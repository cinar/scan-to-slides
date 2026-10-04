#!/usr/bin/env python3
"""Turn a PDF into a .pptx that Google Slides can import.

Two steps:

  scan2slides.py inspect FILE.pdf
      Extracts the scans and writes contact sheets plus strips of each scan's
      top/bottom edge, so the printed page numbers, sideways pages and
      two-page spreads can be identified.

  scan2slides.py build FILE.pdf [--plan FILE.plan.json]
      Reorders, rotates, splits, cleans and resizes the pages according to
      the plan and writes the .pptx files next to the PDF. Without a plan
      the scans are only cleaned and resized, in PDF order.

Needs: Pillow, and pdfimages/pdfinfo/pdftoppm (poppler-utils).
See SKILL.md for the plan format.
"""
import argparse
import json
import statistics
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageMath, ImageOps

JPEG_QUALITY = 82
MARGIN = 20            # white margin around / between pages, px
LONG_SIDE = {1: 1660, 2: 1500}   # page long side in px, by pages per slide
EMU_PER_PX = 6350      # 144 px per inch
BG_WINDOW = 0.12       # paper-colour estimation window, fraction of page width


# --- Extracting scans -------------------------------------------------------

def extract_scans(pdf, tmp):
    """Return one image file per PDF page, in PDF order."""
    info = subprocess.run(["pdfinfo", str(pdf)], check=True, capture_output=True,
                          text=True).stdout
    n_pages = int(next(l.split()[1] for l in info.splitlines() if l.startswith("Pages:")))
    # Scanner PDFs hold exactly one JPEG per page; pull those out untouched.
    subprocess.run(["pdfimages", "-j", str(pdf), str(tmp / "p")], check=True)
    files = sorted(tmp.glob("p-*"))
    if len(files) == n_pages and all(f.suffix == ".jpg" for f in files):
        return files
    # Anything else (several images per page, vector content): render pages.
    for f in files:
        f.unlink()
    subprocess.run(["pdftoppm", "-r", "200", "-jpeg", str(pdf), str(tmp / "r")], check=True)
    return sorted(tmp.glob("r-*.jpg"))


# --- inspect ----------------------------------------------------------------

def _font(size):
    try:
        return ImageFont.load_default(size=size)
    except TypeError:          # Pillow < 10.1
        return ImageFont.load_default()


def _sheet(tiles, cols, cell, path):
    """Paste labelled (label, image) tiles into a grid and save it."""
    cw, ch = cell
    label_h = 30
    rows = -(-len(tiles) // cols)
    sheet = Image.new("RGB", (cols * (cw + 8) + 8, rows * (ch + label_h + 8) + 8), "white")
    draw = ImageDraw.Draw(sheet)
    for i, (label, img) in enumerate(tiles):
        x = 8 + (i % cols) * (cw + 8)
        y = 8 + (i // cols) * (ch + label_h + 8)
        draw.text((x, y), label, fill=(200, 0, 0), font=_font(24))
        draw.rectangle((x - 1, y + label_h - 1, x + cw, y + label_h + ch), outline=(160, 160, 160))
        sheet.paste(img, (x + (cw - img.width) // 2, y + label_h))
    sheet.save(path, "JPEG", quality=85)


def inspect(args):
    pdf = Path(args.pdf)
    out = Path(args.out) if args.out else Path(tempfile.mkdtemp(prefix="scan2slides-"))
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as d:
        files = extract_scans(pdf, Path(d))
        thumbs, tops, bottoms = [], [], []
        print("scan  size        note")
        for i, f in enumerate(files, 1):
            img = Image.open(f).convert("RGB")
            w, h = img.size
            note = "LANDSCAPE - sideways page or two-page spread?" if w > h else ""
            print(f"{i:>4}  {w}x{h:<6} {note}")
            label = f"scan {i}"
            t = img.copy()
            t.thumbnail((380, 540))
            thumbs.append((label, t))
            wide = img.resize((1000, round(1000 * h / w)), Image.LANCZOS)
            band = max(60, round(wide.height * 0.08))
            for strips, box in ((tops, (0, 0, 1000, band)),
                                (bottoms, (0, wide.height - band, 1000, wide.height))):
                s = ImageOps.autocontrast(wide.crop(box), cutoff=1)
                s.thumbnail((640, 110))
                strips.append((label, s))
        for name, tiles, cols, cell, per in (("contact", thumbs, 5, (380, 540), 20),
                                             ("bottom", bottoms, 3, (640, 110), 36),
                                             ("top", tops, 3, (640, 110), 36)):
            for n in range(0, len(tiles), per):
                _sheet(tiles[n:n + per], cols, cell, out / f"{name}-{n // per + 1:02d}.jpg")
    print(f"\n{len(files)} scans. Sheets written to: {out}")
    for f in sorted(out.glob("*.jpg")):
        print(f"  {f.name}")


# --- Page clean-up ----------------------------------------------------------

def _divide(img, bg):
    """Per-channel img / bg, scaled so that paper (img == bg) becomes white."""
    expr = "convert(min(a * 255 / max(b, 1), 255), 'L')"
    ev = getattr(ImageMath, "unsafe_eval", None) or ImageMath.eval
    chans = [ev(expr, a=a, b=b) for a, b in zip(img.split(), bg.split())]
    return Image.merge("RGB", chans)


def clean_page(img, window=BG_WINDOW):
    """Flatten uneven lighting and whiten the paper."""
    w, h = img.size
    small = img.resize((200, max(1, round(200 * h / w))), Image.BILINEAR)
    k = max(3, int(200 * window) | 1)
    # Morphological close removes text/drawings, leaving just the paper.
    bg = small.filter(ImageFilter.MaxFilter(k)).filter(ImageFilter.MinFilter(k))
    bg = bg.filter(ImageFilter.GaussianBlur(k / 3))
    bg = bg.resize((w, h), Image.BICUBIC)
    return _divide(img, bg)


# --- Plan -> pages -> slide images ------------------------------------------

def load_pages(files, plan):
    """Apply the plan to the scans and return {page number: full-size image}."""
    scans = plan["scans"]
    if len(scans) != len(files):
        sys.exit(f"plan lists {len(scans)} scans but the PDF has {len(files)}")
    rotate = {int(k): v for k, v in plan.get("rotate", {}).items()}
    pages = {}
    for i, (entry, f) in enumerate(zip(scans, files), 1):
        if entry is None:                       # skipped scan
            continue
        split = 0.5
        if isinstance(entry, dict):
            entry, split = entry["pages"], entry.get("split", 0.5)
        nums = entry if isinstance(entry, list) else [entry]
        scan = Image.open(f).convert("RGB")
        if i in rotate:
            scan = scan.rotate(-rotate[i], expand=True)
        if len(nums) == 2:
            cut = round(scan.width * split)
            parts = [scan.crop((0, 0, cut, scan.height)),
                     scan.crop((cut, 0, scan.width, scan.height))]
        elif len(nums) == 1:
            parts = [scan]
        else:
            sys.exit(f"scan {i}: expected one page number or a pair, got {nums}")
        for num, img in zip(nums, parts):
            if num in pages:
                sys.exit(f"page {num} appears twice in the plan (second time at scan {i})")
            pages[num] = img
    missing = sorted(set(range(min(pages), max(pages) + 1)) - set(pages))
    if missing:
        print(f"note: no scan for page(s) {missing}", file=sys.stderr)
    return pages


def page_box(pages, per_slide):
    """Box every page is fitted into, shaped like the PDF's typical page."""
    aspect = statistics.median(img.width / img.height for img in pages.values())
    long_side = LONG_SIDE[per_slide]
    if aspect <= 1:
        return round(long_side * aspect), long_side
    return long_side, round(long_side / aspect)


def compose(pages, plan, tmp):
    """Return (slide size, [(title, jpeg path)])."""
    per_slide = plan.get("per_slide", 1)
    label = plan.get("label", "Page")
    do_clean = plan.get("clean", True)
    box_w, box_h = page_box(pages, per_slide)
    size = (per_slide * box_w + (per_slide + 1) * MARGIN, box_h + 2 * MARGIN)

    if per_slide == 1:
        groups = [[n] for n in sorted(pages)]
    else:   # two-page spreads: even page on the left, odd page on the right
        first = min(pages) - min(pages) % 2
        groups = [[n, n + 1] for n in range(first, max(pages) + 1, 2)]
        groups = [g for g in groups if any(n in pages for n in g)]

    slides = []
    for group in groups:
        canvas = Image.new("RGB", size, "white")
        shown = []
        for slot, num in enumerate(group):
            if num not in pages:
                continue
            img = pages[num].copy()
            img.thumbnail((box_w, box_h), Image.LANCZOS)
            if do_clean:
                img = clean_page(img)
            x = MARGIN + slot * (box_w + MARGIN) + (box_w - img.width) // 2
            y = MARGIN + (box_h - img.height) // 2
            canvas.paste(img, (x, y))
            shown.append(str(num))
        path = tmp / f"page-{group[0]:03d}.jpg"
        canvas.save(path, "JPEG", quality=JPEG_QUALITY, optimize=True,
                    progressive=True, subsampling=1)
        slides.append((f"{label} " + "–".join(shown), path))
    return size, slides


# --- PowerPoint (Office Open XML) packaging ---------------------------------

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument.presentationml"
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
EMPTY_TREE = ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
              '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
              '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>{}</p:spTree>')
CLR_MAP = ('bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
           'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" '
           'hlink="hlink" folHlink="folHlink"')
WHITE_BG = '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'


def rels(*items):
    body = "".join(f'<Relationship Id="{i}" Type="{t}" Target="{target}"/>'
                   for i, t, target in items)
    return f'{HEAD}<Relationships xmlns="{PKG_REL}">{body}</Relationships>'


def theme_xml():
    colors = [("dk1", "000000"), ("lt1", "FFFFFF"), ("dk2", "44546A"), ("lt2", "E7E6E6"),
              ("accent1", "4472C4"), ("accent2", "ED7D31"), ("accent3", "A5A5A5"),
              ("accent4", "FFC000"), ("accent5", "5B9BD5"), ("accent6", "70AD47"),
              ("hlink", "0563C1"), ("folHlink", "954F72")]
    clr = "".join(f'<a:{n}><a:srgbClr val="{v}"/></a:{n}>' for n, v in colors)
    font = '<a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/>'
    fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    line = f'<a:ln w="9525">{fill}</a:ln>'
    return (f'{HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Plain">'
            f'<a:themeElements><a:clrScheme name="Plain">{clr}</a:clrScheme>'
            f'<a:fontScheme name="Plain"><a:majorFont>{font}</a:majorFont>'
            f'<a:minorFont>{font}</a:minorFont></a:fontScheme>'
            f'<a:fmtScheme name="Plain"><a:fillStyleLst>{fill * 3}</a:fillStyleLst>'
            f'<a:lnStyleLst>{line * 3}</a:lnStyleLst>'
            f'<a:effectStyleLst>{"<a:effectStyle><a:effectLst/></a:effectStyle>" * 3}</a:effectStyleLst>'
            f'<a:bgFillStyleLst>{fill * 3}</a:bgFillStyleLst></a:fmtScheme>'
            f'</a:themeElements></a:theme>')


def slide_xml(title, w, h, as_background):
    name = escape(title, {'"': "&quot;"})
    blip = '<a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch>'
    if as_background:
        # Slide background: cannot be dragged or deleted by accident.
        bg = f'<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1">{blip}</a:blipFill><a:effectLst/></p:bgPr></p:bg>'
        shapes = ""
    else:
        bg = WHITE_BG
        shapes = (f'<p:pic><p:nvPicPr><p:cNvPr id="2" name="{name}" descr="{name}"/>'
                  f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
                  f'<p:blipFill>{blip}</p:blipFill>'
                  f'<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
                  f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
    return (f'{HEAD}<p:sld {NS}><p:cSld name="{name}">{bg}{EMPTY_TREE.format(shapes)}</p:cSld>'
            f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')


def write_pptx(slides, size, out, as_background):
    w, h = size[0] * EMU_PER_PX, size[1] * EMU_PER_PX
    n = len(slides)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        overrides = "".join(
            f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="{CT}.slide+xml"/>'
            for i in range(1, n + 1))
        z.writestr("[Content_Types].xml", (
            f'{HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="jpg" ContentType="image/jpeg"/>'
            f'<Override PartName="/ppt/presentation.xml" ContentType="{CT}.presentation.main+xml"/>'
            f'<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="{CT}.slideMaster+xml"/>'
            f'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="{CT}.slideLayout+xml"/>'
            '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
            f'{overrides}</Types>'))
        z.writestr("_rels/.rels", rels(("rId1", f"{REL}/officeDocument", "ppt/presentation.xml")))

        ids = "".join(f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, n + 1))
        z.writestr("ppt/presentation.xml", (
            f'{HEAD}<p:presentation {NS}>'
            '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
            f'<p:sldIdLst>{ids}</p:sldIdLst>'
            f'<p:sldSz cx="{w}" cy="{h}"/><p:notesSz cx="{h}" cy="{w}"/></p:presentation>'))
        z.writestr("ppt/_rels/presentation.xml.rels", rels(
            ("rId1", f"{REL}/slideMaster", "slideMasters/slideMaster1.xml"),
            *((f"rId{i + 1}", f"{REL}/slide", f"slides/slide{i}.xml") for i in range(1, n + 1)),
            (f"rId{n + 2}", f"{REL}/theme", "theme/theme1.xml")))

        z.writestr("ppt/theme/theme1.xml", theme_xml())
        z.writestr("ppt/slideMasters/slideMaster1.xml", (
            f'{HEAD}<p:sldMaster {NS}><p:cSld>{WHITE_BG}{EMPTY_TREE.format("")}</p:cSld>'
            f'<p:clrMap {CLR_MAP}/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
            '</p:sldMaster>'))
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", rels(
            ("rId1", f"{REL}/slideLayout", "../slideLayouts/slideLayout1.xml"),
            ("rId2", f"{REL}/theme", "../theme/theme1.xml")))
        z.writestr("ppt/slideLayouts/slideLayout1.xml", (
            f'{HEAD}<p:sldLayout {NS} type="blank" preserve="1"><p:cSld name="Blank">'
            f'{EMPTY_TREE.format("")}</p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'))
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", rels(
            ("rId1", f"{REL}/slideMaster", "../slideMasters/slideMaster1.xml")))

        for i, (title, img) in enumerate(slides, 1):
            z.writestr(f"ppt/slides/slide{i}.xml", slide_xml(title, w, h, as_background))
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", rels(
                ("rId1", f"{REL}/slideLayout", "../slideLayouts/slideLayout1.xml"),
                ("rId2", f"{REL}/image", f"../media/image{i}.jpg")))
            # JPEGs are already compressed; store them as-is.
            z.write(img, f"ppt/media/image{i}.jpg", zipfile.ZIP_STORED)


# --- build ------------------------------------------------------------------

def build(args):
    pdf = Path(args.pdf)
    if not pdf.is_file():
        sys.exit(f"no such PDF: {pdf}")
    plan_path = Path(args.plan) if args.plan else pdf.with_suffix(".plan.json")
    if plan_path.is_file():
        plan = json.loads(plan_path.read_text())
    elif args.plan:
        sys.exit(f"no such plan: {plan_path}")
    else:
        plan = {}       # no plan: keep the scans as they are, in PDF order
    if args.per_slide:
        plan["per_slide"] = args.per_slide
    if plan.get("per_slide", 1) not in LONG_SIDE:
        sys.exit("per_slide must be 1 or 2")
    out_dir = Path(args.out_dir) if args.out_dir else pdf.parent
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        files = extract_scans(pdf, tmp)
        if "scans" not in plan:
            plan["scans"] = list(range(1, len(files) + 1))
            print(f"note: no plan at {plan_path}; using the scans in PDF order, numbered "
                  f"1-{len(files)}, without reordering, rotating or splitting any. "
                  "Run 'inspect' and write a plan to fix those (see SKILL.md).",
                  file=sys.stderr)
        size, slides = compose(load_pages(files, plan), plan, tmp)
        keep = out_dir / "slides" / pdf.stem
        keep.mkdir(parents=True, exist_ok=True)
        for old in keep.glob("page-*.jpg"):
            old.unlink()
        for _, img in slides:
            (keep / img.name).write_bytes(img.read_bytes())
        for suffix, as_bg in (("", True), (" (movable images)", False)):
            out = out_dir / f"{pdf.stem}{suffix}.pptx"
            write_pptx(slides, size, out, as_bg)
            print(f"{out.name}: {len(slides)} slides, {out.stat().st_size / 1e6:.1f} MB")
        print(f"slide images: {keep}")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("inspect", help="write contact sheets and page-edge strips")
    a.add_argument("pdf")
    a.add_argument("--out", help="directory for the sheets (default: a new temp dir)")
    a.set_defaults(func=inspect)
    b = sub.add_parser("build", help="build the .pptx files from a plan")
    b.add_argument("pdf")
    b.add_argument("--plan", help="plan JSON (default: <pdf name>.plan.json next to the PDF; "
                                  "without one the scans are used as they are, in PDF order)")
    b.add_argument("--per-slide", type=int, choices=sorted(LONG_SIDE),
                   help="pages per slide (overrides the plan)")
    b.add_argument("--out-dir", help="output directory (default: next to the PDF)")
    b.set_defaults(func=build)
    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
