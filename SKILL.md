---
name: scan-to-slides
description: Convert a PDF into a Google Slides-ready .pptx, one page per slide. Handles scanned PDFs (phone-scanner output, one photo per page) by putting the pages back in printed order, rotating sideways pages, splitting two-page scans, whitening the paper and downsizing the images, so people can annotate the pages together. Use this whenever the user has a PDF and wants it in Google Slides or PowerPoint, wants its pages as slides to share or write on, or says something like "do the same for this PDF" about a new file, even if they don't say "slides" explicitly.
---

# Scan to Slides

Turns a PDF into a `.pptx` that Google Slides imports, with each
page as a slide background so it can be annotated but not accidentally moved.

The bundled script does all the mechanical work. Your part is the judgement a
script can't do: looking at the scans to work out the printed page order, which
scans are sideways, and which hold two pages.

Script: `scripts/scan2slides.py` (needs Pillow and poppler-utils: `pdfinfo`,
`pdfimages`, `pdftoppm`).

## Keep source material out of this skill

This skill lives in a public git repository; the PDFs it processes may be
copyrighted or private. Nothing from a PDF belongs in the skill directory:
no titles, no plan files, no page images, no quoted page text, not even as an
example or a test fixture.

- Write inspection sheets to a temp or scratch directory (the default).
- Write the plan and all outputs next to the user's PDF (the default).
- If you improve the skill or script after a tricky PDF, describe the case
  generically ("a scan holding two pages"), never by naming the document.

## Workflow

### 1. Inspect

```bash
python3 scripts/scan2slides.py inspect "/path/to/File.pdf"
```

This prints each scan's pixel size (flagging landscape ones) and writes, to a
temp directory it reports:

- `contact-NN.jpg`: thumbnails of every scan, 20 per sheet
- `bottom-NN.jpg` and `top-NN.jpg`: the bottom and top edge of every scan,
  contrast-boosted, where printed page numbers usually sit

Every tile is labelled `scan N`, the scan's position in the PDF (1-based).
Read the images with the Read tool.

### 2. Work out the plan

For every scan decide:

- **Its printed page number.** Read it from the bottom strips (or top strips if
  the pages are numbered at the top). Phone scans are very often out of order
  in adjacent pairs (9, 8, 10, 11, 13, 12 ...), so read every number rather
  than assuming a pattern from the first few.
- **Whether it is sideways.** A landscape scan among portrait pages is either
  sideways or a spread. Open that scan's thumbnail to tell which. For a
  sideways page, work out the clockwise rotation that makes it upright from
  something with an obvious up, like an illustration or the page number.
- **Whether it holds two pages.** A spread shows two page numbers and a gutter
  down the middle. Estimate where the gutter is as a fraction of the width.
- **Whether to drop it.** Blank pages or accidental duplicates.

When a page number is cropped off or unreadable, infer it from its neighbours
and the gap in the sequence, and tell the user which pages you inferred. For
pages that have no number at all (covers, title pages), give them numbers that
sort where they belong, such as 0 for a front cover.

The edge strips of a landscape scan usually miss the page number; look at the
contact sheet or the scan itself for those.

### 3. Write the plan

Save it next to the PDF as `<PDF name>.plan.json`:

```json
{
  "per_slide": 1,
  "label": "Page",
  "scans": [9, 8, 10, 11, {"pages": [12, 13], "split": 0.5}, null, 14],
  "rotate": {"7": 90}
}
```

- `scans`: one entry per scan, in PDF order. A number is that scan's printed
  page number. `{"pages": [a, b], "split": f}` is a two-page scan cut at
  fraction `f` of its width (`[a, b]` alone means 0.5). `null` drops the scan.
- `rotate`: scan position (not page number) to degrees clockwise. Rotation
  happens before a spread is split.
- `per_slide`: 1 for one page per slide (the default, and what works best for
  working on a page together), or 2 for two-page spreads with the even page on
  the left. Use 2 only if the user asks for two pages per slide.
- `label`: word used in slide titles, e.g. "Page 12". Match the document's
  language if you like ("Seite", "Página").
- `clean`: set to `false` to skip paper whitening. Leave it on unless pages
  are mostly full-bleed photographs, where whitening can wash out the image.

### 4. Build

```bash
python3 scripts/scan2slides.py build "/path/to/File.pdf"
```

This writes, next to the PDF:

- `File.pptx`: each page is the slide background, so it can't be dragged or
  deleted while people add text boxes on top
- `File (movable images).pptx`: same slides with the page as an ordinary
  picture, a fallback in case an importer drops slide backgrounds
- `slides/File/page-NNN.jpg`: the processed page images

If there is no plan file it says so and builds from the scans as they are, in
PDF order; that is a fallback for people running the script by hand, not a
substitute for steps 1-3.

It stops if the plan's scan count doesn't match the PDF or a page number
appears twice, and prints a note listing any page numbers missing from the
sequence. Treat that note as a prompt to recheck your reading of the numbers
before assuming the page really wasn't scanned.

### 5. Check the result

Look at a handful of the output images before reporting: every page you
rotated, both halves of every split, the first and last page, and one page
that looked shadowed or tinted in the contact sheet. Rotations are easy to get
backwards and a split in the wrong place cuts text off, and both are obvious
at a glance. Fix the plan and rebuild if needed; building takes well under a
minute per PDF.

### 6. Report

Tell the user:

- slide count, page range and file size compared with the original PDF
- what you changed: which pages were reordered, rotated, split or dropped, and
  any page numbers you inferred
- how to open it: upload the `.pptx` to Google Drive, open with Google Slides,
  then File > Save as Google Slides
- that the import into Google Slides itself was not tested, unless you
  actually did it

## What the script does to each page

Knowing this helps when a result looks off.

- **Extraction.** Scanner PDFs hold one JPEG per page; those are pulled out
  untouched. Other PDFs are rendered at 200 dpi instead.
- **Sizing.** Every page is fitted into a box shaped like the PDF's typical
  page, 1660 px on the long side for one page per slide and 1500 px for two.
  That leaves room to zoom on an ordinary screen while keeping the file well
  below the original PDF's size. The slide is sized to fit the page exactly
  rather than 16:9, so the page is as large as possible.
- **Whitening.** The paper colour is estimated locally and divided out, which
  removes shadows and colour casts and turns the paper white while leaving
  illustrations intact. A page lit very unevenly can keep a faint tint.
- **Packaging.** The `.pptx` is written directly as Office Open XML with a
  blank theme and no document metadata, so nothing about the author or machine
  is embedded.
