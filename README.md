<div align="center">

# scan-to-slides

**Turn a scanned PDF into a clean, annotatable Google Slides deck.**

A [Claude Code](https://claude.com/claude-code) skill that converts a PDF into a
Google Slides-ready `.pptx`, one page per slide.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3](https://img.shields.io/badge/Python-3-3776AB.svg?logo=python&logoColor=white)](https://www.python.org/)
[![Claude Code skill](https://img.shields.io/badge/Claude%20Code-skill-D97757.svg)](https://claude.com/claude-code)

</div>

---

Phone-scanner PDFs are awkward to work on together: pages come out of order,
some are sideways, some hold two pages, and the paper is grey and shadowed.
`scan-to-slides` fixes all of that and puts each page on its own slide, so a
group can write on the pages in Google Slides.

## Features

- **Printed page order.** Claude reads the page numbers off the scans and puts
  the pages back in sequence.
- **Rotation.** Sideways pages are turned upright.
- **Spread splitting.** A scan that holds two pages is cut into two.
- **Paper whitening.** Shadows and colour casts are removed while
  illustrations are left intact.
- **Small files.** Pages are downsized so the deck stays well below the size
  of the original PDF.
- **Made for annotating.** Each page is the slide background, so it can't be
  dragged or deleted while people add text boxes on top.
- **No embedded metadata.** The `.pptx` is written directly as Office Open
  XML, with nothing about the author or machine inside.

Ordinary (non-scanned) PDFs work too; their pages are rendered at 200 dpi.

## Requirements

- [Claude Code](https://claude.com/claude-code)
- Python 3 with [Pillow](https://python-pillow.org/)
- [poppler-utils](https://poppler.freedesktop.org/) (`pdfinfo`, `pdfimages`,
  `pdftoppm`)

```bash
# Debian / Ubuntu
sudo apt install poppler-utils python3-pil

# macOS
brew install poppler && pip3 install Pillow
```

## Installation

Clone the repository into your Claude Code skills directory:

```bash
git clone https://github.com/cinar/scan-to-slides.git ~/.claude/skills/scan-to-slides
```

## Usage

In Claude Code, point it at a PDF and ask for slides:

> Convert `pages.pdf` to Google Slides

Claude inspects the scans, works out the page order, rotations and splits,
builds the deck and checks the result. You get, next to the PDF:

| Output | What it is |
| --- | --- |
| `pages.pptx` | The deck, with each page as a slide background |
| `pages (movable images).pptx` | The same slides with each page as an ordinary picture, in case an importer drops slide backgrounds |
| `slides/pages/page-NNN.jpg` | The processed page images |

To open the deck in Google Slides, upload the `.pptx` to Google Drive, open it
with Google Slides, then choose **File > Save as Google Slides**.

### Running the script by hand

The skill is a thin layer over one script, which also works on its own:

```bash
python3 scripts/scan2slides.py inspect pages.pdf   # contact sheets to read page numbers from
python3 scripts/scan2slides.py build pages.pdf     # uses pages.plan.json next to the PDF
```

Without a plan, `build` whitens and downsizes the scans but keeps them as they
are, in PDF order. Reordering, rotating and splitting need a plan, which is
the part Claude writes after looking at the scans:

```json
{
  "per_slide": 1,
  "label": "Page",
  "scans": [9, 8, 10, 11, {"pages": [12, 13], "split": 0.5}, null, 14],
  "rotate": {"7": 90}
}
```

See [SKILL.md](SKILL.md) for the full workflow and the plan format.

## How it works

1. **Inspect.** The script extracts the scans and writes contact sheets, plus
   contrast-boosted strips of each scan's top and bottom edge, where page
   numbers usually sit.
2. **Plan.** Claude reads those images and records each scan's printed page
   number, rotation and split in a `.plan.json` file next to the PDF.
3. **Build.** The script reorders, rotates, splits, whitens and resizes the
   pages according to the plan, then writes the `.pptx` files.
4. **Check.** Claude looks over the rotated and split pages before reporting
   back.

## Contributing

Issues and pull requests are welcome. Please keep source material out of the
repository: no PDFs, page images, plan files or text quoted from a document,
not even as an example or a test fixture.

## License

[MIT](LICENSE) © Onur Cinar
