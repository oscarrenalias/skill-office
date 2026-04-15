---
name: pypptx
description: Python CLI for reading, editing, and creating PowerPoint .pptx files
author: Renalias, Oscar
tags:
  - powerpoint
  - pptx
  - presentations
  - office
entry_point: pypptx.py
requires:
  optional:
    - LibreOffice (soffice) — needed for thumbnails command
    - Poppler (pdftoppm) — needed for thumbnails command
    - Pillow — needed for thumbnails command (installed automatically via pip)
---

# pypptx

A Python CLI for reading, editing, and creating PowerPoint `.pptx` files.

## Quick reference

| Task | How |
|---|---|
| Read slide text | `python3 pypptx.py extract-text <file>` |
| List slides | `python3 pypptx.py slide list <file>` |
| Add / delete / move slides | `python3 pypptx.py slide add/delete/move <file> ...` |
| Visual overview | `python3 pypptx.py thumbnails <file>` |
| Structural edit (XML) | unpack → edit → clean → pack |
| Create from scratch | Write a python-pptx script, run it via `.venv/bin/python3` |

The entry point is `python3 pypptx.py` at the repo root. It self-bootstraps a
`.venv` on first run with no external tooling required.

---

## Reading content

Extract all text from a presentation:

```bash
python3 pypptx.py extract-text presentation.pptx
```

Limit to specific slides with `--slides 1,3`. Output goes to stdout (no JSON
wrapper) unless `--output <file>` is given, in which case command metadata is
emitted as JSON.

---

## Visual inspection

Generate a labeled thumbnail grid to see slide layout at a glance:

```bash
python3 pypptx.py thumbnails presentation.pptx
```

Requires LibreOffice and Poppler — see README for installation. Use this to
verify edits look correct before delivering or committing a file. Hidden slides
appear as a hatched grey placeholder so grid index always matches slide number.

---

## Editing an existing presentation

Use the unpack → edit → clean → pack workflow for structural changes:

```bash
python3 pypptx.py unpack presentation.pptx      # expand to directory
# manipulate slides, edit XML, or run a python-pptx script against the directory
python3 pypptx.py clean presentation/           # remove orphans
python3 pypptx.py pack presentation/ output.pptx
```

Most slide commands also accept a `.pptx` file directly and handle
unpack/clean/repack internally:

```bash
python3 pypptx.py slide add presentation.pptx --duplicate 2
python3 pypptx.py slide delete presentation.pptx 3
python3 pypptx.py slide move presentation.pptx 3 1
```

Use `slide list` to confirm slide order and `slide layouts` to find layout
indices for `slide add --layout`.

---

## Creating a presentation from scratch

Write a python-pptx script, then run it using the skill's virtual environment.
If this is the first time running the skill, bootstrap the venv first:

```bash
python3 pypptx.py --version   # triggers first-run bootstrap
```

Then write your script and run it:

```bash
.venv/bin/python3 create_deck.py
```

Example `create_deck.py`:

```python
from pptx import Presentation
from pptx.util import Inches, Pt

# If a template file is provided, pass it here — it loads the slide master,
# layouts, theme, and color palette. If no template is provided, omit the
# argument and python-pptx uses its built-in blank default.
prs = Presentation('template.pptx')  # or: Presentation()

slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
slide.shapes.title.text = "Introduction"
slide.placeholders[1].text = "Key points go here"
prs.save("output.pptx")
```

Use `python3 pypptx.py slide layouts <file>` on the template (or any existing
deck) to find the layout indices available in that theme.

### Choosing a layout

**Always run `slide layouts` before choosing a layout index.** Never hardcode
an index — layout assignments vary between templates.

```bash
python3 pypptx.py slide layouts presentation.pptx
```

> **Warning:** Layout index 1 (entry 1) is almost always the cover slide
> ("Title Slide"). Never use it for regular content slides.

Common layout names and when to use them:

| Name | Use for |
|---|---|
| Title Slide | Cover slide only (first slide) |
| Title and Content | Standard body slides |
| Section Header | Section dividers |
| Two Content | Side-by-side content |
| Title Only | Slides with custom content below the title |
| Blank | Fully custom slides (no placeholders) |

### Use placeholders, not `add_textbox()`

Before writing content, inspect what placeholders the chosen layout provides:

```python
for ph in slide.placeholders:
    print(ph.placeholder_format.idx, ph.name)
```

Write into them by index:

```python
slide.placeholders[0].text = "My Title"
slide.placeholders[1].text = "Body text"
```

**Never use `add_textbox()`** on a slide that has a layout with placeholders.
It places an unstyled text box on top of the master design — wrong font, wrong
colour, wrong position, often unreadable against the background.
`add_textbox()` is only appropriate for fully blank slides (layout "Blank")
where no placeholders exist.

### Design guidance

**Colors** — use the theme palette from an existing file where possible.
Extract it with `python3 pypptx.py unpack` and inspect
`ppt/theme/theme1.xml`. Avoid free-form hex values unrelated to the deck's
palette.

**Typography** — respect the slide master's font stack. Prefer placeholder
text frames over adding raw text boxes; placeholders inherit master styles.

**Spacing** — use `Inches()` and `Pt()` from `pptx.util` for all measurements.
Leave slide margins of at least 0.3–0.5 in on all edges. Never place shapes
flush to the slide boundary.

---

### Best practices and common mistakes

#### Text inside shapes

**Do** write text directly into a shape's text frame and control margins there:

```python
tf = shape.text_frame
tf.word_wrap = True
tf.margin_left   = Inches(0.1)
tf.margin_right  = Inches(0.1)
tf.margin_top    = Inches(0.05)
tf.margin_bottom = Inches(0.05)
tf.text = "My label"
```

**Don't** add a separate `add_textbox()` on top of an existing shape to place
text — it creates an invisible stacked box with wrong font, wrong colour, and
no connection to the shape below.

#### Placeholders vs. free shapes

**Do** use `slide.placeholders[idx]` for title and body content — they inherit
the master font, colour, and position. Inspect what a layout offers before
writing:

```python
for ph in slide.placeholders:
    print(ph.placeholder_format.idx, ph.name)
```

**Don't** delete placeholders and replace them with textboxes. `add_textbox()`
is only appropriate on fully blank slides (layout "Blank") where no
placeholders exist.

#### Slide dimensions

**Do** read actual slide dimensions before placing any shape:

```python
W = prs.slide_width   # e.g. 9144000 EMU for a 10-inch-wide slide
H = prs.slide_height
```

**Don't** hardcode `Inches(10)` / `Inches(7.5)` — template slides vary.
Always derive positions and sizes from `prs.slide_width` / `prs.slide_height`.

#### Tables

**Do** use `shapes.add_table(rows, cols, left, top, width, height)` for
tabular data, then set column widths explicitly so they sum to the table width:

```python
tbl = slide.shapes.add_table(4, 3, left, top, width, height).table
tbl.columns[0].width = Inches(3)
tbl.columns[1].width = Inches(2)
tbl.columns[2].width = Inches(2)
```

**Don't** simulate tables with individually positioned text boxes — alignment
breaks as soon as content changes length.

#### Colors

**Do** extract and reuse theme colors from the deck's `ppt/theme/theme1.xml`.
**Don't** invent hex values not present in the palette — they clash with the
master and look unprofessional.

#### Layout index

**Do** always run `slide layouts` before choosing a layout index — assignments
vary between templates:

```bash
python3 pypptx.py slide layouts presentation.pptx
```

**Don't** assume index 0 = blank, index 1 = "Title and Content". Index 1 is
almost always the cover ("Title Slide") in branded templates.

#### After editing

**Do** always run `verify` and `thumbnails` before declaring a file done:

```bash
python3 pypptx.py verify output.pptx
python3 pypptx.py thumbnails output.pptx
```

Read the generated thumbnail image to catch layout issues invisible in code:
overlapping shapes, text clipping, white-on-white text, wrong layout on slide 1.

**Don't** skip the visual check — `verify` catches structural issues but cannot
see rendering artefacts.

---

## Output contract

All commands write a single JSON object to stdout by default.
Pass `--plain` for human-readable text. Errors go to stderr. Exit code 0 on
success, 1 on any error.

---

## QA checklist

Before delivering or committing a modified or newly created `.pptx`:

1. **Structural validation** — run the built-in quality checks to catch unfilled
   placeholders, font size issues, shape overflow, text clipping, and significant
   shape overlap:
   ```bash
   python3 pypptx.py verify output.pptx
   ```
   Exit code 0 means all checks passed. Non-zero exit means errors were found;
   review the `errors` list in the JSON output (or use `--plain` for one line per
   finding).

2. **Re-open with python-pptx** — must not raise:
   ```bash
   .venv/bin/python3 -c "from pptx import Presentation; Presentation('output.pptx')"
   ```
3. **Slide count** — `python3 pypptx.py slide list output.pptx` matches expectation.
4. **Text check** — `python3 pypptx.py extract-text output.pptx` to verify content landed in the right slides.
5. **No orphans** — `python3 pypptx.py clean output.pptx` returns `{"removed": []}`.
6. **Visual check** — generate a thumbnail grid and read the image file to
   verify the output looks correct:
   ```bash
   python3 pypptx.py thumbnails output.pptx
   ```
   Read the generated image (e.g. `thumbnails.jpg`) to view it. Check for:
   - Text is legible — not white-on-white, not obscured by background elements
   - Cover layout ("Title Slide") used only on slide 1
   - No raw text boxes floating over branded backgrounds
   - Content sits within slide boundaries, not clipped or overflowing
   - Each slide uses a layout appropriate to its content type

   Skip only if LibreOffice or Poppler is not available in the environment.

---

## Dependencies

The skill self-bootstraps on first run. The following are installed automatically
into `.venv/`:

- `python-pptx` — core presentation manipulation
- `click` — CLI framework
- `defusedxml` — safe XML parsing

Optional (for `thumbnails`): `Pillow`, LibreOffice (`soffice`), Poppler (`pdftoppm`).
See README for system installation instructions.
