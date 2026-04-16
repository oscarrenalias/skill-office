# skill-office — Microsoft Office skills for AI agents

`skill-office` is a collection of AI agent skills for working with Microsoft Office files.
Each skill exposes file operations as a CLI with structured JSON output, designed to be called
by agents running in Claude Code or similar environments. Skills can also be used directly
from the terminal.

## Skills included

| Skill | File types | Entry point |
|---|---|---|
| `pypptx` | `.pptx` — create, edit, inspect PowerPoint presentations | `python3 pypptx.py` |
| `pyxlsx` | `.xlsx` — read sheets and tables, update cells, manage sheets | `python3 pyxlsx.py` |

Each skill self-bootstraps its own `.venv` on first run — no manual setup required.

---

## Installing the skills

### Via APM

```bash
apm install oscarrenalias/skill-office#vX.Y.Z
```

Please check the latest version of the skill in the [releases page](https://github.com/oscarrenalias/skill-office/releases).

### Via zip (manual)

Download the latest `skills-vX.Y.Z.zip` from the
[releases page](https://github.com/oscarrenalias/skill-office/releases) and
unzip it into your skills directory if you intend the skills to be available globally in your local environment:

```bash
# Claude Code
unzip skills-vX.Y.Z.zip -d ~/.claude/skills/

# Codex / other agents
unzip skills-vX.Y.Z.zip -d ~/.agents/skills/
```

Skills can also be deployed at the project level:

```bash
# Claude Code
unzip skills-vX.Y.Z.zip -d .claude/skills/

# Codex / other agents
unzip skills-vX.Y.Z.zip -d .agents/skills/
```

The zip contains both skills. Each one bootstraps its own `.venv` with the required
Python dependencies on first run.

---

## Pre-requisites

The skills bootstrap their own Python dependencies on first run (first run will be slower than usual).

The `pypptx thumbnails` command additionally requires two system tools:

| Tool | macOS | Debian/Ubuntu |
|---|---|---|
| LibreOffice (`soffice`) | `brew install --cask libreoffice` | `sudo apt-get install libreoffice` |
| Poppler (`pdftoppm`) | `brew install poppler` | `sudo apt-get install poppler-utils` |

All other commands work without these tools.

---

## Agent scenarios

### Scenario 1 — Create a presentation from scratch

Ask Claude Code to build a deck from a template or from nothing:

```
Create a 6-slide project status presentation using the template in @template.pptx.
Slide 1 should be the cover slide with the project name and today's date.
Slides 2–5 should cover: overview, current status, risks, and next steps.
Slide 6 is a closing/Q&A slide.

After generating the file, check the deck for quality issues, then generate thumbnails
so you can visually confirm the output looks correct.
```

Claude will typically:
1. Run `pypptx slide layouts template.pptx` to discover available layout names
2. Write a script using python-pptx and the skill's `.venv`
3. Run `pypptx verify output.pptx` to catch structural issues
4. Run `pypptx thumbnails output.pptx` and read the image to visually confirm the result

### Scenario 2 — Modify an existing presentation

Ask Claude Code to make targeted edits to an existing presentation:

```
I have a deck in @quarterly_review.pptx. Please:
- Move the "Risks" slide (currently slide 5) to be slide 3
- Delete the blank slide at position 7
- Extract all the text so I can review what's there

After making changes, run verify and generate thumbnails so we can confirm everything looks right.
```

Claude will typically:
1. Run `pypptx slide list quarterly_review.pptx` to see the current structure
2. Run `pypptx extract-text quarterly_review.pptx` to read the content
3. Run `pypptx slide move` and `pypptx slide delete` to make the structural changes
4. Run `pypptx verify` and `pypptx thumbnails` for validation and visual QA

### Scenario 3 — Extract a workplan from Excel and build a roadmap in PowerPoint

The two skills work together for cross-format workflows:

```
Read the workplan in @workplan.xlsx, extract all tasks from the "Q2 Plan" sheet,
and create a roadmap slide in @template.pptx — one row per workstream,
with task names, owners, and due dates laid out as a timeline table.
```

Claude will typically:
1. Run `pyxlsx sheet list workplan.xlsx` to discover sheet names
2. Run `pyxlsx table read workplan.xlsx "Q2 Plan"` to extract structured task data as JSON
3. Write a python-pptx script to generate the roadmap slide from that data
4. Run `pypptx verify` and `pypptx thumbnails` for QA

### Scenario 4 — Read and update an Excel table

```
Open @budget.xlsx, find the "Headcount" sheet, and update the Q3 budget
for the "Engineering" row to 450000. Then show me the full updated table.
```

Claude will typically:
1. Run `pyxlsx table read budget.xlsx Headcount` to read current data
2. Run `pyxlsx cell set budget.xlsx Headcount B5 450000` to update the cell
3. Run `pyxlsx table read` again to confirm the change

---

## Human usage

The CLIs can be called directly from the terminal but that is not their intended usage pattern.

---

## Output contract

Both skills follow the same output contract:

- **Default output**: every command writes a single JSON object to stdout.
- **`--plain` flag**: pass `--plain` to receive human-readable text instead.
- **Errors**: all error messages are written to stderr, never stdout.
- **Exit codes**: `0` on success, `1` on any error.

---

## pypptx commands

Run `pypptx --help` or `pypptx <command> --help` for the full option reference.
See [SKILL.md](.apm/skills/pypptx/SKILL.md) for detailed usage guidance.

| Command | Description |
|---|---|
| `verify <file>` | Quality checks: unfilled placeholders, font sizes, overflow, clipping, overlaps |
| `extract-text <file>` | Extract all text from slides |
| `thumbnails <file>` | Generate labeled thumbnail grid (requires LibreOffice + Poppler) |
| `slide list <file>` | List slides in presentation order |
| `slide layouts <file>` | List available slide layouts with indices |
| `slide add <file>` | Add a slide by layout or by duplicating an existing one |
| `slide delete <file> <n>` | Delete slide at 1-based index |
| `slide move <file> <from> <to>` | Move a slide between positions |
| `unpack <file> [dir]` | Extract `.pptx` ZIP to a directory |
| `clean <file_or_dir>` | Remove orphaned XML parts |
| `pack <dir> <file>` | Repack a directory into a `.pptx` |

---

## pyxlsx commands

Run `pyxlsx --help` or `pyxlsx <command> --help` for the full option reference.
See [SKILL.md](.apm/skills/pyxlsx/SKILL.md) for detailed usage guidance.

| Command | Description |
|---|---|
| `info <file>` | Workbook metadata: sheet names and named ranges |
| `sheet list <file>` | List sheets with row/column counts and visibility |
| `sheet read <file> <sheet>` | Read sheet as a raw 2D grid |
| `sheet add/delete/rename <file>` | Add, delete, or rename a sheet |
| `table read <file> <sheet>` | Read sheet as array-of-objects keyed by header row |
| `cell get <file> <sheet> <cell>` | Get a single cell value |
| `cell set <file> <sheet> <cell> <value>` | Set a single cell value or formula |
| `unpack <file> [dir]` | Extract `.xlsx` ZIP to a directory |
| `pack <dir> <file>` | Repack a directory into an `.xlsx` |
