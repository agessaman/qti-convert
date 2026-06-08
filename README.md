# qti-convert

Converts QTI quiz exports from Canvas into JSON or Word (docx) for printing, including randomized test forms with question-pool selection.

Canvas exports must be unzipped first. This tool has been tested with Canvas export packages only.

## Setup

```bash
pip install pipenv
pipenv install
```

Or install dependencies directly:

```bash
pip install logzero lxml python-docx htmldocx
```

Run commands from the `src/` directory:

```bash
cd src
python3 main.py --help
```

## Basic usage

Convert a quiz to JSON (stdout):

```bash
python3 main.py "/path/to/export/imsmanifest.xml"
```

Convert to JSON file:

```bash
python3 main.py -o quiz.json "/path/to/export/imsmanifest.xml"
```

Convert to Word:

```bash
python3 main.py -f docx -o quiz.docx "/path/to/export/imsmanifest.xml"
```

## Command-line options

| Flag | Default | Description |
|------|---------|-------------|
| `input` | — | Path to `imsmanifest.xml` in the unzipped Canvas export |
| `-f`, `--format` | `json` | Output format: `json` or `docx` |
| `-o`, `--output` | — | Output file path |
| `--forms N` | `1` | Number of randomized forms to generate (Form A, B, C, …) |
| `--shuffle-scope` | `group` | `group` = shuffle within question pools; `test` = shuffle all questions |
| `--seed INT` | — | RNG seed for reproducible forms |
| `--form-output` | `separate` | `separate` = one file per form; `combined` = single file with all forms |
| `--strip-html` | off | Remove HTML tags from question and answer text |
| `--split-matching` | off | Split matching questions into batches of ≤5 for scantron (A–E) |
| `-v` | — | Increase log verbosity |
| `-h`, `--help` | — | Show help |

## Examples

### Chapter 30 Test — printable Word forms

Generate three randomized forms as separate docx files with clean text and scantron-friendly matching:

```bash
python3 main.py "/Users/adam/Downloads/Chapter 30 Test/imsmanifest.xml" \
  --forms 3 \
  -f docx \
  --strip-html \
  --split-matching \
  -o Chapter30.docx
```

Output:

- `Chapter30-Form-A.docx`
- `Chapter30-Form-B.docx`
- `Chapter30-Form-C.docx`

Each form has 16 questions: 1 matching block (split into two 5-item tables) + 10 multiple choice + 5 multiple choice.

### Comprehensive Test — many chapter pools

Generate three forms with questions shuffled across the entire test:

```bash
python3 main.py "/Users/adam/Downloads/Comprehensive Test World History/imsmanifest.xml" \
  --forms 3 \
  --shuffle-scope test \
  --seed 42 \
  -f docx \
  --strip-html \
  -o Comprehensive.docx
```

Each form has 54 questions drawn from 15 chapter question pools.

### Combined output (all forms in one file)

```bash
python3 main.py "/Users/adam/Downloads/Chapter 30 Test/imsmanifest.xml" \
  --forms 3 \
  -f docx \
  --strip-html \
  --form-output combined \
  -o Chapter30-all-forms.docx
```

### Reproducible forms with JSON output

```bash
python3 main.py "/Users/adam/Downloads/Comprehensive Test World History/imsmanifest.xml" \
  --forms 3 \
  --seed 42 \
  -o comprehensive.json
```

Produces `form-a.json`, `form-b.json`, and `form-c.json`.

### Legacy export (all questions, no randomization)

```bash
python3 main.py "/Users/adam/Downloads/Chapter 30 Test/imsmanifest.xml" -o chapter30-full.json
```

Exports every question in the source file (20 for Chapter 30, 82 for Comprehensive Test) in original order.

## Output formatting

**Word documents** apply these conventions automatically:

- **Matching** — labeled once per matching block; prompts numbered sequentially; word bank lettered A, B, C, …
- **Multiple Choice** — labeled once per question-pool section; stems numbered sequentially; options indented 0.25″ and lettered A, B, C, D, …
- Questions flow continuously without a page break between each one

**`--split-matching`** divides matching questions into even batches of at most 5 (e.g. 10 items → 5 + 5), with each batch using only A–E in the word bank for scantron bubble sheets.

**`--strip-html`** is recommended for Canvas exports that embed inline HTML styling in question text.

## Form generation behavior

| Mode | When | Pool selection | Ordering |
|------|------|----------------|----------|
| Legacy | `--forms 1` | All questions exported | Original document order |
| Group shuffle | `--forms N` (default) | `selection_number` per Canvas question pool | Shuffle within each pool; pool order preserved |
| Test shuffle | `--forms N --shuffle-scope test` | Same pool selection | Shuffle entire question list |

Form labels: Form A, Form B, … Form Z, Form AA, …

## Bundled test case

A small sample export is included in the repo:

```bash
python3 main.py ../cases/imsmanifest.xml -f docx -o sample.docx
```
