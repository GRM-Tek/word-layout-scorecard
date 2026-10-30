# The scorecard

One command measures an editor against Microsoft Word on a set of documents and prints the spec table with that editor's column filled:

```bash
scorecard/score wordcraft scorecard/public-set
```

`score <editor> <set-folder>`. With `wordcraft` the pinned engine is run in the sealed container to make the
outputs first; any other editor name scores outputs already on disk (`--outputs DIR`). The table goes to the
screen and everything, per document, to `score.json` in the run folder (`<set>/out/<editor>/`). Counts only:
no document text is printed or written outside that folder.

Options: `--outputs DIR` (score these outputs instead of making them), `--out DIR` (run folder), `--fonts DIR;DIR`
(font folders mounted read-only into the container; default: this computer's Windows fonts and the Office cloud
fonts, or `/usr/share/fonts` on Linux), `--proto-rescue DIR` (a proto-rescue checkout, for the "nothing lost on
save" measure, which reads both files through its Word import; also `$PROTO_RESCUE`), `--no-import`.

## The four measures, and where they come from

| Row of the table | Measured by | How |
|---|---|---|
| Page breaks on the same word; within one line; same page count; the flowing copies | `pages.py` | For every page of Word's PDF after the first, its first 8 body words are found in the editor's PDF; the distance to the editor's nearest page start is 0 (same word), up to 15 words (within one line) or more. Running headers, footers and page numbers are left out. |
| Words, paragraphs, pictures after a save | `import_compare.ts` | The keyed original and the editor's save are both read through proto-rescue's own Word import and the blocks are aligned by type and words. "All kept" means every word, every paragraph (same type, same words) and every picture came back. |
| Our paragraph keys kept; Word's paragraph IDs kept | `keys.py` | Every body paragraph of the original gets a hidden bookmark `_kNNNNN` spanning it (`keys.py inject`); after the save, a key is kept when it sits on exactly one paragraph with the same words. Word's own `w14:paraId` is counted the same way. |
| Edits behave like Word's bookmarks | `edits.py` | The keyed copy goes through eight scripted edits (type, split, delete inside a paragraph, copy and paste in front of another, cut and paste in front of another, delete a paragraph, type a new one, join two) and each case is judged against what Word's bookmarks do. |
| Pages identical when Word reopens the save | `pixels.py` | Word's print of the original against Word's print of the editor's save, page by page at 40 dpi in grey: identical means under 0.5% of pixels differ. Word's print of the save is made by hand on this computer (`scripts/word-pdf.ps1`); where it is missing the editor's own PDF stands in and the row says so. |

`flow.py` makes a flowing copy (every manual page break and per-page section removed) so the editor has to work out
every page itself; its answer key is Word's print of that copy.

## A document set

A folder with `set.json`, `docs/<name>.docx` and the answer key `word/<name>.pdf` (plus `word/<name>-flow.pdf`
for each flowing copy). The answer key is Word's own print, made on this computer with

```bash
powershell -ExecutionPolicy Bypass -File scripts/word-pdf.ps1 -DocsDir <folder of docx> -OutDir <folder for pdf>
```

(Word prints through "Microsoft Print to PDF"; saving as PDF freezes in this computer's Word.) The keyed copies
and flowing copies are derived at run time into the run folder, deterministically.

`set.json`:

```
{"name": "public", "documents": ["a", "b"], "flow": ["a"],
 "edit": {"file": "a", "cases": {
   "type": {"in": PHRASE, "mark": "EDITMARKONE"},         "split": {"in": PHRASE, "mark": "SPLITMARK"},
   "delete_text": {"in": PHRASE},                          "copy_paste": {"copy": PHRASE, "before": PHRASE},
   "move": {"cut": PHRASE, "before": PHRASE},              "delete_paragraph": {"in": PHRASE},
   "new_paragraph": {"after": PHRASE, "mark": "NEWPARAMARK"}, "join": {"first": PHRASE, "second": PHRASE}}}}
```

Each PHRASE occurs once in the edit file and names the paragraph the edit happens in (the first words of the
paragraph for split, copy, cut, before, delete_paragraph and join; the last words for new_paragraph; words in
the middle for type and delete_text). The engine's job is generated from these in that order; an editor driven
by hand instead can put its own `edit/edit-cases.json` beside its outputs, listing only the cases it tried.

Two sets exist:

- **Public** (`scorecard/public-set/`, in this repo): invented documents with the private set's features, built by
  `public-set/make.py`, with their committed Word answer key. This is what the GitHub Actions job runs on every push.
- **Private** (never published): ten real documents from the publisher's own files, their keyed copies and
  two flowing copies, kept outside every repository on one computer. Results from it are never published;
  only the spec column's percentages come from it.

## An editor's outputs

```
pdf/<name>.pdf, pdf/<name>-flow.pdf   the editor's own PDF of each document and flowing copy
saved/<name>-keyed.docx               the editor's save of the keyed copy
edit/<file>-edited.docx               the keyed copy after the scripted edits (edit-cases.json for hand edits)
word-of-saved/<name>.pdf              Word's print of saved/, made by hand with scripts/word-pdf.ps1
timings.json                          seconds per conversion, when known
```

## The spec column

`spec.json` holds ONLYOFFICE's results from the 2026-10-28 test on the private set, as percentages. They are the
specification and are never re-run.

## Cross-check of the pixel measure

The pinned tree ships its own page-by-page comparison (`cargo xtask visual-diff`). `scripts/visual-diff.sh
REF.pdf TEST.pdf OUT` runs it sealed (poppler inside the container) and writes `OUT/result.json`. Run once on
2026-10-29 on a public pair (Word's print of `handbook-calibri` against the engine's own PDF of it, 60 pages):

| | `visual-diff` | ours |
|---|---|---|
| Pages starting on the same line / page breaks on the same word | 60 of 60 | 59 of 59 breaks (`pages.py`) |
| Pages under 0.5% of pixels different | 0 | 0 (`pixels.py`) |
| Pages under 3% different | 56 | 50 |
| Per-page share of pixels different | | correlation 0.78 with ours; the two sort 50 of 60 pages into the same bucket |

The page-break measure agrees exactly. The pixel shares track each other but are not the same number (different
rasterisers, 100 dpi against 40 dpi), and neither tool finds a page identical between Word's print and the
engine's own PDF: that comparison is meant for two Word prints (the original and the reopened save), which is
how the "pages identical" row uses it.

## Continuous runs

`.github/workflows/scorecard.yml` builds the engine image from the pin, scores the public set with the free
stand-in fonts (Carlito, Caladea, Liberation), appends one row to `scores.csv` (`csv-row.py`) and commits it.
No scheduled runs: only a push, which is where a layer change, a patch or a pin move arrives. The "nothing lost
on save" measure runs there only when the repository secret `PROTO_RESCUE_TOKEN` (a token that can read the
proto-rescue repository) exists; without it the row says "not run".
