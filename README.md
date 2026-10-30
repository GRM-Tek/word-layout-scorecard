# word-layout-scorecard

A page-by-page scorecard of a document editor against Microsoft Word, and a public set of invented
Word documents with Word's own prints as the answer key. Published by Proto Rescue, which builds a
document editor on the WordCraft engine and measures it before every change.

One command prints the table and writes `score.json`:

```bash
scorecard/score wordcraft scorecard/public-set --no-import
```

Read `scorecard/README.md` for the four measures and the layout of a document set, and
`scorecard/public-set/README.md` for the documents. Every document is invented by `make.py` from a
built-in vocabulary (byte-identical on each run); no real document, organisation or person is in them.

## What you need

- Docker. `scripts/build.sh` builds the engine's command-line tool from the `wordcraft/` submodule, pinned
  to one commit (`git submodule update --init` first; moving the pin is a deliberate change with a score
  attached). `scripts/run-sealed.sh` runs it with no network and the inputs read-only.
- Python 3.13 with `pymupdf` (`py -3.13 -I` on Windows; `python3 -I` elsewhere).
- Fonts: the runner mounts this computer's fonts read-only. With Microsoft's fonts installed the numbers
  are the ones we publish; with Carlito, Caladea and Liberation the page-break rows move and the table says so.
- Microsoft Word on Windows, only to make an answer key (`scripts/word-pdf.ps1`) or to print an editor's
  saved copies for the "pages identical when Word reopens the save" row. The committed answer key already
  covers the public set.
- The "words, paragraphs, pictures after a save" row reads both files through a private Word import and is
  skipped with `--no-import`.

## Results on this set (engine `wordcraft-cli` at `aadb077e`, Microsoft's fonts, 2026-10-29)

| Measure | Result |
|---|---|
| Page breaks on the same word | 741 of 830 (89%); 751 within a line; 11 of 11 page counts |
| The two flowing copies, same word | 7 of 38 and 9 of 17, all within a line |
| Pages identical when Word reopens the save | 783 of 843 (93%) |
| Hidden paragraph bookmarks kept through a save | 9,575 of 9,585 |
| 414-page book to PDF | 8 s |

## Licences

This repository is MIT (`LICENSE`). The `wordcraft/` submodule is the WordCraft engine by the ArtCraft
team, under its own MIT or Apache-2.0 licence; the Docker image keeps its licence files. The answer-key
PDFs are Word's prints of invented documents: they may be copied anywhere except into WordCraft's own
repository, whose clean-room rule excludes files made by Word.
