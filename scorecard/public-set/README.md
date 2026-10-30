# Public document set

This folder makes the public set of Word documents for the page-layout scorecard.
Every document is invented: the text is made up by the script from a built-in word list,
the pictures are coloured rectangles drawn by the script, and the place names are made up.
No real document, organisation or person is in any of them, so the set and its results can be
published.

To make the set again, run this from the top folder of the repo:

```bash
py -3.13 -I scorecard/public-set/make.py
```

It writes the eleven documents into `docs/` (byte-identical each time) and `set.json`, which
lists the documents and the phrases the editing checks look for in the handbook.

The first eight carry the general features (fixed page breaks, a 400-page flowing book, nested
lists, merged cells, triple borders, text boxes, pictures, headers and footers, a section per page).
Three were added on 2026-10-29 because the private set lost pages after a save on features the
eight did not have, each shaped like the private file that loses them:

- `protocols-arial`: a protocol book, one section per protocol, bullet lists in four spacings (a
  Bullet style that switches contextual spacing off, plain numbered paragraphs with space after,
  space before set to nothing, an exact line height), a black box whose bullets come from the style
  itself, and two-column comparison tables with a tall row. It carries a theme part, so its defaults
  name theme fonts and colours the way Word's own files do.
- `notices-calibri`: sections with page borders measured from the page edge in two weights, beside
  sections with none and one with a border measured from the text; a continuous first section;
  margins that change from section to section; fixed-layout tables on most pages.
- `register-aptos`: tables of seven kinds: widths from the grid alone, widths in twips under a fixed
  layout, a table in percent with cell spacing and centred rows, rows with a height, a floating
  table, automatic cell widths, merged cells with uneven fixed widths.

The Word answer key in `word/` is made by printing each document through Microsoft Print to
PDF on a Windows computer with Word; that is a separate step and this script does not do it.
