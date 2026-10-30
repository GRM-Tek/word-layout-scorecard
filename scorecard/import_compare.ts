// Measure 2: read a keyed original and the copy an editor saved, both through proto-rescue's own Word import
// (parseDocxAsNewDocument), and count what changed. Prints counts only, never document text.
//
//   cd <proto-rescue checkout> && npx tsx <this file> ORIGINAL.docx SAVED.docx...
//
// The import is loaded from the current working directory, which must be a proto-rescue checkout.
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

type B = Record<string, unknown> & { type: string };

async function main() {
  const lib = (f: string) => import(pathToFileURL(path.join(process.cwd(), 'api', '_lib', f)).href);
  const { parseDocxAsNewDocument } = await lib('docxToBlocks.ts');
  const { blockText } = await lib('documentChunks.ts');
  const words = (b: B) => String(blockText(b as never)).replace(/\s+/g, ' ').trim();
  const full = (b: B) => { const { key: _k, block_key: _bk, ...rest } = b as Record<string, unknown>; return JSON.stringify(rest); };

  function read(file: string) {
    try {
      const bytes = new Uint8Array(fs.readFileSync(file));
      let p;
      try { p = parseDocxAsNewDocument(bytes); }
      catch (e) {
        // Crossed-out words: answer "keep them as normal words" for every passage, as a person could.
        const passages = (e as { passages?: Array<{ id: string; fingerprint: string }> }).passages;
        if (!passages) throw e;
        const crossedOut = new Map(passages.map((x) => [x.id, { id: x.id, fingerprint: x.fingerprint, choice: 'keep' as const }]));
        p = parseDocxAsNewDocument(bytes, { crossedOut });
      }
      return { blocks: p.blocks as unknown as B[], report: p.report as unknown as Record<string, unknown>, pictures: p.pictures.length as number };
    } catch (e) { return { error: (e as Error).name || 'Error' }; }
  }

  // Myers diff on interned signatures: returns matched index pairs.
  function align(a: string[], b: string[]): Array<[number, number]> {
    const n = a.length, m = b.length, max = n + m, off = max, v = new Int32Array(2 * max + 2), trace: Int32Array[] = [];
    for (let d = 0; d <= max; d++) {
      trace.push(v.slice());
      for (let k = -d; k <= d; k += 2) {
        let x = k === -d || (k !== d && v[off + k - 1] < v[off + k + 1]) ? v[off + k + 1] : v[off + k - 1] + 1;
        let y = x - k;
        while (x < n && y < m && a[x] === b[y]) { x++; y++; }
        v[off + k] = x;
        if (x >= n && y >= m) {
          const pairs: Array<[number, number]> = [];
          let cx = n, cy = m;
          for (let dd = d; dd > 0; dd--) {
            const vv = trace[dd], kk = cx - cy;
            const pk = kk === -dd || (kk !== dd && vv[off + kk - 1] < vv[off + kk + 1]) ? kk + 1 : kk - 1;
            const px = vv[off + pk], py = px - pk;
            while (cx > px && cy > py) { pairs.push([cx - 1, cy - 1]); cx--; cy--; }
            cx = px; cy = py;
          }
          while (cx > 0 && cy > 0) { pairs.push([cx - 1, cy - 1]); cx--; cy--; }
          return pairs.reverse();
        }
      }
    }
    return [];
  }

  function typeCounts(bs: B[]) { const c: Record<string, number> = {}; for (const b of bs) c[b.type] = (c[b.type] ?? 0) + 1; return c; }
  const wordsAll = (bs: B[]) => bs.map(words).join(' ').split(' ').filter(Boolean).length;

  // Definition blocks (styles, list definitions, the sheet) are not content: an editor may write more or fewer
  // of them without a word changing. They are counted apart and left out of the alignment.
  const DEFINITIONS = new Set(['style', 'listdef', 'sheet']);
  const content = (bs: B[]) => bs.filter((b) => !DEFINITIONS.has(b.type));
  const [orig, ...saved] = process.argv.slice(2);
  const o0 = read(orig);
  if ('error' in o0) { console.log(JSON.stringify({ file: path.basename(orig), error: o0.error })); return; }
  const o = { ...o0, blocks: content(o0.blocks) };
  for (const s of saved) {
    const r0 = read(s);
    if ('error' in r0) { console.log(JSON.stringify({ file: path.basename(s), error: r0.error })); continue; }
    const r = { ...r0, blocks: content(r0.blocks) };
    const pairs = align(o.blocks.map((b) => b.type + '\u0001' + words(b)), r.blocks.map((b) => b.type + '\u0001' + words(b)));
    const sameLook = pairs.filter(([i, j]) => full(o.blocks[i]) === full(r.blocks[j])).length;
    // Which fields differ on blocks with the same words (field names only).
    const fieldDiffs: Record<string, number> = {};
    for (const [i, j] of pairs) {
      const a = o.blocks[i] as Record<string, unknown>, b = r.blocks[j] as Record<string, unknown>;
      for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) {
        if (k === 'key' || k === 'block_key') continue;
        if (JSON.stringify(a[k]) !== JSON.stringify(b[k])) fieldDiffs[`${a.type}.${k}`] = (fieldDiffs[`${a.type}.${k}`] ?? 0) + 1;
      }
    }
    const matchedO = new Set(pairs.map((p) => p[0])), matchedR = new Set(pairs.map((p) => p[1]));
    const oWords = o.blocks.filter((_, i) => !matchedO.has(i)).map(words).join(' ').split(' ').filter(Boolean);
    const rWords = r.blocks.filter((_, j) => !matchedR.has(j)).map(words).join(' ').split(' ').filter(Boolean);
    const reportDiff: Record<string, [unknown, unknown]> = {};
    for (const k of new Set([...Object.keys(o.report), ...Object.keys(r.report)])) {
      if (JSON.stringify(o.report[k]) !== JSON.stringify(r.report[k])) reportDiff[k] = [o.report[k], r.report[k]];
    }
    console.log(JSON.stringify({
      file: path.basename(s), blocks: [o.blocks.length, r.blocks.length], words: [wordsAll(o.blocks), wordsAll(r.blocks)],
      sameTypeAndWords: pairs.length, ofThoseSameFormattingAndLook: sameLook,
      onlyInOriginal: o.blocks.length - pairs.length, onlyInSaved: r.blocks.length - pairs.length,
      unmatchedWords: [oWords.length, rWords.length], pictures: [o.pictures, r.pictures],
      definitionBlocks: [o0.blocks.length - o.blocks.length, r0.blocks.length - r.blocks.length],
      types: [typeCounts(o.blocks), typeCounts(r.blocks)], fieldsChanged: fieldDiffs, reportChanged: reportDiff,
    }));
  }
}
main().catch((e) => { console.log(JSON.stringify({ error: String((e as Error).message || e) })); process.exit(1); });
