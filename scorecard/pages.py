"""Measure 1: where pages break, each candidate PDF against Microsoft Word's own PDF of the same file.

For every Word page after the first, take the first 8 body words and find them in the candidate's word stream; the
distance (in words) from there to the candidate's nearest page start says how far that page break moved: 0 is the
same word, up to 15 words is within one line. Running headers and footers (a line on 30%+ of pages, page numbers)
are left out so they do not count as body words. Counts only, never document text.

  pages.py WORD.pdf CANDIDATE.pdf      prints the comparison as JSON
"""
import collections, json, re, sys
import pymupdf

NUM = re.compile(r'\d+')
K = 8               # words that identify a page start
LINE = 15           # words: a break this close is "within one line"

def norm_line(s): return re.sub(r'\s+', ' ', NUM.sub('#', s.lower())).strip()
def tok(s): return re.sub(r'[^0-9a-z]', '', s.lower())

def pages(pdf):
    """Each page as a list of text lines, in reading order."""
    doc = pymupdf.open(pdf)
    out = []
    for pg in doc:
        lines = collections.OrderedDict()
        for x0, y0, x1, y1, w, b, l, n in pg.get_text('words', sort=True):
            lines.setdefault((b, l), []).append(w)
        out.append([' '.join(ws) for ws in lines.values()])
    doc.close()
    return out

def body_tokens(pgs):
    """Each page's body words, running headers, footers and page numbers left out."""
    count = collections.Counter()
    for p in pgs:
        for ln in set(norm_line(x) for x in p): count[ln] += 1
    rep = {ln for ln, c in count.items() if len(pgs) >= 3 and c >= max(2, 0.3 * len(pgs))}
    pn = re.compile(r'^(page )?#( of #)?$|^# ?/ ?#$|^-? ?# ?-?$')
    res = []
    for p in pgs:
        toks = []
        for ln in p:
            k = norm_line(ln)
            if k in rep or pn.match(k): continue
            toks += [t for t in (tok(w) for w in ln.split()) if t]
        res.append(toks)
    return res

def stream(pt):
    s, starts = [], []
    for t in pt:
        starts.append(len(s)); s += t
    return s, starts

def breaks(word_pages, cand_pages):
    """How each of Word's page breaks fared in the candidate (token lists per page)."""
    ws, wst = stream(word_pages); cs, cst = stream(cand_pages)
    idx = collections.defaultdict(list)
    for i in range(len(cs) - K + 1): idx[tuple(cs[i:i + K])].append(i)
    same = near = far = lost = 0
    for k in range(1, len(wst)):
        s = wst[k]
        if len(word_pages[k]) < K: continue          # a near-empty page: nothing to match on
        hits = None
        for off in range(0, 6):
            hits = idx.get(tuple(ws[s + off:s + off + K]))
            if hits: break
        if not hits: lost += 1; continue
        expect = s / max(1, len(ws)) * len(cs)
        m = min(hits, key=lambda h: abs(h - expect)) - off
        d = min((m - c for c in cst), key=abs)
        if d == 0: same += 1
        elif abs(d) <= LINE: near += 1
        else: far += 1
    return same, near, far, lost, len(ws), len(cs)

def compare(word_pdf, cand_pdf):
    wp, cp = body_tokens(pages(word_pdf)), body_tokens(pages(cand_pdf))
    same, near, far, lost, nw, nc = breaks(wp, cp)
    b_same, b_near, b_far, b_lost, _, _ = breaks(cp, wp)
    return {'word_pages': len(wp), 'cand_pages': len(cp), 'breaks_checked': same + near + far,
            'same_word': same, 'within_a_line': near, 'further': far, 'unmatched': lost,
            'extra_breaks': b_near + b_far + b_lost, 'word_tokens': nw, 'cand_tokens': nc}

if __name__ == '__main__':
    print(json.dumps(compare(sys.argv[1], sys.argv[2])))
