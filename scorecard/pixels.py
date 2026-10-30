"""Measure 4: how alike two PDFs of the same document look, page by page. Each page is drawn at 40 dpi in grey and
the share of pixels that differ clearly (by more than 48 of 255 levels) is bucketed: the same (under 0.5%), close
(0.5 to 3%), different (over 3%). Counts only.

  pixels.py A.pdf B.pdf      prints the buckets as JSON
"""
import json, sys
import pymupdf

DPI = 40
LEVELS = 48

def grey(page): return page.get_pixmap(dpi=DPI, colorspace=pymupdf.csGRAY)

def compare(a, b):
    A, B = pymupdf.open(a), pymupdf.open(b)
    buckets = {'same': 0, 'close': 0, 'different': 0}
    shares = []
    for i in range(min(A.page_count, B.page_count)):
        pa, pb = grey(A[i]), grey(B[i])
        if (pa.width, pa.height) != (pb.width, pb.height):
            buckets['different'] += 1; shares.append(1.0); continue
        sa, sb = pa.samples, pb.samples
        diff = sum(1 for x, y in zip(sa, sb) if abs(x - y) > LEVELS) / len(sa)
        shares.append(round(diff, 4))
        buckets['same' if diff < .005 else 'close' if diff < .03 else 'different'] += 1
    r = {'pages_a': A.page_count, 'pages_b': B.page_count, 'compared': len(shares), **buckets, 'share_per_page': shares}
    A.close(); B.close()
    return r

if __name__ == '__main__':
    r = compare(sys.argv[1], sys.argv[2]); r.pop('share_per_page')
    print(json.dumps(r))
