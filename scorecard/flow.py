"""A flowing copy: take out every manual page break and every per-page section break, so the editor has to work out
every page break itself. Words, tables and pictures are untouched; the last section's page set-up applies.

  flow.py IN.docx OUT.docx
"""
import re, sys, zipfile
import xml.etree.ElementTree as ET

TAG = re.compile(r'<w:sectPr\b[^>]*?(/?)>|</w:sectPr>')

def flow_xml(x):
    spans, depth, start = [], 0, None
    for m in TAG.finditer(x):
        if m.group(0) == '</w:sectPr>':
            depth -= 1
            if depth == 0: spans.append((start, m.end()))
        elif m.group(1) == '/':
            if depth == 0: spans.append((m.start(), m.end()))
        else:
            if depth == 0: start = m.start()
            depth += 1
    for s, e in reversed(spans[:-1]):          # keep the body's own (last) section
        x = x[:s] + x[e:]
    x = re.sub(r'<w:br [^>]*w:type="page"[^>]*/>', '', x)
    x = re.sub(r'<w:pageBreakBefore(?: [^>]*)?/>', '', x)
    ET.fromstring(x.encode('utf8'))           # refuse to write broken XML
    return x, len(spans) - 1

def flow(src, dst):
    z = zipfile.ZipFile(src)
    x, removed = flow_xml(z.read('word/document.xml').decode('utf8'))
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as w:
        for item in z.infolist():
            w.writestr(item, x.encode('utf8') if item.filename == 'word/document.xml' else z.read(item.filename))
    return removed

if __name__ == '__main__':
    print('sections removed:', flow(sys.argv[1], sys.argv[2]))
