"""Paragraph keys through an editor save.

  keys.py inject <in.docx> <out.docx>   give every body paragraph a hidden bookmark _kNNNNN spanning it
  keys.py check <keyed.docx> <saved.docx>...   where each key ended up; counts only, never document text

A key is 'kept' when it sits on exactly one paragraph whose words equal its original paragraph's, 'kept, edited'
when that one paragraph's words changed, 'moved' when it sits on another paragraph's words, 'lost' when it is gone,
'doubled' when two or more paragraphs carry it. Word's own paragraph ids (w14:paraId) are counted the same way."""
import re, sys, zipfile, collections, json, os, html

TAG = re.compile(r'<w:p(?=[\s>/])[^>]*?(/?)>|</w:p>')
def paragraphs(xml):
    """(start, end, inner xml) of every <w:p>, nesting-aware; a <w:p/> is an empty paragraph."""
    out, stack = [], []
    for m in TAG.finditer(xml):
        if m.group(0) == '</w:p>':
            s, inner_start = stack.pop(); out.append((s, m.end(), xml[inner_start:m.start()], xml[s:inner_start]))
        elif m.group(1) == '/':
            out.append((m.start(), m.end(), '', m.group(0)))
        else:
            stack.append((m.start(), m.end()))
    return sorted(out)

def text(inner):
    inner = re.sub(r'<w:delText[^>]*>.*?</w:delText>', '', inner, flags=re.S)
    return re.sub(r'\s+', ' ', html.unescape(''.join(re.findall(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>', inner)))).strip()

def own(inner):
    """The paragraph's own markup, without paragraphs nested inside it (text boxes)."""
    return re.sub(r'<w:txbxContent>.*?</w:txbxContent>', '', inner, flags=re.S)

def inject(src, dst):
    z = zipfile.ZipFile(src)
    xml = z.read('word/document.xml').decode('utf8')
    edits, n = [], 0
    stack = []
    for m in TAG.finditer(xml):
        if m.group(0) == '</w:p>':
            k = stack.pop()
            edits.append((m.start(), f'<w:bookmarkEnd w:id="{900000 + k}"/>'))
        else:
            n += 1
            start = f'<w:bookmarkStart w:id="{900000 + n}" w:name="_k{n:05d}"/>'
            if m.group(1) == '/':      # <w:p .../> -> <w:p ...>bookmark</w:p>
                open_tag = m.group(0)[:-2] + '>'
                edits.append((m.start(), None, m.end(), open_tag + start + f'<w:bookmarkEnd w:id="{900000 + n}"/></w:p>'))
                continue
            after = m.end()
            pp = re.compile(r'<w:pPr\b.*?</w:pPr>|<w:pPr/>', re.S).match(xml, after)
            edits.append((pp.end() if pp else after, start))
            stack.append(n)
    out, last = [], 0
    for e in sorted(edits, key=lambda e: e[0]):
        if len(e) == 4:
            out.append(xml[last:e[0]]); out.append(e[3]); last = e[2]
        else:
            out.append(xml[last:e[0]]); out.append(e[1]); last = e[0]
    out.append(xml[last:])
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as w:
        for item in z.infolist():
            data = ''.join(out).encode('utf8') if item.filename == 'word/document.xml' else z.read(item.filename)
            w.writestr(item, data)
    print(f'{os.path.basename(dst)}: {n} paragraphs keyed')

def where(docx):
    xml = zipfile.ZipFile(docx).read('word/document.xml').decode('utf8')
    keys, ids, texts = collections.defaultdict(list), collections.defaultdict(list), []
    for i, (s, e, inner, open_tag) in enumerate(paragraphs(xml)):
        t = text(own(inner)); texts.append(t)
        for k in re.findall(r'<w:bookmarkStart[^>]*w:name="(_k\d{5})"', own(inner)): keys[k].append(i)
        pid = re.search(r'w14:paraId="([0-9A-Fa-f]{8})"', open_tag)
        if pid: ids[pid.group(1).upper()].append(i)
    return keys, ids, texts

def classify(orig_map, orig_texts, new_map, new_texts):
    c = collections.Counter()
    for k, (i,) in ((k, v) for k, v in orig_map.items() if len(v) == 1):
        at = new_map.get(k, [])
        if not at: c['lost'] += 1
        elif len(at) > 1: c['doubled'] += 1
        elif new_texts[at[0]] == orig_texts[i]: c['kept'] += 1
        elif orig_texts[i] and orig_texts[i] in new_texts[at[0]] or new_texts[at[0]] and new_texts[at[0]] in orig_texts[i]:
            c['kept, edited'] += 1
        else: c['moved'] += 1
    return dict(c)

if __name__ == '__main__':
    if sys.argv[1] == 'inject': inject(sys.argv[2], sys.argv[3])
    else:
        ok, oi, ot = where(sys.argv[2])
        for f in sys.argv[3:]:
            nk, ni, nt = where(f)
            print(json.dumps({'file': os.path.basename(f), 'paragraphs': [len(ot), len(nt)],
                              'bookmark keys': classify(ok, ot, nk, nt), 'word paraIds': classify(oi, ot, ni, nt),
                              'paraIds in original': sum(len(v) == 1 for v in oi.values())}))
