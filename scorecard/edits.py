"""Measure 3b: after real edits, did each paragraph key behave like one of Word's own bookmarks?

  edits.py KEYED.docx EDITED.docx CASES.json      prints the verdicts as JSON (never document text)

CASES.json names each edit by a phrase that occurs once in the keyed document (see scorecard/README.md):
  type              {"in": phrase, "mark": MARK}        text typed inside a paragraph: its key stays, words intact
  split             {"in": phrase, "mark": MARK}        Enter inside a paragraph: one half keeps the key, the other has none
  delete_text       {"in": phrase}                      characters deleted inside a paragraph: its key stays
  copy_paste        {"copy": phrase, "before": phrase}  a paragraph copied in front of another: the other keeps its own key
  move              {"cut": phrase, "before"?: phrase}  a paragraph cut and pasted in front of another: its key travels
                                                      with it and the other paragraph keeps its own
  delete_paragraph  {"in": phrase}                      a paragraph deleted: its key goes with it
  new_paragraph     {"mark": MARK}                      a paragraph typed new: it has no key
  join              {"first": phrase, "second": phrase} two paragraphs joined: the joined paragraph keeps a key of theirs
"""
import collections, difflib, json, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from keys import where

MARKS = re.compile(r'(EDITMARKONE|SPLITMARK|NEWPARAMARK) ?')
def unmarked(t): return re.sub(r'\s+', ' ', MARKS.sub('', t)).strip()

class Judge:
    def __init__(self, keyed, edited):
        self.ok, self.oi, self.ot = where(keyed)
        self.nk, self.ni, self.nt = where(edited)
        self.okey = {v[0]: k for k, v in self.ok.items() if len(v) == 1}     # original paragraph -> its key
        self.nkeys = collections.defaultdict(list)                            # edited paragraph -> its keys
        for k, v in self.nk.items():
            for j in v: self.nkeys[j].append(k)

    def orig(self, phrase):
        hits = [i for i, t in enumerate(self.ot) if phrase in t]
        if len(hits) != 1: raise ValueError('phrase found %d times in the keyed document' % len(hits))
        return hits[0]

    def key_of(self, phrase):
        i = self.orig(phrase)
        if i not in self.okey: raise ValueError('that paragraph carries no single key')
        return i, self.okey[i]

    def at(self, key): return self.nk.get(key, [])

    def case_type(self, c):
        i, k = self.key_of(c['in']); at = self.at(k)
        if len(at) != 1: return False, 'key on %d paragraphs' % len(at)
        t = self.nt[at[0]]
        if c['mark'] not in t: return False, 'key not on the paragraph with the typed text'
        same = unmarked(t) == self.ot[i]
        return same, 'key stays, words intact' if same else 'key stays but the words changed'

    def case_split(self, c):
        i, k = self.key_of(c['in']); at = self.at(k)
        if len(at) != 1: return False, 'key on %d paragraphs' % len(at)
        j = at[0]; half = unmarked(self.nt[j])
        if not half or half not in self.ot[i]: return False, 'key on a paragraph that is not a half of the original'
        others = [x for x in (j - 1, j + 1) if 0 <= x < len(self.nt)
                  and unmarked(self.nt[x]) and unmarked(self.nt[x]) in self.ot[i]]
        if not others: return False, 'the other half was not found beside it'
        if any(self.nkeys.get(x) for x in others): return False, 'both halves carry a key'
        return True, 'one half keeps the key, the other has none'

    def case_delete_text(self, c):
        i, k = self.key_of(c['in']); at = self.at(k)
        if len(at) != 1: return False, 'key on %d paragraphs' % len(at)
        t = unmarked(self.nt[at[0]]); o = self.ot[i]
        close = len(o) - 40 <= len(t) < len(o) and difflib.SequenceMatcher(None, o, t).ratio() > 0.9
        return close, 'key stays on its shortened paragraph' if close else 'key on other words'

    def case_copy_paste(self, c):
        ic, kc = self.key_of(c['copy']); ib, kb = self.key_of(c['before'])
        at = self.at(kb)
        if len(at) != 1: return False, 'the paragraph pasted in front of has its key on %d paragraphs' % len(at)
        if unmarked(self.nt[at[0]]) != self.ot[ib]:
            return False, 'the key of the paragraph pasted in front of moved to other words'
        copies = [j for j in self.at(kc) if unmarked(self.nt[j]) == self.ot[ic]]
        if not copies: return False, 'the copied paragraph lost its key'
        return True, 'the paragraph pasted in front of keeps its key' + (' (the copy carries the same key)' if len(copies) > 1 else '')

    def case_move(self, c):
        ic, kc = self.key_of(c['cut']); at = self.at(kc)
        if len(at) != 1: return False, 'key on %d paragraphs' % len(at)
        j = at[0]
        if unmarked(self.nt[j]) != self.ot[ic]: return False, 'key on other words'
        if 'before' in c:
            ib, kb = self.key_of(c['before'])
            if j + 1 >= len(self.nt) or unmarked(self.nt[j + 1]) != self.ot[ib]:
                return False, 'key travelled but not to the new place'
            if self.at(kb) != [j + 1]:
                return False, 'key travelled but the paragraph pasted in front of lost its own key'
        return True, 'key travelled with its paragraph'

    def case_delete_paragraph(self, c):
        i, k = self.key_of(c['in'])
        gone = not self.at(k)
        return gone, 'key gone with its paragraph' if gone else 'key still present'

    def case_new_paragraph(self, c):
        js = [j for j, t in enumerate(self.nt) if c['mark'] in t]
        if len(js) != 1: return False, 'new paragraph found %d times' % len(js)
        none = not self.nkeys.get(js[0])
        return none, 'new paragraph has no key' if none else 'new paragraph carries a key'

    def case_join(self, c):
        ia, ka = self.key_of(c['first']); ib, kb = self.key_of(c['second'])
        js = [j for j, t in enumerate(self.nt)
              if unmarked(t).startswith(self.ot[ia][:40]) and self.ot[ib][:40] in unmarked(t)]
        if len(js) != 1: return False, 'joined paragraph found %d times' % len(js)
        held = [k for k in self.nkeys.get(js[0], []) if k in (ka, kb)]
        if not held: return False, 'joined paragraph holds neither key'
        if any(len(self.at(k)) != 1 for k in held): return False, 'a held key also sits elsewhere'
        which = 'both keys' if len(held) == 2 else 'the first key' if held[0] == ka else 'the second key'
        return True, 'joined paragraph keeps ' + which

    def untouched(self):
        """Paragraphs whose words exist once in both files: did the key travel with the words?"""
        oc, nc = collections.Counter(self.ot), collections.Counter(self.nt)
        pos = {t: j for j, t in enumerate(self.nt)}
        c = collections.Counter()
        for i, t in enumerate(self.ot):
            if not t or oc[t] != 1 or nc[t] != 1 or i not in self.okey: continue
            c['key travelled with its words' if self.okey[i] in self.nkeys.get(pos[t], []) else 'words kept, key not on them'] += 1
        return dict(c)

def judge(keyed, edited, cases):
    j = Judge(keyed, edited)
    out, passed = {}, 0
    for name, c in cases.items():
        fn = getattr(j, 'case_' + name, None)
        if fn is None: out[name] = {'ok': False, 'why': 'unknown case'}; continue
        try: ok, why = fn(c)
        except ValueError as e: ok, why = False, 'cannot judge: %s' % e
        out[name] = {'ok': bool(ok), 'why': why}; passed += bool(ok)
    lost = [k for k in j.ok if k not in j.nk]; doubled = [k for k, v in j.nk.items() if len(v) > 1]
    return {'cases': out, 'passed': passed, 'of': len(cases), 'untouched': j.untouched(),
            'keys_gone': len(lost), 'keys_doubled': len(doubled),
            'paragraphs_without_key': sum(1 for x in range(len(j.nt)) if x not in j.nkeys),
            'paragraphs': [len(j.ot), len(j.nt)]}

if __name__ == '__main__':
    with open(sys.argv[3], encoding='utf8') as f: cases = json.load(f)        # set.json, edit-cases.json or bare cases
    cases = cases.get('edit', cases); cases = cases.get('cases', cases)
    print(json.dumps(judge(sys.argv[1], sys.argv[2], cases), indent=1))
