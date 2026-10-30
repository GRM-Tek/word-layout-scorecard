"""The scorecard: one command that measures an editor against Microsoft Word on a document set.

  score <editor> <set-folder> [--outputs DIR] [--out DIR] [--fonts DIR[;DIR]] [--proto-rescue DIR] [--no-import]

<editor> is `wordcraft` (the pinned engine is run, sealed, to make its outputs) or any other name whose outputs
already lie in --outputs DIR (for example `onlyoffice`, from the 2026-10-28 test). The spec table from PLAN.md is
printed with the editor's column filled, and score.json is written beside the outputs. Counts only; no document
text is printed or written anywhere but the outputs folder.

A set folder holds set.json, docs/<name>.docx and the answer key word/<name>.pdf (Word's own print, see
scripts/word-pdf.ps1), plus word/<name>-flow.pdf for each flowing copy. An outputs folder holds:
  pdf/<name>.pdf, pdf/<name>-flow.pdf   the editor's own PDF of each document and flowing copy
  saved/<name>-keyed.docx               the editor's save of the keyed copy
  edit/<file>-edited.docx               the keyed copy after the scripted edits (plus edit-cases.json for hand edits)
  word-of-saved/<name>.pdf              Word's print of saved/ (made on this computer; without it the editor's own
                                        PDF stands in for the "pages identical" row and the row says so)
  timings.json                          seconds per conversion, when known
"""
import argparse, datetime, json, os, pathlib, platform, shutil, subprocess, sys, time
HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
import edits, flow, keys, pages, pixels

FAILURES = []                      # engine runs that failed; the command exits 1 when any did

def say(*a): print(*a, file=sys.stderr, flush=True)

def pct(a, b): return '%d%%' % round(100.0 * a / b) if b else '-'
def frac(a, b, with_pct=True): return '{:,} of {:,}'.format(a, b) + (' (%s)' % pct(a, b) if with_pct else '')

# ----------------------------------------------------------------------------------------------- the edit job
def job_commands(cases):
    """The engine's scripted edits for a set's cases, in the 2026-10-28 order, the join case last."""
    sel = lambda t: {'command': 'select.text', 'params': {'text': t}}
    ins = lambda t: {'command': 'text.insert', 'params': {'text': t}}
    out = []
    if 'type' in cases:
        c = cases['type']; out += ['caret.docStart', sel(c['in']), 'caret.right', ins(' ' + c['mark'])]
    if 'split' in cases:
        c = cases['split']; out += ['caret.docStart', sel(c['in']), 'caret.right', 'text.newParagraph', ins(c['mark'] + ' ')]
    if 'delete_text' in cases:
        c = cases['delete_text']; out += ['caret.docStart', sel(c['in']), 'caret.right', 'text.backspace']
    if 'copy_paste' in cases:
        c = cases['copy_paste']
        out += ['caret.docStart', sel(c['copy']), 'select.paragraph', 'edit.copy', 'caret.docStart', sel(c['before']), 'caret.left', 'edit.paste']
    if 'move' in cases:
        c = cases['move']
        out += ['caret.docStart', sel(c['cut']), 'select.paragraph', 'edit.cut', 'caret.docStart', sel(c['before']), 'caret.left', 'edit.paste']
    if 'delete_paragraph' in cases:
        c = cases['delete_paragraph']; out += ['caret.docStart', sel(c['in']), 'select.paragraph', 'text.delete']
    if 'new_paragraph' in cases:
        c = cases['new_paragraph']
        out += ['caret.docStart', sel(c['after']), 'caret.right', 'text.newParagraph', ins(c['mark'] + ' typed as a brand new paragraph.')]
    if 'join' in cases:
        # caret.home, not caret.left: a backspace right after collapsing the selection at the paragraph start
        # deletes the paragraph's bookmark instead of joining (seen on the pinned commit).
        c = cases['join']; out += ['caret.docStart', sel(c['second']), 'caret.home', 'text.backspace']
    return out

# ----------------------------------------------------------------------------------------------- the engine, sealed
def default_fonts():
    dirs = []
    if platform.system() == 'Windows':
        win = pathlib.Path(os.environ.get('SystemRoot', 'C:/Windows')) / 'Fonts'
        if win.is_dir(): dirs.append(win)
        cloud = pathlib.Path(os.environ.get('LOCALAPPDATA', '')) / 'Microsoft/FontCache/4/CloudFonts'
        for d in sorted(cloud.glob('Aptos*')) if cloud.is_dir() else []: dirs.append(d)
    else:
        for d in ('/usr/share/fonts', '/usr/local/share/fonts'):
            if pathlib.Path(d).is_dir(): dirs.append(pathlib.Path(d))
    return dirs

def fonts_report(dirs):
    names = set()
    for d in dirs:
        for p in pathlib.Path(d).rglob('*'):
            if p.suffix.lower() in ('.ttf', '.ttc', '.otf'): names.add(p.name.lower())
    have = lambda *fs: any(f in names for f in fs)
    return {'folders': [str(d) for d in dirs],
            'microsoft': {'calibri': have('calibri.ttf'), 'cambria': have('cambria.ttc'), 'arial': have('arial.ttf'),
                          'aptos': any(str(d).lower().endswith('aptos') for d in dirs) or have('aptos.ttf')},
            'stand_ins': {'carlito': have('carlito-regular.ttf'), 'caladea': have('caladea-regular.ttf'),
                          'liberation_sans': have('liberationsans-regular.ttf')}}

def git_bash():
    """The bash that runs scripts/run-sealed.sh: on Windows, Git's (the one on PATH may be WSL's, which cannot
    see C:/ paths); the `score` wrapper passes its own in SCORE_BASH."""
    if os.environ.get('SCORE_BASH'): return os.environ['SCORE_BASH']
    if platform.system() == 'Windows':
        for c in (pathlib.Path(os.environ.get('EXEPATH', '')) / 'bash.exe', pathlib.Path('C:/Program Files/Git/bin/bash.exe')):
            if c.is_file(): return str(c)
    return 'bash'

def sealed(in_dir, out_dir, fonts, *cli_args):
    posix = lambda p: str(p).replace('\\', '/')            # bash eats the backslashes of a Windows path
    cmd = [git_bash(), posix(ROOT / 'scripts/run-sealed.sh'), posix(in_dir), posix(out_dir)]
    if fonts: cmd.append(';'.join(posix(f) for f in fonts))
    cmd += ['--', *cli_args]
    t0 = time.time()
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        say('  engine failed:', ' '.join(cli_args), '\n', (r.stdout + r.stderr)[-2000:])
        FAILURES.append(' '.join(cli_args))
        return None
    return round(time.time() - t0, 1)

def make_outputs(setdir, cfg, out, fonts):
    """Run the pinned engine over the set: PDFs, saves of the keyed copies, the scripted edits."""
    stage = out / 'in'; stage.mkdir(parents=True, exist_ok=True)
    derived = out / 'derived'
    for d in cfg['documents']:
        shutil.copyfile(setdir / 'docs' / (d + '.docx'), stage / (d + '.docx'))
        shutil.copyfile(derived / (d + '-keyed.docx'), stage / (d + '-keyed.docx'))
    for d in cfg.get('flow', []):
        shutil.copyfile(derived / (d + '-flow.docx'), stage / (d + '-flow.docx'))
    for sub in ('pdf', 'saved', 'edit', 'job'): (out / sub).mkdir(exist_ok=True)
    timings = {}
    for d in cfg['documents'] + [d + '-flow' for d in cfg.get('flow', [])]:
        say('  pdf', d)
        t = sealed(stage, out, fonts, 'convert', '/in/%s.docx' % d, '/out/pdf/%s.pdf' % d)
        if t is not None: timings[d + '.pdf'] = t
    for d in cfg['documents']:
        say('  save', d)
        t = sealed(stage, out, fonts, 'convert', '/in/%s-keyed.docx' % d, '/out/saved/%s-keyed.docx' % d)
        if t is not None: timings[d + '-keyed.docx'] = t
    edit = cfg.get('edit')
    if edit:
        job = {'file': '/in/%s-keyed.docx' % edit['file'], 'commands': job_commands(edit['cases']),
               'save': ['/out/edit/%s-edited.docx' % edit['file']]}
        (out / 'job/job.json').write_text(json.dumps(job, indent=1), encoding='utf8')
        say('  edits', edit['file'])
        t = sealed(stage, out, fonts, 'run', '--script', '/out/job/job.json')
        if t is not None: timings['edits'] = t
    (out / 'timings.json').write_text(json.dumps(timings, indent=1), encoding='utf8')
    (out / 'fonts.json').write_text(json.dumps(fonts_report(fonts), indent=1), encoding='utf8')
    shutil.rmtree(stage)

# ----------------------------------------------------------------------------------------------- the measures
def derive(setdir, cfg, out):
    derived = out / 'derived'; derived.mkdir(parents=True, exist_ok=True)
    for d in cfg['documents']:
        k = derived / (d + '-keyed.docx')
        if not k.exists():
            say('  keying', d); keys.inject(str(setdir / 'docs' / (d + '.docx')), str(k))
    for d in cfg.get('flow', []):
        f = derived / (d + '-flow.docx')
        if not f.exists():
            say('  flowing copy', d); flow.flow(str(setdir / 'docs' / (d + '.docx')), str(f))
    return derived

def measure_pages(setdir, cfg, outputs):
    res = {}
    for d in cfg['documents'] + [d + '-flow' for d in cfg.get('flow', [])]:
        w, c = setdir / 'word' / (d + '.pdf'), outputs / 'pdf' / (d + '.pdf')
        if w.exists() and c.exists():
            say('  pages', d); res[d] = pages.compare(str(w), str(c))
        else: res[d] = {'missing': [str(p.name) for p in (w, c) if not p.exists()]}
    return res

def measure_keys(cfg, derived, outputs):
    res = {}
    for d in cfg['documents']:
        s = outputs / 'saved' / (d + '-keyed.docx')
        if not s.exists(): res[d] = {'missing': s.name}; continue
        say('  keys', d)
        ok, oi, ot = keys.where(str(derived / (d + '-keyed.docx')))
        nk, ni, nt = keys.where(str(s))
        res[d] = {'paragraphs': [len(ot), len(nt)], 'keys': sum(len(v) == 1 for v in ok.values()),
                  'bookmark_keys': keys.classify(ok, ot, nk, nt),
                  'paraids': sum(len(v) == 1 for v in oi.values()), 'word_paraids': keys.classify(oi, ot, ni, nt)}
    return res

def measure_edits(cfg, derived, outputs):
    edit = cfg.get('edit')
    if not edit: return {'skipped': 'the set names no edit file'}
    e = outputs / 'edit' / (edit['file'] + '-edited.docx')
    if not e.exists(): return {'missing': e.name}
    cases, source = edit['cases'], 'set.json'
    hand = outputs / 'edit' / 'edit-cases.json'
    if hand.exists():
        j = json.loads(hand.read_text(encoding='utf8')); cases, source = j['cases'], 'edit-cases.json (edits made by hand)'
        note = j.get('note')
    else: note = None
    say('  edits', edit['file'])
    r = edits.judge(str(derived / (edit['file'] + '-keyed.docx')), str(e), cases)
    r['cases_from'] = source
    if note: r['note'] = note
    return r

def measure_import(cfg, derived, outputs, proto):
    if proto is None: return {'skipped': 'no proto-rescue checkout (give --proto-rescue DIR)'}
    res = {}
    for d in cfg['documents']:
        s = outputs / 'saved' / (d + '-keyed.docx')
        if not s.exists(): res[d] = {'missing': s.name}; continue
        say('  import', d)
        r = subprocess.run(['npx', 'tsx', str(HERE / 'import_compare.ts'), str(derived / (d + '-keyed.docx')), str(s)],
                           cwd=str(proto), capture_output=True, text=True, shell=(platform.system() == 'Windows'))
        line = [l for l in r.stdout.splitlines() if l.startswith('{')]
        res[d] = json.loads(line[-1]) if line else {'error': (r.stderr or r.stdout)[-500:]}
    return res

def measure_pixels(setdir, cfg, outputs):
    res, via = {}, None
    for d in cfg['documents']:
        w = setdir / 'word' / (d + '.pdf'); own = outputs / 'pdf' / (d + '.pdf')
        ws = outputs / 'word-of-saved' / (d + '.pdf')
        if not ws.exists(): ws = outputs / 'word-of-saved' / (d + '-keyed.pdf')
        if not w.exists(): res[d] = {'missing': w.name}; continue
        if ws.exists(): c, how = ws, 'word-of-saved'
        elif own.exists(): c, how = own, 'own-pdf'
        else: res[d] = {'missing': ws.name}; continue
        via = via or how
        if via != how: res[d] = {'missing': 'mixed sources'}; continue
        say('  pixels', d, '(%s)' % how)
        r = pixels.compare(str(w), str(c)); r.pop('share_per_page'); r['via'] = how; res[d] = r
    return res, via

# ----------------------------------------------------------------------------------------------- the table
def total(res, key, docs):
    return sum(res[d].get(key, 0) for d in docs if d in res and 'missing' not in res[d])

def summarize(cfg, editor, P, K, E, I, X, via, timings, fonts, pin):
    docs = cfg['documents']; flows = cfg.get('flow', [])
    n = len([d for d in docs if 'missing' not in P.get(d, {'missing': 1})])
    same, near, checked = total(P, 'same_word', docs), total(P, 'within_a_line', docs), total(P, 'breaks_checked', docs)
    count_ok = sum(1 for d in docs if 'missing' not in P.get(d, {'missing': 1}) and P[d]['word_pages'] == P[d]['cand_pages'])
    flow_rows = {}
    for d in flows:
        r = P.get(d + '-flow', {})
        if 'missing' in r or not r: flow_rows[d] = None
        else: flow_rows[d] = [r['same_word'], r['within_a_line'], r['breaks_checked'], r['word_pages'], r['cand_pages']]
    kd = [d for d in docs if 'missing' not in K.get(d, {'missing': 1})]
    keys_kept = sum(K[d]['bookmark_keys'].get('kept', 0) for d in kd); keys_all = sum(K[d]['keys'] for d in kd)
    pid_kept = sum(K[d]['word_paraids'].get('kept', 0) for d in kd); pid_all = sum(K[d]['paraids'] for d in kd)
    xd = [d for d in docs if 'missing' not in X.get(d, {'missing': 1})]
    px_same, px_all = sum(X[d]['same'] for d in xd), sum(X[d]['compared'] for d in xd)
    imp = None
    if 'skipped' not in I:
        idd = [d for d in docs if d in I and 'error' not in I[d] and 'missing' not in I[d]]
        imp = {'documents': len(idd), 'errors': [d for d in docs if d in I and 'error' in I[d]],
               'words': [sum(I[d]['words'][0] for d in idd), sum(I[d]['words'][1] for d in idd)],
               'paragraphs': [sum(I[d]['blocks'][0] for d in idd), sum(I[d]['sameTypeAndWords'] for d in idd), sum(I[d]['blocks'][1] for d in idd)],
               'pictures': [sum(I[d]['pictures'][0] for d in idd), sum(I[d]['pictures'][1] for d in idd)],
               'same_look': sum(I[d]['ofThoseSameFormattingAndLook'] for d in idd)}
        imp['all_kept'] = bool(idd) and not imp['errors'] and imp['words'][0] == imp['words'][1] and \
            imp['paragraphs'][0] == imp['paragraphs'][1] == imp['paragraphs'][2] and imp['pictures'][0] == imp['pictures'][1]
    biggest = max(docs, key=lambda d: P.get(d, {}).get('word_pages', 0)) if docs else None
    secs = timings.get(biggest + '.pdf') if timings and biggest else None
    return {'editor': editor, 'set': cfg['name'], 'date': datetime.date.today().isoformat(), 'pin': pin,
            'documents': n, 'same_word': [same, checked], 'within_line': [same + near, checked],
            'page_count_same': [count_ok, n], 'flow': flow_rows,
            'save': imp, 'pixels': [px_same, px_all], 'pixels_via': via,
            'keys': [keys_kept, keys_all], 'paraids': [pid_kept, pid_all],
            'edits': None if 'cases' not in E else {'passed': E['passed'], 'of': E['of'],
                     'failed': [c for c, v in E['cases'].items() if not v['ok']], 'from': E['cases_from'], 'note': E.get('note')},
            'pdf_seconds': {'document': biggest, 'seconds': secs}, 'fonts': fonts}

def table(S, spec, is_private):
    rows = []
    sp = spec['rows']
    def specv(v, textual=None, with_pct=True):
        if textual: return textual
        if isinstance(v, list): return frac(v[0], v[1], with_pct) if is_private else pct(v[0], v[1]) + ' (private spec)'
        return str(v)
    n = S['documents']
    rows.append(('Page breaks on the same word, %d files' % n, specv(sp['same_word']), frac(*S['same_word'])))
    rows.append(('Within one line, %d files' % n, specv(sp['within_line'], with_pct=False), frac(*S['within_line'], with_pct=False)))
    rows.append(('Same page count as Word', specv(sp['page_count'], with_pct=False), frac(*S['page_count_same'], with_pct=False)))
    for d, r in S['flow'].items():
        s = sp['flow'].get(d) if is_private else None
        spec_cell = frac(*s) if s else pct(*sp['flow']['sog']) + ' (private spec, SOG)'
        rows.append(('%s flowing copy, same word' % d.upper(), spec_cell,
                     'not run' if r is None else '%s; %s within a line; pages %d vs %d' % (frac(r[0], r[2]), r[0] + r[1], r[3], r[4])))
    imp = S['save']
    if imp is None: cell = 'not run (needs a proto-rescue checkout)'
    elif imp['all_kept']: cell = 'all kept (%d files)' % imp['documents']
    else:
        cell = 'words %s, paragraphs %s, pictures %s' % (frac(imp['words'][1], imp['words'][0], False),
               frac(imp['paragraphs'][1], imp['paragraphs'][0], False), frac(imp['pictures'][1], imp['pictures'][0], False))
        if imp['errors']: cell += '; import failed on ' + ', '.join(imp['errors'])
    rows.append(('Words, paragraphs, pictures after a save', specv(None, sp['save']), cell))
    px = frac(*S['pixels'])
    if S['pixels_via'] == 'own-pdf': px += " (the editor's own PDF; Word's print of the save not made)"
    rows.append(('Pages identical when Word reopens the save', specv(sp['pixels']), px))
    rows.append(('Our paragraph keys kept, no edits', specv(sp['keys']), frac(*S['keys'])))
    pid = S['paraids']
    rows.append(("Word's own paragraph IDs kept", specv(None, sp['paraids']),
                 'all' if pid[1] and pid[0] == pid[1] else 'none' if pid[1] and pid[0] == 0 else frac(*pid)))
    e = S['edits']
    if e is None: cell = 'not run'
    else:
        cell = ('yes, %d of %d cases' % (e['passed'], e['of'])) if e['passed'] == e['of'] else \
               '%d of %d cases (failed: %s)' % (e['passed'], e['of'], ', '.join(e['failed']).replace('_', ' '))
        if e.get('note'): cell += '; ' + e['note']
    rows.append(('Edits behave like Word\'s bookmarks', specv(None, sp['edits']), cell))
    t = S['pdf_seconds']
    rows.append(('%s to PDF' % (t['document'] or 'biggest file').upper(), specv(None, sp['pdf_seconds']),
                 'not timed' if t['seconds'] is None else '%s s' % t['seconds']))
    rows.append(('Runs inside another web page', specv(None, sp['embed']), 'not measured here (step 3)'))
    w = [max(len(r[i]) for r in rows + [('Measure', 'Spec (' + spec['source'] + ')', S['editor'])]) for i in range(3)]
    head = ('Measure', 'Spec (' + spec['source'] + ')', S['editor'])
    lines = ['| ' + ' | '.join(h.ljust(w[i]) for i, h in enumerate(head)) + ' |', '|' + '|'.join('-' * (x + 2) for x in w) + '|']
    lines += ['| ' + ' | '.join(c.ljust(w[i]) for i, c in enumerate(r)) + ' |' for r in rows]
    return '\n'.join(lines)

# ----------------------------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description='Score an editor against Microsoft Word on a document set.')
    ap.add_argument('editor'); ap.add_argument('set')
    ap.add_argument('--outputs', help='folder holding the editor\'s outputs (default: make them, wordcraft only)')
    ap.add_argument('--out', help='where derived files and score.json go (default: <set>/out/<editor>)')
    ap.add_argument('--fonts', help='font folders for the engine, separated by ; (default: this computer\'s)')
    ap.add_argument('--proto-rescue', help='proto-rescue checkout for the import measure (default: $PROTO_RESCUE)')
    ap.add_argument('--no-import', action='store_true', help='skip the import measure')
    a = ap.parse_args()
    setdir = pathlib.Path(a.set).resolve()
    cfg = json.loads((setdir / 'set.json').read_text(encoding='utf8'))
    out = pathlib.Path(a.out).resolve() if a.out else setdir / 'out' / a.editor
    out.mkdir(parents=True, exist_ok=True)
    outputs = pathlib.Path(a.outputs).resolve() if a.outputs else out
    fonts = [pathlib.Path(f) for f in a.fonts.split(';') if f] if a.fonts else default_fonts()
    proto = a.proto_rescue or os.environ.get('PROTO_RESCUE')
    proto = pathlib.Path(proto).resolve() if proto and not a.no_import and (pathlib.Path(proto) / 'api/_lib/docxToBlocks.ts').exists() else None
    pin = subprocess.run(['git', '-C', str(ROOT), 'ls-files', '-s', 'wordcraft'], capture_output=True, text=True).stdout.split()
    pin = pin[1][:8] if len(pin) > 1 else 'unknown'

    say('set %s: %d documents, %d flowing copies' % (cfg['name'], len(cfg['documents']), len(cfg.get('flow', []))))
    derived = derive(setdir, cfg, out)
    if a.editor == 'wordcraft' and not a.outputs:
        say('running the engine, sealed (fonts: %s)' % ('; '.join(str(f) for f in fonts) or 'none'))
        make_outputs(setdir, cfg, out, fonts)
    timings = json.loads((outputs / 'timings.json').read_text()) if (outputs / 'timings.json').exists() else {}
    say('measuring')
    P = measure_pages(setdir, cfg, outputs)
    K = measure_keys(cfg, derived, outputs)
    E = measure_edits(cfg, derived, outputs)
    I = measure_import(cfg, derived, outputs, proto) if not a.no_import else {'skipped': '--no-import'}
    X, via = measure_pixels(setdir, cfg, outputs)
    fr = json.loads((outputs / 'fonts.json').read_text(encoding='utf8')) if (outputs / 'fonts.json').exists() else {'folders': 'the editor\'s own'}
    S = summarize(cfg, a.editor, P, K, E, I, X, via, timings, fr, pin)
    spec = json.loads((HERE / 'spec.json').read_text(encoding='utf8'))
    is_private = cfg['name'] == spec['set']
    S['table'] = table(S, spec, is_private)
    S['detail'] = {'pages': P, 'keys': K, 'edits': E, 'import': I, 'pixels': X, 'timings': timings}
    (out / 'score.json').write_text(json.dumps(S, indent=1), encoding='utf8')
    print()
    print(S['table'])
    print()
    print('score.json:', out / 'score.json')
    if FAILURES:
        say('%d engine runs failed; the table above is incomplete' % len(FAILURES)); sys.exit(1)

if __name__ == '__main__':
    main()
