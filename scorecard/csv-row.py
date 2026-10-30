"""One line of scores.csv from a score.json:  csv-row.py score.json [--header]"""
import json, sys
COLS = ['date', 'pin', 'patch_set', 'set', 'editor', 'same_word', 'same_word_of', 'within_line', 'page_count_same', 'documents',
        'flow_same_word', 'words_kept', 'pages_identical', 'pages_identical_of', 'pixels_via', 'keys_kept', 'keys_of',
        'paraids_kept', 'paraids_of', 'edit_cases_ok', 'edit_cases', 'pdf_seconds', 'fonts_microsoft', 'fonts_stand_ins']
def row(S):
    f = S.get('fonts', {})
    flags = lambda d: '+'.join(k for k, v in (d or {}).items() if v) or 'none'
    return [S['date'], S['pin'], '0', S['set'], S['editor'], S['same_word'][0], S['same_word'][1], S['within_line'][0],
            S['page_count_same'][0], S['documents'],
            ';'.join('%s:%d/%d' % (d, r[0], r[2]) for d, r in S['flow'].items() if r),
            '' if S['save'] is None else 'yes' if S['save']['all_kept'] else 'no',
            S['pixels'][0], S['pixels'][1], S.get('pixels_via') or '', S['keys'][0], S['keys'][1], S['paraids'][0], S['paraids'][1],
            '' if S['edits'] is None else S['edits']['passed'], '' if S['edits'] is None else S['edits']['of'],
            '' if S['pdf_seconds']['seconds'] is None else S['pdf_seconds']['seconds'],
            flags(f.get('microsoft')) if isinstance(f, dict) else '', flags(f.get('stand_ins')) if isinstance(f, dict) else '']
if __name__ == '__main__':
    if '--header' in sys.argv: print(','.join(COLS))
    else: print(','.join(str(x) for x in row(json.load(open(sys.argv[1], encoding='utf8')))))
