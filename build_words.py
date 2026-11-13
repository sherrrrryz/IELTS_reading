import sys, glob, os, zipfile, re, json
import xml.etree.ElementTree as ET
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def col_idx(ref):
    s = re.match(r'[A-Z]+', ref).group(0); n = 0
    for ch in s: n = n*26 + ord(ch)-64
    return n-1
def read(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    ss = []
    if 'xl/sharedStrings.xml' in names:
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', NS):
            ss.append(''.join(t.text or '' for t in si.iter('{%s}t' % NS['m'])))
    sheets = sorted(n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
    out = []
    for sn in sheets:
        rows = []
        for row in ET.fromstring(z.read(sn)).iter('{%s}row' % NS['m']):
            cells = {}
            for c in row.findall('m:c', NS):
                t = c.get('t'); v = c.find('m:v', NS)
                if t == 's' and v is not None: val = ss[int(v.text)]
                elif t == 'inlineStr': val = ''.join(x.text or '' for x in c.iter('{%s}t' % NS['m']))
                else: val = v.text if v is not None else None
                if val is not None: cells[col_idx(c.get('r'))] = val
            if cells: rows.append([cells.get(i) for i in range(max(cells)+1)])
        out.append((sn, rows))
    return out
def files(d):
    return sorted(glob.glob(os.path.join(d, '*.xlsx')), key=lambda p: int(p.rsplit('_', 1)[1].split('.')[0]))

# 用法: python3 -I build_words.py <表格文件夹> <输出 words.js>
if __name__ == '__main__':
    src, dst = sys.argv[1], sys.argv[2]
    # 同目录下的 exclude.txt：不放进单词卡的词
    ex_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'exclude.txt')
    exclude = set()
    if os.path.exists(ex_path):
        with open(ex_path, encoding='utf-8') as fh:
            exclude = {ln.strip().lower() for ln in fh if ln.strip() and not ln.startswith('#')}
    words, skipped, hit = [], 0, set()
    for f in files(src):
        n = int(f.rsplit('_', 1)[1].split('.')[0])
        for _, rows in read(f):
            for r in rows[1:]:
                r = [(x or '').strip() for x in r] + [''] * 5
                if not r[0]:
                    continue
                if r[0].lower() in exclude:
                    skipped += 1
                    hit.add(r[0].lower())
                    continue
                d = '\n'.join(s.strip().rstrip('；;') for s in r[2].split('\n') if s.strip())
                words.append({'l': n, 'w': r[0], 'p': r[1], 'd': d, 'e': r[3], 't': r[4]})
    with open(dst, 'w', encoding='utf-8') as fh:
        fh.write('window.WORDS = ' + json.dumps(words, ensure_ascii=False, separators=(',', ':')) + ';\n')
    print(len(words), 'words ->', dst, '| skipped', skipped, 'entries for', len(hit), 'excluded words')
    if exclude - hit:
        print('exclude.txt 里没匹配到的词:', sorted(exclude - hit))
