#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""docx_deep_extract.py — deep-parse .docx files for document-set analysis.

Parses identity from the file name (code / family / revision / date /
language / title — glued names like "PR.01DocumentControlProcedure-R6&..."
included), then reads the package read-only: heading tree, tables,
headers/footers, core properties, tracked-change/comment/macro flags, and
document codes cross-referenced in the text. Falls back to raw-XML parsing
when python-docx cannot open the file; classifies EMPTY / NOT_ZIP precisely.

Usage:
  docx_deep_extract.py FILE [FILE ...] [--out out.json] [--no-table-text]
  docx_deep_extract.py DIRECTORY        # recursive *.docx (cap 2000)

Output: JSON (object for one file, list for many). Never writes to inputs.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import zipfile

TEXT_CAP = 300_000
TABLES_CAP, ROWS_CAP, COLS_CAP, CELL_CAP = 80, 40, 20, 300

# --------------------------------------------------------------- identity --
CODE_RE = re.compile(r'^((?i:KEK|OEK|QM|PR|TL|TB|LS|FR|DD|SD|PG))[._\- ]?'
                     r'(\d{2,3}(?:[a-z](?![A-Za-z]))?)(?![0-9])')
REV_PATTERNS = [
    re.compile(r'[Rr]ev(?:izyon)?[ ._&\-]*(\d{1,2}(?:[.,]\d)?)'),
    re.compile(r'[\s._&\-(]R ?(\d{1,2}(?:[.,]\d)?)(?![\d.,])'),
    re.compile(r'[\s._&\-]REV[ ._&\-]*(\d{1,2}(?:[.,]\d)?)'),
    re.compile(r'(?<=[A-Za-zÇĞİÖŞÜçğıöşü])[Rr](?:ev)? ?0*(\d{1,2})(?=[&_\-.]|[0-9]{2}[._])'),
]
DATE_RE = re.compile(r'(\d{1,2})[._](\d{1,2})[._](\d{4})(?![0-9])')
DATE2_RE = re.compile(r'(\d{1,2})[._](\d{1,2})[._](\d{2})(?![0-9])')
TR_HINTS = ('prosedür', 'prosedur', 'talimat', 'formu', 'listesi', 'tablosu',
            'sözleşme', 'sozlesme', 'beyan', 'görev', 'gorev', 'eğitim',
            'egitim', 'belgelendirme', 'başvuru', 'basvuru', 'uygunsuzluk',
            'politika', 'müşteri', 'musteri', 'denetim', 'denetçi',
            'sertifiker', 'sertifikası', 'sertifikasi', 'tarafsızlık',
            'gizlilik', 'kalite', 'sınav', 'sinav', 'el kitabı')
EN_HINTS = ('procedure', 'instruction', 'manual', 'certificate', 'report',
            'policy', 'agreement', 'contract', 'application', 'checklist',
            'questionnaire', 'declaration', 'guide', 'training', 'exam',
            'answer', 'job description', 'audit', 'certification', 'offer',
            'review', 'plan', 'minutes', 'invitation')
REF_RE = re.compile(r'\b((?:KEK|OEK|QM|PR|TL|TB|LS|FR|DD|SD|PG)[._ ]?\d{2,3}[a-z]?)(?![0-9])')
EXT_REF_RE = re.compile(r'\b((?:ASR|CCS|GRS|OCS|RCS)[-_.]?\d{2,3})(?![0-9])')


def norm_ref(code):
    return re.sub(r'[-_. ]', '', code).upper()


def parse_identity(basename):
    code = prefix = None
    m = CODE_RE.match(basename)
    if m:
        prefix = m.group(1).upper()
        code = '%s.%s' % (prefix, m.group(2).lower())
    rev_raw = rev = None
    cands = []
    for pat in REV_PATTERNS:
        for mm in pat.finditer(basename):
            v = mm.group(1).replace(',', '.')
            try:
                f = float(v)
            except ValueError:
                continue
            cands.append((mm.start(), mm.group(0).strip(), v, f))
    if cands:
        cands.sort(key=lambda t: t[0])
        _pos, rev_raw, _v, rev = cands[-1]
        rev_raw = rev_raw.strip(' ._&-(,')
    eff = None
    dates = []
    for mm in DATE_RE.finditer(basename):
        d, mo, y = int(mm.group(1)), int(mm.group(2)), int(mm.group(3))
        if 1 <= d <= 31 and 1 <= mo <= 12:
            dates.append((mm.start(), '%04d-%02d-%02d' % (y, mo, d)))
    if not dates:
        for mm in DATE2_RE.finditer(basename):
            d, mo, y = int(mm.group(1)), int(mm.group(2)), int(mm.group(3))
            if 1 <= d <= 31 and 1 <= mo <= 12:
                dates.append((mm.start(), '%04d-%02d-%02d' % (2000 + y, mo, d)))
    if dates:
        dates.sort(key=lambda t: t[0])
        eff = dates[-1][1]
    lang = parse_lang(basename)
    title = clean_title(basename, code, rev_raw)
    fam = family_of(prefix, title.lower()) if prefix else None
    return {'code': code, 'prefix': prefix, 'family': fam, 'rev': rev,
            'rev_raw': rev_raw, 'eff_date': eff, 'lang': lang, 'title': title}


def parse_lang(name):
    up = name.upper()
    if 'TR-EN' in up or 'TR/EN' in up or 'EN-TR' in up:
        return 'BOTH'
    low = name.lower()
    tr = sum(1 for h in TR_HINTS if h in low)
    en = 0
    for h in EN_HINTS:
        ok = re.search(r'(?<![a-zçğıöşü])' + re.escape(h) + r'(?![a-zçğıöşü])', low)
        if ok or (len(h) >= 7 and h in low):
            en += 1
    return 'TR' if tr > en else ('EN' if en > tr else 'UNKNOWN')


def clean_title(name, code=None, rev_raw=None):
    t = name
    if code:
        p, num = code.split('.', 1)
        t = re.sub(re.escape(p) + r'[._\- ]?' + re.escape(num), ' ', t,
                   count=1, flags=re.I)
    if rev_raw:
        t = t.replace(rev_raw, ' ')
    t = DATE_RE.sub(' ', t)
    t = DATE2_RE.sub(' ', t)
    t = re.sub(r'[\s._\-&()\[\]]+', ' ', t)
    return re.sub(r'\s{2,}', ' ', t).strip(' .-_')


def family_of(prefix, tl):
    if prefix == 'PR':
        return 'PROCEDURE'
    if prefix == 'TL':
        return 'INSTRUCTION'
    if prefix == 'TB':
        return 'TABLE'
    if prefix == 'LS':
        return 'LIST'
    if prefix == 'FR':
        return 'FORM'
    if prefix in ('QM', 'KEK', 'OEK'):
        return 'MANUAL'
    if prefix == 'PG':
        return 'PROGRAM_DOC'
    if prefix in ('DD', 'SD'):
        if 'görev tanım' in tl or 'gorev tanim' in tl or 'job desc' in tl:
            return 'JOB_DESCRIPTION'
        if 'proses' in tl or 'process' in tl:
            return 'PROCESS'
        if 'politika' in tl or 'policy' in tl or 'beyan' in tl or 'commitment' in tl:
            return 'POLICY'
        return 'SUPPORT_DOC'
    return 'OTHER'


def find_refs(text, self_code=None):
    text = text[:TEXT_CAP]
    seen = {}
    disp = {}
    total = 0
    skip = norm_ref(self_code) if self_code else None
    for pat in (REF_RE, EXT_REF_RE):
        for m in pat.finditer(text):
            raw = m.group(1)
            key = norm_ref(raw)
            if key == skip:
                continue
            total += 1
            seen[key] = seen.get(key, 0) + 1
            disp.setdefault(key, raw)
    return {'codes': sorted(disp.values())[:200], 'distinct': len(seen),
            'mention_count': total}


NUM_HEAD_RE = re.compile(r'^(\d{1,2}(?:\.\d{1,2}){0,3})[.)]?\s+(\S(?:.{0,110}\S)?)$')


def guess_numbered_headings(paras, cap=200):
    """Heuristic outline for documents with no heading styles: '1.2 Text' lines."""
    out = []
    for t in paras:
        if len(t) > 112:
            continue
        m = NUM_HEAD_RE.match(t)
        if m:
            out.append({'level': m.group(1).count('.') + 1, 'text': t[:112]})
            if len(out) >= cap:
                break
    return out


# ------------------------------------------------------------ path helper --
def path_candidates(p):
    seen = set()

    def push(x):
        if x and x not in seen:
            seen.add(x)
            return True
        return False

    if push(p):
        yield p
    if os.name != 'nt':
        return
    q = p
    if q.startswith('\\\\') and not q.startswith('\\\\?\\'):
        q2 = '\\\\?\\UNC\\' + q[2:]
        if push(q2):
            yield q2
    elif re.match(r'^[A-Za-z]:', q):
        q2 = '\\\\?\\' + q
        if push(q2):
            yield q2
    if q.startswith('//'):
        b = '\\\\' + q[2:].replace('/', '\\')
        if push(b):
            yield b
        b2 = '\\\\?\\UNC\\' + b[2:]
        if push(b2):
            yield b2


# ---------------------------------------------------------------- readers --
def _zip_facts(z, names, rec):
    rec['zip_parts'] = len(names)
    rec['has_comments'] = 'word/comments.xml' in names
    rec['has_macros'] = 'word/vbaProject.bin' in names
    rec['media_files'] = sum(1 for n in names if n.startswith('word/media/'))
    docxml = z.read('word/document.xml') if 'word/document.xml' in names else b''
    rec['tracked_ins'] = docxml.count(b'<w:ins ')
    rec['tracked_del'] = docxml.count(b'<w:del ')
    rec['has_tracked_changes'] = bool(rec['tracked_ins'] or rec['tracked_del'])
    rec['hyperlink_count'] = docxml.count(b'<w:hyperlink')
    return docxml


def _sample_tables_pydocx(d, rec, table_text=True):
    tables = []
    for ti, tbl in enumerate(d.tables[:TABLES_CAP]):
        try:
            cells = []
            if table_text:
                cells = [[(c.text or '')[:CELL_CAP] for c in r.cells[:COLS_CAP]]
                         for r in tbl.rows[:ROWS_CAP]]
            tables.append({'index': ti, 'rows': len(tbl.rows),
                           'cols': len(tbl.columns),
                           'style': (tbl.style.name if tbl.style else None),
                           'cells_sample': cells})
        except Exception as ex:
            tables.append({'index': ti, 'error': str(ex)})
    rec['table_count'] = len(d.tables)
    rec['tables'] = tables


def _extract_pydocx(path, rec, table_text=True):
    import docx
    z = zipfile.ZipFile(path)
    names = z.namelist()
    _zip_facts(z, names, rec)
    z.close()
    d = docx.Document(path)
    cp = d.core_properties
    rec['doc_props'] = {
        'title': cp.title, 'author': cp.author,
        'last_modified_by': cp.last_modified_by,
        'created': str(cp.created) if cp.created else None,
        'modified': str(cp.modified) if cp.modified else None,
        'revision': cp.revision, 'category': cp.category,
        'comments': cp.comments, 'subject': cp.subject, 'keywords': cp.keywords,
    }
    heads, paras = [], []
    for p in d.paragraphs:
        try:
            t = (p.text or '').strip()
        except Exception:
            t = ''
        if not t:
            continue
        paras.append(t)
        try:
            st = p.style.name or '' if p.style is not None else ''
            sid = p.style.style_id or '' if p.style is not None else ''
        except Exception:
            st, sid = '', ''
        m = (re.match(r'(?i)heading\s*(\d+)', st) or re.match(r'(?i)Heading(\d+)', sid)
             or re.match(r'(?i)ba[şs]l[ıi]k\s*(\d+)', st))
        if m:
            heads.append({'level': int(m.group(1)), 'text': t[:300], 'style': st or sid})
        elif st.lower() in ('title', 'subtitle') or sid in ('Title', 'Subtitle'):
            heads.append({'level': 0, 'text': t[:300], 'style': st or sid})
    rec['headings'] = heads[:400]
    rec['heading_count'] = len(heads)
    rec['paragraph_count'] = len(paras)
    rec['numbered_outline_guess'] = guess_numbered_headings(paras)
    _sample_tables_pydocx(d, rec, table_text)
    hf = []
    try:
        for si, s in enumerate(d.sections[:6]):
            for pname, part in (('header', s.header), ('footer', s.footer)):
                try:
                    tt = ' | '.join((p.text or '').strip()
                                    for p in part.paragraphs if (p.text or '').strip())
                    if tt:
                        hf.append({'section': si, 'part': pname, 'text': tt[:400]})
                except Exception:
                    pass
    except Exception:
        pass
    rec['header_footer'] = hf
    text = '\n'.join(paras)
    if len(text) > TEXT_CAP:
        text = text[:TEXT_CAP] + '\n[[TRUNCATED]]'
    rec['text_len'] = len(text)
    reftext = text
    if table_text:
        reftext += '\n' + '\n'.join(c for t in rec['tables']
                                    for row in (t.get('cells_sample') or [])
                                    for c in row)
    rec['refs'] = find_refs(reftext, rec['identity']['code'])
    rec['status'] = 'OK' if (text.strip() or rec['table_count']) else 'EMPTY'
    return rec


def _extract_raw_xml(path, rec, table_text=True):
    import xml.etree.ElementTree as ET
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    z = zipfile.ZipFile(path)
    names = z.namelist()
    docxml = _zip_facts(z, names, rec)
    style_map = {}
    if 'word/styles.xml' in names:
        try:
            sroot = ET.fromstring(z.read('word/styles.xml'))
            for st in sroot.iter(W + 'style'):
                sid = st.get(W + 'styleId') or st.get('styleId')
                nm_el = st.find(W + 'name')
                nm = nm_el.get(W + 'val') if nm_el is not None else None
                if sid and nm:
                    style_map[sid] = nm
        except Exception:
            style_map = {}
    core = None
    if 'docProps/core.xml' in names:
        try:
            croot = ET.fromstring(z.read('docProps/core.xml'))
            core = {
                'title': croot.findtext('{http://purl.org/dc/elements/1.1/}title'),
                'author': croot.findtext('{http://purl.org/dc/elements/1.1/}creator'),
                'last_modified_by': croot.findtext(
                    '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy'),
                'created': croot.findtext('{http://purl.org/dc/terms/}created'),
                'modified': croot.findtext('{http://purl.org/dc/terms/}modified'),
            }
        except Exception:
            core = None
    z.close()
    rec['doc_props'] = core
    root = ET.fromstring(docxml)
    heads, paras = [], []
    for p in root.iter(W + 'p'):
        t = ''.join(p.itertext()).strip()
        if not t:
            continue
        paras.append(t)
        ps = p.find(W + 'pPr/' + W + 'pStyle')
        style = (ps.get(W + 'val') or ps.get('val')) if ps is not None else None
        if style:
            lvl = None
            for c in (style, style_map.get(style) or ''):
                if not c:
                    continue
                m = (re.match(r'(?i)heading\s*(\d+)$', c)
                     or re.match(r'(?i)ba[şs]l[ıi]k\s*(\d+)$', c)
                     or re.match(r'^(\d)$', c))
                if m:
                    lvl = int(m.group(1))
                    break
            if lvl is not None:
                heads.append({'level': lvl, 'text': t[:300], 'style': style})
            elif style in ('Title', 'Subtitle') or (
                    (style_map.get(style) or '').lower() in ('title', 'subtitle')):
                heads.append({'level': 0, 'text': t[:300], 'style': style})
    rec['headings'] = heads[:400]
    rec['heading_count'] = len(heads)
    rec['paragraph_count'] = len(paras)
    rec['numbered_outline_guess'] = guess_numbered_headings(paras)
    tables = []
    tbls = list(root.iter(W + 'tbl'))
    for ti, tbl in enumerate(tbls[:TABLES_CAP]):
        rows = []
        for r in tbl.iter(W + 'tr'):
            cells = [''.join(c.itertext()).strip()[:CELL_CAP]
                     for c in list(r.iter(W + 'tc'))[:COLS_CAP]]
            rows.append(cells)
            if len(rows) >= ROWS_CAP:
                break
        tables.append({'index': ti, 'rows': len(rows),
                       'cells_sample': rows if table_text else []})
    rec['table_count'] = len(tbls)
    rec['tables'] = tables
    rec['header_footer'] = []
    text = '\n'.join(paras)
    if len(text) > TEXT_CAP:
        text = text[:TEXT_CAP] + '\n[[TRUNCATED]]'
    rec['text_len'] = len(text)
    reftext = text
    if table_text:
        reftext += '\n' + '\n'.join(c for t in tables
                                    for row in (t.get('cells_sample') or [])
                                    for c in row)
    rec['refs'] = find_refs(reftext, rec['identity']['code'])
    rec['status'] = 'OK' if (text.strip() or rec['table_count']) else 'EMPTY'
    return rec


# ------------------------------------------------------------------- main --
def extract(path, table_text=True):
    rec = {'path': path, 'identity': parse_identity(os.path.basename(path))}
    try:
        rec['size'] = os.stat(path).st_size
        if rec['size'] == 0:
            rec['status'] = 'EMPTY'
            return rec
    except OSError:
        rec['size'] = None
    last_err = None
    for cand in path_candidates(path):
        try:
            return _extract_pydocx(cand, rec, table_text)
        except zipfile.BadZipFile as ex:
            rec['status'] = 'NOT_ZIP'
            rec['error'] = 'BadZipFile: %s' % ex
            return rec
        except ImportError:
            last_err = 'python-docx not installed'
            break
        except Exception as ex:
            last_err = '%s: %s' % (type(ex).__name__, ex)
            rec['fallback'] = 'raw_xml'
            try:
                return _extract_raw_xml(cand, rec, table_text)
            except zipfile.BadZipFile:
                rec['status'] = 'NOT_ZIP'
                rec['error'] = 'BadZipFile (raw fallback)'
                return rec
            except Exception as ex2:
                last_err = '%s | raw-xml: %s: %s' % (last_err, type(ex2).__name__, ex2)
                continue
    if last_err == 'python-docx not installed':
        rec['fallback'] = 'raw_xml'
        for cand in path_candidates(path):
            try:
                return _extract_raw_xml(cand, rec, table_text)
            except Exception as ex2:
                last_err = 'raw-xml: %s: %s' % (type(ex2).__name__, ex2)
                continue
    rec['status'] = 'FAILED'
    rec['error'] = last_err or 'unreadable (tried long-path variants)'
    return rec


def main():
    ap = argparse.ArgumentParser(
        description='Deep-parse .docx for document-set analysis.')
    ap.add_argument('paths', nargs='+', help='file(s) or directory (recursive *.docx)')
    ap.add_argument('--out', help='write JSON here')
    ap.add_argument('--quiet', action='store_true', help='do not print JSON to stdout')
    ap.add_argument('--no-table-text', action='store_true',
                    help='table shapes only - no cell text (smaller output)')
    args = ap.parse_args()
    files = []
    for p in args.paths:
        if os.path.isdir(p):
            for base, _dirs, fs in os.walk(p):
                for f in fs:
                    if f.lower().endswith('.docx'):
                        files.append(os.path.join(base, f))
            files = files[:2000]
        else:
            files.append(p)
    out = [extract(f, table_text=not args.no_table_text) for f in files]
    data = out[0] if len(out) == 1 else out
    txt = json.dumps(data, ensure_ascii=False, indent=1)
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as fh:
            fh.write(txt)
    if not args.quiet:
        print(txt)


if __name__ == '__main__':
    main()
