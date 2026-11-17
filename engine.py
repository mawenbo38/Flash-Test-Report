"""Index and filter LONGi flash report OOXML without recreating the formatting."""
from __future__ import annotations
import copy
import io
import re
import sqlite3
import zipfile
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from lxml import etree as ET

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'m': MAIN, 'r': REL, 'p': PKG, 'c': CT}
V = '{%s}v' % MAIN
DATA_START = 7
DATA_END = 726  # This file family has a styled trailing row 727; determine dynamically when importing.


def normalize_sn(value):
    return str(value or '').strip().upper().replace('\u200b', '')


def get_db(path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(path, timeout=30)
    con.execute('PRAGMA foreign_keys=ON')
    con.execute('''CREATE TABLE IF NOT EXISTS sources (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, uploaded_at TEXT NOT NULL,
        content BLOB NOT NULL, rows_count INTEGER NOT NULL DEFAULT 0)''')
    con.execute('''CREATE TABLE IF NOT EXISTS serials (
        source_id INTEGER NOT NULL, sheet_name TEXT NOT NULL, sheet_path TEXT NOT NULL,
        row_num INTEGER NOT NULL, sn TEXT NOT NULL,
        PRIMARY KEY (source_id, sheet_path, row_num),
        FOREIGN KEY (source_id) REFERENCES sources(id) ON DELETE CASCADE)''')
    con.execute('CREATE INDEX IF NOT EXISTS ix_serials_sn ON serials(sn)')
    con.commit()
    return con


def read_shared(z):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    return [''.join(node.itertext()) for node in root]


def sheet_paths(z):
    """Use workbook relationships instead of assuming sheet names match filenames."""
    book = ET.fromstring(z.read('xl/workbook.xml'))
    relroot = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    relmap = {e.get('Id'): e.get('Target') for e in relroot}
    result = []
    for node in book.xpath('./m:sheets/m:sheet', namespaces=NS):
        target = relmap.get(node.get('{%s}id' % REL), '')
        if not target:
            continue
        if target.startswith('/'):
            part = target.lstrip('/')
        elif target.startswith('xl/'):
            part = target
        else:
            part = 'xl/' + target
        part = str(Path(part).as_posix())
        if part in z.namelist():
            result.append((node.get('name'), part))
    return result


def cell_value(cell, shared):
    if cell is None:
        return None
    typ = cell.get('t')
    if typ == 'inlineStr':
        return ''.join(cell.xpath('./m:is//m:t/text()', namespaces=NS))
    v = cell.find(V)
    if v is None:
        return None
    if typ == 's':
        return shared[int(v.text)]
    return v.text


def row_cells(row):
    return {re.match(r'([A-Z]+)', c.get('r', '')).group(1): c
            for c in row.findall('m:c', namespaces=NS) if re.match(r'([A-Z]+)', c.get('r', ''))}


def extract_rows(z, sheet_path, shared):
    root = ET.fromstring(z.read(sheet_path))
    out = []
    for row in root.xpath('./m:sheetData/m:row', namespaces=NS):
        index = int(row.get('r'))
        if index < DATA_START:
            continue
        cells = row_cells(row)
        sn = normalize_sn(cell_value(cells.get('A'), shared))
        # Skip styled/empty footer rows and protect against other export footer labels.
        if not sn or not re.match(r'^[A-Z0-9]{12,}$', sn):
            continue
        if not any(cell_value(cells.get(col), shared) is not None for col in ('B', 'C', 'D', 'E', 'F')):
            continue
        out.append((index, sn))
    return out


def index_upload(con, filename, data):
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise ValueError('ไฟล์ไม่ใช่ .xlsx ที่ถูกต้อง')
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        shared = read_shared(z)
        sheets = sheet_paths(z)
        if not sheets:
            raise ValueError('ไม่พบ worksheet ในไฟล์')
        entries = []
        for sheet_name, path in sheets:
            entries.extend((sheet_name, path, index, sn) for index, sn in extract_rows(z, path, shared))
    if not entries:
        raise ValueError('ไม่พบ Serial Number ตามโครงสร้างตัวอย่าง (คอลัมน์ A เริ่มแถว 7)')
    with con:
        cur = con.execute('INSERT INTO sources(name,uploaded_at,content,rows_count) VALUES (?,?,?,?)',
                          (filename, datetime.now(timezone.utc).isoformat(), data, len(entries)))
        source_id = cur.lastrowid
        con.executemany('INSERT INTO serials(source_id,sheet_name,sheet_path,row_num,sn) VALUES (?,?,?,?,?)',
                        ((source_id, *entry) for entry in entries))
    return source_id, len(entries), len(sheets)


def parse_sn_text(text):
    """Accept one-per-line, comma-, semicolon-, tab-, or space-separated SN."""
    return list(dict.fromkeys(normalize_sn(x) for x in re.split(r'[,;\s]+', text or '') if normalize_sn(x)))


def query_matches(con, sns, source_ids=None):
    if not sns:
        return []
    matches = []
    for offset in range(0, len(sns), 400):
        batch = sns[offset:offset + 400]
        q = 'SELECT s.source_id, s.sheet_name, s.sheet_path, s.row_num, s.sn, f.name FROM serials s JOIN sources f ON s.source_id=f.id WHERE s.sn IN (%s)' % ','.join('?' * len(batch))
        params = list(batch)
        if source_ids is not None:
            if not source_ids:
                return []
            q += ' AND s.source_id IN (%s)' % ','.join('?' * len(source_ids))
            params.extend(source_ids)
        matches.extend(con.execute(q, params).fetchall())
    return matches


def replace_numeric(cell, value):
    if cell is None:
        return
    cell.attrib.pop('t', None)
    for sub in list(cell):
        if sub.tag in (V, '{%s}is' % MAIN, '{%s}f' % MAIN):
            cell.remove(sub)
    ET.SubElement(cell, V).text = str(value)


def compact_sheet(source_xml, selected_row_nums, shared):
    root = ET.fromstring(source_xml)
    data = root.find('m:sheetData', NS)
    if data is None:
        raise ValueError('Invalid worksheet: missing sheetData')
    chosen = set(selected_row_nums)
    original = list(data.findall('m:row', NS))
    header = [copy.deepcopy(r) for r in original if int(r.get('r')) < DATA_START]
    rows = [copy.deepcopy(r) for r in original if int(r.get('r')) in chosen]
    # Do not retain arbitrary trailing rows: output must contain only requested serials.
    for child in list(data):
        data.remove(child)
    for r in header:
        data.append(r)
    numeric = defaultdict(list)
    for i, row in enumerate(rows, DATA_START):
        row.set('r', str(i))
        for c in row.findall('m:c', NS):
            ref = c.get('r', '')
            m = re.match(r'^([A-Z]+)\d+$', ref)
            if m:
                c.set('r', m.group(1) + str(i))
        cells = row_cells(row)
        for col in 'BCDEFG':
            v = cell_value(cells.get(col), shared)
            try:
                numeric[col].append(float(v))
            except (ValueError, TypeError):
                pass
        data.append(row)
    # Original rows 3-5 summarize *all* original rows. Recalculate for selected rows.
    header_by_num = {int(r.get('r')): r for r in header}
    for column, values in numeric.items():
        if not values:
            continue
        for rnum, aggregate in ((3, sum(values) / len(values)), (4, min(values)), (5, max(values))):
            replace_numeric(row_cells(header_by_num[rnum]).get(column), aggregate)
    dim = root.find('m:dimension', NS)
    if dim is not None:
        dim.set('ref', 'A1:J%d' % max(6, 6 + len(rows)))
    # Ensure print area / print settings aren't tied to old 727-row document.
    for node in root.xpath('./m:ignoredErrors/m:ignoredError', namespaces=NS):
        if node.get('sqref'):
            node.set('sqref', 'A1:J%d' % max(6, 6 + len(rows)))
    return ET.tostring(root, encoding='UTF-8', xml_declaration=True)


def build_filtered_report(original_bytes, matches):
    """Filter by exact indexed row numbers; keep matching sheets only and XML formatting intact."""
    selected = defaultdict(set)
    for source_id, sheet_name, sheet_path, row_num, sn, filename in matches:
        selected[sheet_path].add(row_num)
    if not selected:
        raise ValueError('ไม่มีข้อมูลที่เลือก')
    src = io.BytesIO(original_bytes)
    output = io.BytesIO()
    with zipfile.ZipFile(src, 'r') as zi:
        shared = read_shared(zi)
        names = sheet_paths(zi)
        keep_paths = set(selected)
        wb = ET.fromstring(zi.read('xl/workbook.xml'))
        relroot = ET.fromstring(zi.read('xl/_rels/workbook.xml.rels'))
        relmap = {e.get('Id'): e for e in relroot}
        by_name = dict(names)
        kept_rids = set()
        sheets_el = wb.find('m:sheets', NS)
        for sheet in list(sheets_el):
            rid = sheet.get('{%s}id' % REL)
            if by_name.get(sheet.get('name')) in keep_paths:
                kept_rids.add(rid)
            else:
                sheets_el.remove(sheet)
        for el in list(relroot):
            if 'worksheet' in el.get('Type', '') and el.get('Id') not in kept_rids:
                relroot.remove(el)
        ct = ET.fromstring(zi.read('[Content_Types].xml'))
        for el in list(ct):
            if el.get('PartName', '').lstrip('/') in (set(path for _, path in names) - keep_paths):
                ct.remove(el)
        updates = {
            'xl/workbook.xml': ET.tostring(wb, encoding='UTF-8', xml_declaration=True),
            'xl/_rels/workbook.xml.rels': ET.tostring(relroot, encoding='UTF-8', xml_declaration=True),
            '[Content_Types].xml': ET.tostring(ct, encoding='UTF-8', xml_declaration=True),
        }
        for path in keep_paths:
            updates[path] = compact_sheet(zi.read(path), selected[path], shared)
        with zipfile.ZipFile(output, 'w') as zo:
            for info in zi.infolist():
                if info.filename in (set(path for _, path in names) - keep_paths):
                    continue
                zo.writestr(info, updates.get(info.filename, zi.read(info.filename)))
    return output.getvalue()


def extract_sns_from_upload(file_name, raw):
    suffix = Path(file_name).suffix.lower()
    if suffix in ('.txt', '.csv'):
        return parse_sn_text(raw.decode('utf-8-sig', errors='replace'))
    if suffix != '.xlsx':
        raise ValueError('รองรับไฟล์ S/N แบบ .xlsx, .csv หรือ .txt')
    all_values = []
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        shared = read_shared(z)
        for name, path in sheet_paths(z):
            root = ET.fromstring(z.read(path))
            for row in root.xpath('./m:sheetData/m:row', namespaces=NS):
                cells = row_cells(row)
                # S/N list in first column; ignore header e.g. Serial Number.
                val = normalize_sn(cell_value(cells.get('A'), shared))
                if val and re.fullmatch(r'[A-Z0-9_-]{8,}', val) and val not in ('SERIALNUMBER', 'SERIAL_NUMBER'):
                    all_values.append(val)
    return list(dict.fromkeys(all_values))
