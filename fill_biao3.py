"""
Fill subject content (cells 4-12) into the 6 existing 表3 DOCX files.
Rows filled: 4,5,6,8,9,10 in Table 1 (semester 1) and Table 3 (semester 2).

Column mapping (0-indexed in 17-cell data rows):
  4=國語文  5=本土語文  6=英語文  7=數學
  8=社會    9=自然      10=藝術   11=綜合  12=健康體育

Notes:
  - Grades 1-2: KX table has different structure (col8=生活課程, col9=健體, col10=彈性)
    → col9(健體) remapped to col12; col10 cleared; 自然/藝術/綜合 left empty (合科)
  - Grades 5-6: 數學 from 翰林 (not 南一)
  - parse_weekly_file scans ALL tables filtered by 週次 header (handles 翰林 structure)
"""
import sys, re, os, shutil, subprocess
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SKILL_BASE = (r"C:\Users\admin\AppData\Roaming\Claude\local-agent-mode-sessions"
              r"\skills-plugin\f044c158-29da-435c-ae59-7ad9c15f098a"
              r"\7ba8d970-f36e-4ee9-afa9-35fb2a74155e\skills\docx")
UNPACK = os.path.join(SKILL_BASE, r"scripts\office\unpack.py")
PACK   = os.path.join(SKILL_BASE, r"scripts\office\pack.py")

BASE    = r'D:\教學組\114教務\教學\課程計畫\claude 操作\三大書局總表資料'
KX_DIR  = BASE + r'\13臺中市-康軒\13臺中市'
HK_DIR  = BASE + r'\13臺中市康軒客語'
N1_DIR  = BASE + r'\01.課程計畫-南一\01.課程計畫\01橫式'
HL_DIR  = BASE + r'\115年翰林版課程計畫\01. 翰林版\1.課程計畫\01.標準版(橫式)'
OUT_DIR = r'D:\教學組\114教務\教學\課程計畫\claude 操作'
WORK    = r'C:\Users\admin\AppData\Local\Temp\fill_biao3_work'
TPL_DOCX = r'D:\教學組\114教務\教學\課程計畫\claude 操作\表3.docx'

GRADES = [
    (1, '一年級', '客', 'BA'),
    (2, '二年級', '閩', 'BA'),
    (3, '三年級', '客', 'WW'),
    (4, '四年級', '閩', 'WW'),
    (5, '五年級', '客', 'WW'),
    (6, '六年級', '閩', 'WW'),
]

CN_NUM = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,
          '十':10,'十一':11,'十二':12,'十三':13,'十四':14,'十五':15,
          '十六':16,'十七':17,'十八':18,'十九':19,'二十':20,'二十一':21,'二十二':22}

def parse_week_num(s):
    m = re.search(r'第([一二三四五六七八九十]+)週', s)
    if not m: return 99
    return CN_NUM.get(m.group(1), 99)

# ── XML helpers ─────────────────────────────────────────────────────────────

def get_tables(xml):
    """Return (start, end) for each TOP-LEVEL <w:tbl>."""
    tables = []
    pos = 0
    while pos < len(xml):
        s = xml.find('<w:tbl>', pos)
        if s < 0: break
        depth = 1; p = s + 7; found = False
        while depth > 0 and p < len(xml):
            o = xml.find('<w:tbl>', p); c = xml.find('</w:tbl>', p)
            if c < 0: p = len(xml); break
            if o >= 0 and o < c: depth += 1; p = o + 7
            else:
                depth -= 1; p = c + 8
                if depth == 0: tables.append((s, p)); pos = p; found = True; break
        if not found: pos = s + 7
    return tables

def get_rows(tbl_xml):
    rows = []; pos = 0
    while pos < len(tbl_xml):
        m = re.search(r'<w:tr[ >]', tbl_xml[pos:])
        if not m: break
        s = pos + m.start(); depth = 1; p = pos + m.end(); found = False
        while depth > 0 and p < len(tbl_xml):
            om = re.search(r'<w:tr[ >]', tbl_xml[p:]); c = tbl_xml.find('</w:tr>', p)
            if c < 0: p = len(tbl_xml); break
            o = p + om.start() if om else len(tbl_xml)
            if o < c: depth += 1; p = p + om.end()
            else:
                depth -= 1; p = c + 7
                if depth == 0: rows.append((s, p)); pos = p; found = True; break
        if not found: pos = pos + m.end()
    return rows

def get_cells(row_xml):
    cells = []; pos = 0
    while pos < len(row_xml):
        s = row_xml.find('<w:tc>', pos)
        if s < 0: break
        depth = 1; p = s + 6; found = False
        while depth > 0 and p < len(row_xml):
            o = row_xml.find('<w:tc>', p); c = row_xml.find('</w:tc>', p)
            if c < 0: p = len(row_xml); break
            if o >= 0 and o < c: depth += 1; p = o + 6
            else:
                depth -= 1; p = c + 7
                if depth == 0: cells.append((s, p)); pos = p; found = True; break
        if not found: pos = s + 6
    return cells

def get_cell_text(cell_xml):
    return ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', cell_xml))

def make_plain_para(text, rpr_xml=''):
    text = text[:MAX_CELL]
    t_attr = ' xml:space="preserve"' if text != text.strip() or ' ' in text else ''
    escaped = text.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    return (f'<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
            f'<w:r>{rpr_xml}<w:t{t_attr}>{escaped}</w:t></w:r></w:p>')

def replace_cell_content(row_xml, cell_idx, new_content_xml):
    cells = get_cells(row_xml)
    if cell_idx >= len(cells): return row_xml
    cs, ce = cells[cell_idx]
    cell_xml = row_xml[cs:ce]
    tcpr_end = cell_xml.find('</w:tcPr>')
    if tcpr_end >= 0:
        new_cell = cell_xml[:tcpr_end+len('</w:tcPr>')] + new_content_xml + '</w:tc>'
    else:
        new_cell = cell_xml[:cell_xml.index('>')+1] + new_content_xml + '</w:tc>'
    return row_xml[:cs] + new_cell + row_xml[ce:]

# ── Parse 康軒 integrated table ──────────────────────────────────────────────

def parse_kx_table(tbl_xml, row_indices, cell_indices):
    rows = get_rows(tbl_xml)
    result = {}
    for ri in row_indices:
        if ri >= len(rows): continue
        rs, re_ = rows[ri]
        row_xml = tbl_xml[rs:re_]
        cells = get_cells(row_xml)
        result[ri] = {}
        for ci in cell_indices:
            if ci >= len(cells): continue
            cs, ce = cells[ci]
            result[ri][ci] = get_cell_text(row_xml[cs:ce])
    return result

# ── Parse per-week table (南一 / 翰林 / 客語) ────────────────────────────────

ASSESS_KW = {'評量', '測驗', '口頭', '紙筆', '實作', '觀察', '討論', '表演'}

def detect_col_indices(header_row_texts, sample_data=None):
    topic_ci = unit_ci = teach_ci = assess_ci = week_ci = 0
    for i, h in enumerate(header_row_texts):
        if '週次' in h or '起訖週次' in h: week_ci = i
        if h in ('主題', '單元/主題名稱'): topic_ci = i
        if '單元名稱' in h and '主題' not in h: unit_ci = i
        if '教學活動重點' in h: teach_ci = i
        if '教學重點' in h and '教學活動' not in h: teach_ci = i
        if '評量方式' in h: assess_ci = i
    if sample_data:
        scores = {}
        for row_txts in sample_data:
            for i, t in enumerate(row_txts):
                if any(kw in t for kw in ASSESS_KW):
                    scores[i] = scores.get(i, 0) + 1
        if scores:
            best = max(scores, key=scores.get)
            if scores.get(best, 0) > scores.get(assess_ci, 0):
                assess_ci = best
    return week_ci, topic_ci, unit_ci, teach_ci, assess_ci

def parse_weekly_file(xml, split_at=10):
    """
    Parse a per-week source file.
    Scans ALL tables, only processes those with 週次 in header.
    Returns list of (period1, period2) tuples per qualifying table.
    """
    tables = get_tables(xml)
    periods = []
    for ts, te in tables:   # scan ALL tables; filter below by 週次 header
        tbl_xml = xml[ts:te]
        rows = get_rows(tbl_xml)
        if len(rows) < 3: continue
        hr = tbl_xml[rows[0][0]:rows[0][1]]
        hcells = get_cells(hr)
        header_texts = [''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>',
                        hr[cs:ce])) for cs, ce in hcells]
        # Only process tables that have a week column (skip 參考資料/學習目標 tables)
        if not any('週次' in h or '起訖週次' in h for h in header_texts):
            continue
        sample = []
        for ri in range(1, min(6, len(rows))):
            rxml = tbl_xml[rows[ri][0]:rows[ri][1]]
            rcells = get_cells(rxml)
            sample.append([''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>',
                           rxml[cs:ce])) for cs, ce in rcells])
        week_ci, topic_ci, unit_ci, _, assess_ci = detect_col_indices(header_texts, sample)

        p1 = {'units': [], 'assess': []}
        p2 = {'units': [], 'assess': []}
        for ri in range(1, len(rows)):
            rs, re_ = rows[ri]
            row_xml = tbl_xml[rs:re_]
            cells = get_cells(row_xml)
            txts = [''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>',
                    row_xml[cs:ce])) for cs, ce in cells]
            if not any(txts): continue
            week_str = txts[week_ci] if week_ci < len(txts) else ''
            wn = parse_week_num(week_str)
            period = p1 if wn <= split_at else p2
            topic = txts[topic_ci] if topic_ci < len(txts) else ''
            unit  = txts[unit_ci]  if unit_ci > 0 and unit_ci < len(txts) else ''
            combined = (topic + ('　' + unit if unit and unit != topic else '')).strip()
            if combined and combined not in period['units']:
                period['units'].append(combined)
            a = txts[assess_ci] if assess_ci < len(txts) else ''
            if a and a not in period['assess']:
                period['assess'].append(a)
        periods.append((p1, p2))
    return periods

MAX_CELL = 500

def aggregate_text(items, sep=''):
    return sep.join(items)

def period_to_texts(period):
    units = aggregate_text(period['units'])[:MAX_CELL]
    assess = aggregate_text(list(dict.fromkeys(period['assess'])))[:MAX_CELL]
    return {
        'units':  units,
        'teach':  units,
        'assess': assess,
    }

# ── Main ─────────────────────────────────────────────────────────────────────

os.makedirs(WORK, exist_ok=True)

for grade_num, grade_name, hoklo_type, eng_ver in GRADES:
    print(f'\n=== {grade_name} ===')
    grade_work = os.path.join(WORK, f'g{grade_num}')
    os.makedirs(grade_work, exist_ok=True)

    def unpack(src, dst_name):
        dst = os.path.join(grade_work, dst_name)
        if os.path.exists(dst): shutil.rmtree(dst)
        r = subprocess.run(['python', UNPACK, src, dst],
                           capture_output=True, text=True, encoding='utf-8', errors='replace')
        if r.returncode != 0:
            print(f'  Unpack error ({dst_name}): {r.stderr[:150]}')
            return None
        return dst

    def read_xml(unpacked_dir):
        p = os.path.join(unpacked_dir, 'word', 'document.xml')
        with open(p, encoding='utf-8') as f:
            return f.read()

    # 康軒 general
    kx_file = os.path.join(KX_DIR, f'臺中市115學年{grade_num}年級教學進度總表(英{eng_ver}版).docx')
    kx_dir = unpack(kx_file, 'kx')
    if not kx_dir: continue
    kx_xml = read_xml(kx_dir)

    # 客語 (grades 1,3,5)
    hk_xml = None
    if hoklo_type == '客':
        hk_file = os.path.join(HK_DIR, f'臺中市115學年小客{grade_num}年級課程計畫.docx')
        hk_dir = unpack(hk_file, 'hk')
        if hk_dir: hk_xml = read_xml(hk_dir)

    # 數學
    if grade_num <= 4:
        math_s1_file = os.path.join(N1_DIR, f'01.上學期\\{grade_num}年級\\橫式115(一)小數{grade_num}上課程計畫.docx')
        math_s2_file = os.path.join(N1_DIR, f'02.下學期\\{grade_num}年級\\橫式115(二)小數{grade_num}下課程計畫.docx')
    else:
        grade_cn = {5:'五',6:'六'}[grade_num]
        math_s1_file = os.path.join(HL_DIR, f'{grade_cn}年級\\上學期\\115{grade_cn}上數學.docx')
        math_s2_file = os.path.join(HL_DIR, f'{grade_cn}年級\\下學期\\115{grade_cn}下數學.docx')

    math_s1_dir = unpack(math_s1_file, 'math_s1')
    math_s2_dir = unpack(math_s2_file, 'math_s2')
    math_s1_xml = read_xml(math_s1_dir) if math_s1_dir else None
    math_s2_xml = read_xml(math_s2_dir) if math_s2_dir else None

    # 健康體育 3-6 (南一)
    pe_s1_xml = pe_s2_xml = None
    if grade_num >= 3:
        pe_s1_file = os.path.join(N1_DIR, f'01.上學期\\{grade_num}年級\\橫式115(一)小健體{grade_num}上課程計畫.docx')
        pe_s2_file = os.path.join(N1_DIR, f'02.下學期\\{grade_num}年級\\橫式115(二)小健體{grade_num}下課程計畫.docx')
        pe_s1_dir = unpack(pe_s1_file, 'pe_s1')
        pe_s2_dir = unpack(pe_s2_file, 'pe_s2')
        pe_s1_xml = read_xml(pe_s1_dir) if pe_s1_dir else None
        pe_s2_xml = read_xml(pe_s2_dir) if pe_s2_dir else None

    print(f'  Unpacked all sources')

    # ── 2. Extract content ──────────────────────────────────────────────────

    kx_tables = get_tables(kx_xml)
    KX_DATA_ROWS = [3, 4, 5, 7, 8, 9]

    # Grades 1-2: different KX table structure
    #   col8=生活課程, col9=健體(→remapped to col12), col10=彈性(ignored)
    if grade_num <= 2:
        KX_COPY_CELLS = [4, 5, 6, 8]
        KX_EXTRACT_CELLS = [4, 5, 6, 7, 8, 9]
    else:
        KX_COPY_CELLS = [4, 5, 6, 8, 9, 10, 11]
        KX_EXTRACT_CELLS = KX_COPY_CELLS

    kx_s1 = parse_kx_table(kx_xml[kx_tables[0][0]:kx_tables[0][1]],
                            KX_DATA_ROWS, KX_EXTRACT_CELLS if grade_num <= 2 else KX_COPY_CELLS)
    kx_s2 = parse_kx_table(kx_xml[kx_tables[2][0]:kx_tables[2][1]],
                            KX_DATA_ROWS, KX_EXTRACT_CELLS if grade_num <= 2 else KX_COPY_CELLS)

    # Grades 1-2: remap col9(健體)→col12
    if grade_num <= 2:
        for kx_data in [kx_s1, kx_s2]:
            for ri in kx_data:
                if 9 in kx_data[ri]:
                    kx_data[ri][12] = kx_data[ri][9]
                    del kx_data[ri][9]
        KX_COPY_CELLS.append(12)

    # 客語
    hk_periods = None
    if hk_xml:
        all_periods = parse_weekly_file(hk_xml)
        hk_periods = all_periods[:2] if len(all_periods) >= 2 else [all_periods[0], all_periods[0]]

    # 數學
    math_s1_periods = parse_weekly_file(math_s1_xml)[0] if math_s1_xml else None
    math_s2_periods = parse_weekly_file(math_s2_xml)[0] if math_s2_xml else None

    # 健康體育 3-6
    pe_s1_periods = parse_weekly_file(pe_s1_xml)[0] if pe_s1_xml else None
    pe_s2_periods = parse_weekly_file(pe_s2_xml)[0] if pe_s2_xml else None

    print(f'  Parsed all content')

    # ── 3. Unpack and modify target 表3 file ──────────────────────────────

    target_file = os.path.join(OUT_DIR, f'115表3課程計畫_{grade_name}.docx')
    tgt_dir = os.path.join(grade_work, 'tgt')
    if os.path.exists(tgt_dir): shutil.rmtree(tgt_dir)
    r = subprocess.run(['python', UNPACK, target_file, tgt_dir],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        print(f'  Unpack target error: {r.stderr[:150]}'); continue

    tgt_doc = os.path.join(tgt_dir, 'word', 'document.xml')
    with open(tgt_doc, encoding='utf-8') as f:
        tgt_xml = f.read()

    def apply_to_table(tgt_xml, tgt_tbl_idx, kx_data, hk_period_pair, math_periods, pe_periods):
        """Fill cells 4-12 in rows 4,5,6,8,9,10 of one semester table."""
        ts, te = get_tables(tgt_xml)[tgt_tbl_idx]
        tbl_xml = tgt_xml[ts:te]

        ROW_PERIOD = {3: 0, 4: 0, 5: 0, 7: 1, 8: 1, 9: 1}
        ROW_TYPE   = {3: 'units', 4: 'teach', 5: 'assess', 7: 'units', 8: 'teach', 9: 'assess'}

        initial_rows = get_rows(tbl_xml)

        for ri, (period_idx, rtype) in [(r, (ROW_PERIOD[r], ROW_TYPE[r]))
                                         for r in ROW_PERIOD if r < len(initial_rows)]:
            tbl_rows = get_rows(tbl_xml)
            if ri >= len(tbl_rows): continue
            rs, re_ = tbl_rows[ri]
            row_xml = tbl_xml[rs:re_]

            # Grades 1-2: clear 自然(9)/藝術(10)/綜合(11) — not separate subjects
            if grade_num <= 2:
                for ci_clear in [9, 10, 11]:
                    row_xml = replace_cell_content(row_xml, ci_clear, make_plain_para(''))

            # Fill from 康軒 integrated
            for ci in KX_COPY_CELLS:
                if ri in kx_data and ci in kx_data[ri]:
                    row_xml = replace_cell_content(row_xml, ci,
                        make_plain_para(kx_data[ri][ci]))

            # Override cell 5 with 客語
            if hk_period_pair:
                period = hk_period_pair[period_idx]
                texts = period_to_texts(period)
                row_xml = replace_cell_content(row_xml, 5,
                    make_plain_para(texts[rtype] if rtype in texts else texts['units']))

            # Fill 數學 (cell 7)
            if math_periods:
                period = math_periods[period_idx]
                texts = period_to_texts(period)
                row_xml = replace_cell_content(row_xml, 7,
                    make_plain_para(texts[rtype] if rtype in texts else texts['units']))

            # Fill 健康體育 (cell 12) for grades 3-6
            if pe_periods:
                period = pe_periods[period_idx]
                texts = period_to_texts(period)
                row_xml = replace_cell_content(row_xml, 12,
                    make_plain_para(texts[rtype] if rtype in texts else texts['units']))

            tbl_xml = tbl_xml[:tbl_rows[ri][0]] + row_xml + tbl_xml[tbl_rows[ri][1]:]

        return tgt_xml[:ts] + tbl_xml + tgt_xml[te:]

    tgt_xml = apply_to_table(tgt_xml, 0, kx_s1,
                              hk_periods[0] if hk_periods else None,
                              math_s1_periods, pe_s1_periods)
    tgt_xml = apply_to_table(tgt_xml, 2, kx_s2,
                              hk_periods[1] if hk_periods else None,
                              math_s2_periods, pe_s2_periods)

    with open(tgt_doc, 'w', encoding='utf-8') as f:
        f.write(tgt_xml)

    # ── 4. Repack ──────────────────────────────────────────────────────────
    r = subprocess.run(['python', PACK, tgt_dir, target_file,
                        '--original', TPL_DOCX, '--validate', 'false'],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        print(f'  Pack error: {r.stderr[:200]}')
    else:
        size = os.path.getsize(target_file)
        print(f'  ✅ Saved: {os.path.basename(target_file)} ({size:,} bytes)')

print('\n全部完成！')
