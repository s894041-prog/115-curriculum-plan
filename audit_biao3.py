"""
audit_biao3.py
反查 115表3課程計畫 各年級各科目來源是否正確。
對照 表2選用教科書一覽表，逐欄列出 ✅ / ❌ / ⚠️。
"""
import sys, re, os, shutil, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

OUT_DIR = r"D:\教學組\114教務\教學\課程計畫\claude 操作"
WORK    = r"C:\Users\admin\AppData\Local\Temp\biao3_audit2"

GRADES = [
    (1, '一年級'),
    (2, '二年級'),
    (3, '三年級'),
    (4, '四年級'),
    (5, '五年級'),
    (6, '六年級'),
]

COL_NAMES = {
    4:  '國語文',
    5:  '本土語文',
    6:  '英語文',
    7:  '數學',
    8:  '社會/生活',
    9:  '自然',
    10: '藝術',
    11: '綜合',
    12: '健體',
}

# 表2 選用教科書一覽表：每個年級各科預期來源關鍵字（出現在儲存格文字中即算符合）
# 1-2年級的 9/10/11 欄應為空白（合科為生活課程，不單獨列出）
EXPECTED_SRC = {
    1: {
        4:  '康軒',
        5:  '客',       # 客語
        6:  '康軒',
        7:  '南一',
        8:  '康軒',     # 生活課程
        12: '南一',     # 健體
    },
    2: {
        4:  '康軒',
        5:  '閩',       # 閩南語
        6:  '康軒',
        7:  '南一',
        8:  '康軒',     # 生活課程
        12: '南一',     # 健體
    },
    3: {
        4:  '康軒',
        5:  '客',
        6:  '康軒',
        7:  '南一',
        8:  '康軒',
        9:  '康軒',
        10: '康軒',
        11: '康軒',
        12: '南一',
    },
    4: {
        4:  '康軒',
        5:  '閩',
        6:  '康軒',
        7:  '南一',
        8:  '康軒',
        9:  '康軒',
        10: '康軒',
        11: '康軒',
        12: '南一',
    },
    5: {
        4:  '康軒',
        5:  '客',
        6:  '康軒',
        7:  '翰林',
        8:  '康軒',
        9:  '康軒',
        10: '康軒',
        11: '康軒',
        12: '南一',
    },
    6: {
        4:  '康軒',
        5:  '閩',
        6:  '康軒',
        7:  '翰林',
        8:  '康軒',
        9:  '康軒',
        10: '康軒',
        11: '康軒',
        12: '南一',
    },
}

# 1-2年級這些欄位應該是空的（無此科）
EXPECTED_EMPTY_1_2 = {9, 10, 11}

# ─── XML helpers ────────────────────────────────────────────────────────────

def get_tables(xml):
    result = []
    for m in re.finditer(r'<w:tbl>', xml):
        end = xml.index('</w:tbl>', m.start()) + len('</w:tbl>')
        result.append((m.start(), end))
    return result

def get_rows(tbl_xml):
    result = []
    for m in re.finditer(r'<w:tr[ >]', tbl_xml):
        end = tbl_xml.index('</w:tr>', m.start()) + len('</w:tr>')
        result.append((m.start(), end))
    return result

def get_cells(row_xml):
    result = []
    for m in re.finditer(r'<w:tc>', row_xml):
        end = row_xml.index('</w:tc>', m.start()) + len('</w:tc>')
        result.append((m.start(), end))
    return result

def cell_text(cell_xml):
    """Extract plain text from a cell."""
    return ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', cell_xml))

def extract_data_row(tbl_xml, row_1indexed):
    """Return list of cell texts for the given row (1-indexed)."""
    rows = get_rows(tbl_xml)
    ri = row_1indexed - 1
    if ri >= len(rows):
        return []
    row = tbl_xml[rows[ri][0]:rows[ri][1]]
    cells = get_cells(row)
    return [cell_text(row[cs:ce]) for cs, ce in cells]

# ─── Main ───────────────────────────────────────────────────────────────────

os.makedirs(WORK, exist_ok=True)

print('=' * 70)
print('115表3課程計畫 反查報告')
print('=' * 70)

total_ok = 0
total_err = 0
total_warn = 0

for grade_num, grade_name in GRADES:
    docx_path = os.path.join(OUT_DIR, f'115表3課程計畫_{grade_name}.docx')
    if not os.path.exists(docx_path):
        print(f'\n[{grade_name}] ❌ 找不到檔案: {docx_path}')
        continue

    # Unzip DOCX
    extract_dir = os.path.join(WORK, grade_name)
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(extract_dir)

    doc_xml_path = os.path.join(extract_dir, 'word', 'document.xml')
    with open(doc_xml_path, encoding='utf-8', errors='replace') as f:
        xml = f.read()

    tables = get_tables(xml)
    if len(tables) < 3:
        print(f'\n[{grade_name}] ⚠️ 表格數量不足 ({len(tables)})')
        continue

    tbl0 = xml[tables[0][0]:tables[0][1]]
    tbl2 = xml[tables[2][0]:tables[2][1]]

    print(f'\n{"─" * 50}')
    print(f'【{grade_name}】')
    print(f'{"─" * 50}')

    expected = EXPECTED_SRC.get(grade_num, {})

    for sem_label, tbl_xml, row_idx in [
        ('上學期（表1）', tbl0, 3),
        ('下學期（表3）', tbl2, 3),
    ]:
        print(f'  {sem_label}：')
        cells = extract_data_row(tbl_xml, row_idx)
        if not cells:
            print(f'    ⚠️ 無法讀取資料列')
            total_warn += 1
            continue

        for col_idx, col_name in COL_NAMES.items():
            ci = col_idx - 1  # 0-based
            if ci >= len(cells):
                text = ''
            else:
                text = cells[ci].strip()

            # 1-2年級 自然/藝術/綜合 應為空
            if grade_num <= 2 and col_idx in EXPECTED_EMPTY_1_2:
                if text == '':
                    print(f'    col{col_idx} {col_name}: ✅ 空白（無此科）')
                    total_ok += 1
                else:
                    print(f'    col{col_idx} {col_name}: ❌ 應為空但有內容 → "{text[:40]}"')
                    total_err += 1
                continue

            # 不在預期清單中的欄位
            if col_idx not in expected:
                continue

            keyword = expected[col_idx]
            preview = text[:60].replace('\n', ' ')
            if text == '':
                print(f'    col{col_idx} {col_name}: ⚠️ 空白（預期含「{keyword}」）')
                total_warn += 1
            elif keyword in text:
                print(f'    col{col_idx} {col_name}: ✅ 含「{keyword}」→ "{preview}"')
                total_ok += 1
            else:
                print(f'    col{col_idx} {col_name}: ❌ 未含「{keyword}」→ "{preview}"')
                total_err += 1

print(f'\n{"=" * 70}')
print(f'反查結果：✅ {total_ok} 項正確  ❌ {total_err} 項錯誤  ⚠️ {total_warn} 項警告')
print('=' * 70)
