"""
Generate 6 表3 DOCX files (115 version) for grades 1-6.
Strategy: copy 115 blank template, fill 校訂課程 cells (14,15,16) and 週次 cell (1)
from the corresponding 114 source DOCX for each grade.
"""
import sys, re, os, shutil, subprocess
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SKILL_BASE = (r"C:\Users\admin\AppData\Roaming\Claude\local-agent-mode-sessions"
              r"\skills-plugin\f044c158-29da-435c-ae59-7ad9c15f098a"
              r"\7ba8d970-f36e-4ee9-afa9-35fb2a74155e\skills\docx")
UNPACK = os.path.join(SKILL_BASE, r"scripts\office\unpack.py")
PACK   = os.path.join(SKILL_BASE, r"scripts\office\pack.py")

TEMPLATE_FOLDER = r"D:\教學組\114教務\教學\課程計畫\claude 操作\表3_unpacked"
SRC_DIR   = r"D:\教學組\114教務\教學\課程計畫\114表三表四"
OUT_DIR   = r"D:\教學組\114教務\教學\課程計畫\claude 操作"
WORK_DIR  = r"C:\Users\admin\AppData\Local\Temp\biao3_work"
TPL_DOCX  = r"D:\教學組\114教務\教學\課程計畫\claude 操作\表3.docx"

GRADES = [
    (1, '一年級', '表3總表1-1140707修正.docx'),
    (2, '二年級', '表3總表2-1140707修正.docx'),
    (3, '三年級', '表3總表3-1140707修正.docx'),
    (4, '四年級', '表3總表4-1140707修正.docx'),
    (5, '五年級', '表3總表5-1140707修正.docx'),
    (6, '六年級', '表3總表6-1140707修正.docx'),
]

# ─── XML helpers ────────────────────────────────────────────────────────────

def get_tables(xml):
    """Return list of (start, end) for each top-level <w:tbl>...</w:tbl>."""
    return [(m.start(), xml.index('</w:tbl>', m.start()) + len('</w:tbl>'))
            for m in re.finditer(r'<w:tbl>', xml)]

def get_rows(tbl_xml):
    """Return list of (start, end) for each <w:tr>...</w:tr> in tbl_xml."""
    result = []
    for m in re.finditer(r'<w:tr[ >]', tbl_xml):
        end = tbl_xml.index('</w:tr>', m.start()) + len('</w:tr>')
        result.append((m.start(), end))
    return result

def get_cells(row_xml):
    """Return list of (start, end) for each <w:tc>...</w:tc> in row_xml."""
    result = []
    for m in re.finditer(r'<w:tc>', row_xml):
        end = row_xml.index('</w:tc>', m.start()) + len('</w:tc>')
        result.append((m.start(), end))
    return result

def get_para_content(cell_xml):
    """Return everything inside cell after </w:tcPr> and before </w:tc>."""
    idx = cell_xml.rfind('</w:tcPr>')
    if idx >= 0:
        return cell_xml[idx + len('</w:tcPr>'):cell_xml.rfind('</w:tc>')]
    inner_start = cell_xml.index('>') + 1
    return cell_xml[inner_start:cell_xml.rfind('</w:tc>')]

def replace_cell_paras(row_xml, cell_idx_0based, new_paras):
    cells = get_cells(row_xml)
    cs, ce = cells[cell_idx_0based]
    cell_xml = row_xml[cs:ce]
    tcpr_end_idx = cell_xml.rfind('</w:tcPr>')
    if tcpr_end_idx >= 0:
        new_cell = cell_xml[:tcpr_end_idx + len('</w:tcPr>')] + new_paras + '</w:tc>'
    else:
        open_tag_end = cell_xml.index('>') + 1
        new_cell = cell_xml[:open_tag_end] + new_paras + '</w:tc>'
    return row_xml[:cs] + new_cell + row_xml[ce:]

def process_table(tpl_xml, tpl_tbl_idx, src_tbl_xml):
    """
    Fill data rows 4 and 8 (1-indexed) of template table tpl_tbl_idx
    with content from src_tbl_xml rows 4 and 8.
    Copies: cell 1 (週次), cells 11→14, 12→15, 13→16
    """
    tpl_tables = get_tables(tpl_xml)
    ts, te = tpl_tables[tpl_tbl_idx]
    tbl_xml = tpl_xml[ts:te]

    src_rows = get_rows(src_tbl_xml)
    tbl_rows = get_rows(tbl_xml)

    for row_1indexed in [4, 8]:
        ri = row_1indexed - 1
        tbl_rows = get_rows(tbl_xml)
        if ri >= len(src_rows) or ri >= len(tbl_rows):
            print(f'    WARNING: row index {ri} out of range')
            continue

        src_row = src_tbl_xml[src_rows[ri][0]:src_rows[ri][1]]
        tpl_row = tbl_xml[tbl_rows[ri][0]:tbl_rows[ri][1]]

        src_cells = get_cells(src_row)
        tpl_cells = get_cells(tpl_row)

        # Cell 1: 週次 range (0-indexed: 0)
        if len(src_cells) > 0 and len(tpl_cells) > 0:
            src_c = src_row[src_cells[0][0]:src_cells[0][1]]
            paras = get_para_content(src_c)
            tpl_row = replace_cell_paras(tpl_row, 0, paras)
            tpl_cells = get_cells(tpl_row)

        # Cells 11→14, 12→15, 13→16 (0-indexed: 10→13, 11→14, 12→15)
        for src_ci, tpl_ci in [(10, 13), (11, 14), (12, 15)]:
            if src_ci < len(src_cells) and tpl_ci < len(tpl_cells):
                src_c = src_row[src_cells[src_ci][0]:src_cells[src_ci][1]]
                paras = get_para_content(src_c)
                tpl_row = replace_cell_paras(tpl_row, tpl_ci, paras)
                tpl_cells = get_cells(tpl_row)

        tbl_rows_cur = get_rows(tbl_xml)
        tbl_xml = tbl_xml[:tbl_rows_cur[ri][0]] + tpl_row + tbl_xml[tbl_rows_cur[ri][1]:]

    tpl_xml = tpl_xml[:ts] + tbl_xml + tpl_xml[te:]
    return tpl_xml

# ─── Main ───────────────────────────────────────────────────────────────────

os.makedirs(WORK_DIR, exist_ok=True)

for grade_num, grade_name, src_filename in GRADES:
    print(f'\n=== Grade {grade_num}: {grade_name} ===')

    src_file = os.path.join(SRC_DIR, src_filename)
    src_dir  = os.path.join(WORK_DIR, f'src{grade_num}')
    if os.path.exists(src_dir):
        shutil.rmtree(src_dir)
    r = subprocess.run(['python', UNPACK, src_file, src_dir],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        print(f'  Unpack error: {r.stderr[:200]}'); continue
    print(f'  Unpacked source OK')

    with open(os.path.join(src_dir, 'word', 'document.xml'), encoding='utf-8') as f:
        src_xml = f.read()

    src_tables = get_tables(src_xml)
    src_tbl1 = src_xml[src_tables[0][0]:src_tables[0][1]]
    src_tbl3 = src_xml[src_tables[2][0]:src_tables[2][1]]
    print(f'  Src Table1 rows={len(get_rows(src_tbl1))}, Table3 rows={len(get_rows(src_tbl3))}')

    tpl_dir = os.path.join(WORK_DIR, f'tpl{grade_num}')
    if os.path.exists(tpl_dir):
        shutil.rmtree(tpl_dir)
    shutil.copytree(TEMPLATE_FOLDER, tpl_dir)

    tpl_doc = os.path.join(tpl_dir, 'word', 'document.xml')
    with open(tpl_doc, encoding='utf-8') as f:
        tpl_xml = f.read()

    tpl_xml = tpl_xml.replace('___區', '外埔區')
    tpl_xml = tpl_xml.replace('___國民小學', '馬鳴國民小學')
    tpl_xml = tpl_xml.replace('___年級', grade_name)

    tpl_tables = get_tables(tpl_xml)
    print(f'  Tpl Table1 rows={len(get_rows(tpl_xml[tpl_tables[0][0]:tpl_tables[0][1]]))}, '
          f'Table3 rows={len(get_rows(tpl_xml[tpl_tables[2][0]:tpl_tables[2][1]]))}')

    tpl_xml = process_table(tpl_xml, 0, src_tbl1)
    tpl_xml = process_table(tpl_xml, 2, src_tbl3)
    print(f'  Data rows filled OK')

    with open(tpl_doc, 'w', encoding='utf-8') as f:
        f.write(tpl_xml)

    out_file = os.path.join(OUT_DIR, f'115表3課程計畫_{grade_name}.docx')
    r = subprocess.run(['python', PACK, tpl_dir, out_file,
                        '--original', TPL_DOCX, '--validate', 'false'],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        print(f'  Pack error: {r.stderr[:300]}')
    else:
        size = os.path.getsize(out_file)
        print(f'  Saved: {out_file}  ({size} bytes)')

print('\nAll done!')
