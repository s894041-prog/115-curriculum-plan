"""
gen_biao4.py
Generate 115表4彈性學習課程計畫.docx from 114 source.

Directly edits document.xml inside the ZIP at XML run level.

Changes:
1. Year: 114→115  (the '4' sits in a separate run between '...11' and '學年度')
   Replace: <w:t>4</w:t></w:r>...<w:t>學年度</w:t>
          → <w:t>5</w:t></w:r>...<w:t>學年度</w:t>

2. 安全教育→交通安全 in <w:t> content, EXCEPT these 4 run patterns (keep as-is):
   - <w:t>安全教育)</w:t>          (2 occurrences — unit label )
   - <w:t xml:space="preserve"> (安全教育)</w:t>  (1 occurrence — unit label)
   - <w:t>、安全教育</w:t>          (1 occurrence — topic content text)
"""
import sys, re, os, shutil, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR = r"D:\教學組\114教務\教學\課程計畫"
SRC_DOCX = os.path.join(BASE_DIR, r"114表三表四\8-2 表4彈性學習課程計畫1140716益修(已經入攀樹課程)ok.docx")
OUT_DOCX = os.path.join(BASE_DIR, r"claude 操作\115表4彈性學習課程計畫.docx")

# ── Read ZIP ─────────────────────────────────────────────────────────────────
with zipfile.ZipFile(SRC_DOCX, 'r') as z:
    all_files = {name: z.read(name) for name in z.namelist()}

xml = all_files['word/document.xml'].decode('utf-8')

# ── 1. Year replacement ───────────────────────────────────────────────────────
# Pattern: <w:t>4</w:t></w:r> immediately followed (within ~300 chars) by <w:t>學年度</w:t>
year_pattern = r'(<w:t>)4(</w:t></w:r>(?:(?!</w:p>).){1,300}?<w:t>學年度</w:t>)'
n_year = len(re.findall(year_pattern, xml, flags=re.DOTALL))
xml = re.sub(year_pattern, r'\g<1>5\2', xml, flags=re.DOTALL)
print(f'Year 4→5: {n_year} replacement(s)')

# ── 2. Protect exceptions (replace with placeholders) ─────────────────────────
EXCEPTIONS = [
    '<w:t>安全教育)</w:t>',                          # (安全教育)2.This / 2.Time — unit labels
    '<w:t xml:space="preserve"> (安全教育)</w:t>',   # (安全教育)6.食在足夠 — unit label
    # NOTE: 品德、安全教育 spans 2 runs in ref XML: <w:t>品德、安全</w:t>+<w:t>教育、國際教育</w:t>
    # The source has <w:t>、安全教育</w:t> which SHOULD be replaced → 、交通安全
]
PLACEHOLDERS = {}
for i, exc in enumerate(EXCEPTIONS):
    ph = f'__KEEP_AQ_{i}__'
    count = xml.count(exc)
    xml = xml.replace(exc, ph)
    PLACEHOLDERS[ph] = exc
    print(f'Protected {count}x: {exc!r}')

# ── 3. Replace all remaining 安全教育 → 交通安全 ──────────────────────────────
n_replace = xml.count('安全教育')
xml = xml.replace('安全教育', '交通安全')
print(f'安全教育 → 交通安全: {n_replace} replacement(s)')

# ── 4. Restore exceptions ─────────────────────────────────────────────────────
for ph, original in PLACEHOLDERS.items():
    xml = xml.replace(ph, original)
    print(f'Restored: {original!r}')

# ── 5. Verify ────────────────────────────────────────────────────────────────
text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml))
print(f'\nVerification:')
print(f'  114學年度 in text: {text.count("114學年度")} (should be 0)')
print(f'  115學年度 in text: {text.count("115學年度")} (should be 2)')
print(f'  安全教育  in text: {text.count("安全教育")} (should be 4)')
print(f'  交通安全  in text: {text.count("交通安全")}')

# ── 6. Write ZIP ──────────────────────────────────────────────────────────────
all_files['word/document.xml'] = xml.encode('utf-8')
tmp = OUT_DOCX + '.tmp'
with zipfile.ZipFile(tmp, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for name, data in all_files.items():
        zout.writestr(name, data)
os.replace(tmp, OUT_DOCX)
print(f'\nSaved: {OUT_DOCX}  ({os.path.getsize(OUT_DOCX):,} bytes)')
print('Done!')
