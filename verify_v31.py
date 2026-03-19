# -*- coding: utf-8 -*-
import re
import os

DEBUG_DIR = r"C:\Users\Administrator\Desktop\deepseek_listing_tool"

HEADER_PATTERN = re.compile(r'^A\s+G\s+AD\s+AE\s+AF\s+AG\s+AH\s+AI\s+AJ\s+AK\s+AL\s+AM\s+AN\s*$', re.IGNORECASE)

def extract_last_reply(text, oe_number):
    if not text or len(text) < 50:
        return ''
    lines = text.split('\n')
    last_header_idx = -1
    for i, line in enumerate(lines):
        if HEADER_PATTERN.match(line.strip()):
            last_header_idx = i
    if last_header_idx < 0:
        return _extract_by_oe_fallback(text, oe_number)
    data_lines = []
    for i in range(last_header_idx + 1, len(lines)):
        stripped = lines[i].strip()
        if HEADER_PATTERN.match(stripped):
            break
        if re.match(r'^(已读取|Read).*网页', stripped, re.IGNORECASE):
            continue
        if re.match(r'^\d+$', stripped):
            continue
        if stripped in ['思考过程', '深度思考', 'DeepSeek 思考']:
            break
        if stripped:
            data_lines.append(stripped)
    if not data_lines:
        return _extract_by_oe_fallback(text, oe_number)
    merged = ' '.join(data_lines)
    merged = re.sub(r'\)\s+[\d\s]+([,;.])', r')\1', merged)
    merged = re.sub(r'\s+', ' ', merged)
    return merged.strip()

def _extract_by_oe_fallback(text, oe_number):
    lines = text.split('\n')
    last_oe_idx = -1
    for i, line in enumerate(lines):
        if HEADER_PATTERN.match(line.strip()):
            for j in range(i + 1, min(i + 5, len(lines))):
                if oe_number in lines[j]:
                    last_oe_idx = j
                    break
    if last_oe_idx < 0:
        last_pos = text.rfind(oe_number)
        if last_pos < 0:
            return ''
        remaining = text[last_pos:]
    else:
        remaining = '\n'.join(lines[last_oe_idx:])
    cleaned_lines = []
    for line in remaining.split('\n'):
        stripped = line.strip()
        if re.match(r'^(已读取|Read).*网页', stripped, re.IGNORECASE):
            continue
        if re.match(r'^\d+$', stripped):
            continue
        if stripped in ['思考过程', '深度思考', 'DeepSeek 思考']:
            continue
        if stripped:
            cleaned_lines.append(stripped)
    merged = ' '.join(cleaned_lines)
    merged = re.sub(r'\)\s+[\d\s]+([,;.])', r')\1', merged)
    merged = re.sub(r'\s+', ' ', merged)
    return merged.strip()

def parse_deepseek_fixed(text, oe_number):
    if not text or len(text) < 100:
        return None
    cleaned = extract_last_reply(text, oe_number)
    if not cleaned or len(cleaned) < 50:
        print(f"    X extract failed")
        return None
    print(f"    extracted: {len(cleaned)} chars")
    result = {
        'A': oe_number, 'G': '', 'AD': '', 'AE': '', 'AF': '', 'AG': '',
        'AH': '', 'AI': '', 'AJ': '', 'AK': '', 'AL': '', 'AM': '', 'AN': '',
    }
    first_oe_pos = cleaned.find(oe_number)
    if first_oe_pos < 0:
        return None
    vf_match = re.search(r'Vehicle\s+Fitment:', cleaned, re.IGNORECASE)
    if not vf_match:
        return None
    vf_end = vf_match.end()
    oe_matches = list(re.finditer(re.escape(oe_number), cleaned))
    last_oe_end = oe_matches[-1].end() if oe_matches else len(cleaned)
    title_raw = cleaned[first_oe_pos + len(oe_number):vf_match.start()].strip()
    title_raw = re.sub(r'^\s*\d{4,5}[-]\d{4,5}\s*', '', title_raw).strip()
    title_raw = re.sub(r'\s+\d{4,5}[-]\d{4,5}\s*$', '', title_raw).strip()
    result['G'] = title_raw
    if not result['G']:
        return None
    bullet_patterns = [
        r'Verify\s+vehicle\s+fitment', r'Direct\s+fit\s+for', r'Replaces\s+OEM',
        r'Provides\s+accurate', r'Measures\s+intake', r'Precise\s+Signal',
        r'Precise\s+Air', r'Plug\s*&?\s*play', r'Direct\s+plug', r'No\s+splicing',
        r'Factory-Grade', r'Meets\s+Toyota', r'Meets\s+OE',
        r'Corrosion-resistant', r'Delivers\s+accurate',
        r'100%\s+tested', r'1-?year\s+warranty',
    ]
    after_vf = cleaned[vf_end:]
    bullet_positions = []
    for pattern in bullet_patterns:
        for m in re.finditer(pattern, after_vf, re.IGNORECASE):
            bullet_positions.append({'abs_pos': vf_end + m.start(), 'end_pos': vf_end + m.end(), 'match': m.group(0)})
    bullet_positions.sort(key=lambda x: x['abs_pos'])
    filtered = []
    for bp in bullet_positions:
        if not filtered or bp['abs_pos'] - filtered[-1]['abs_pos'] > 5:
            filtered.append(bp)
    filtered = [bp for bp in filtered if vf_end < bp['abs_pos'] < last_oe_end]
    print(f"    {len(filtered)} bullets")
    if filtered:
        result['AD'] = cleaned[vf_end:filtered[0]['abs_pos']].strip()
    else:
        result['AD'] = cleaned[vf_end:last_oe_end].strip()
    result['AD'] = re.sub(r'\s*\.?\s*$', '', result['AD'])
    col_names = ['AE', 'AF', 'AG', 'AH', 'AI']
    bullets_to_use = filtered[:5]
    for i, bp in enumerate(bullets_to_use):
        content_start = bp['abs_pos']
        if i + 1 < len(bullets_to_use):
            content_end = bullets_to_use[i + 1]['abs_pos']
        else:
            content_end = last_oe_end
        content = cleaned[content_start:content_end].strip()
        content = re.sub(r'\s+', ' ', content)
        if i < len(col_names):
            result[col_names[i]] = content
    if last_oe_end < len(cleaned):
        kw_text = cleaned[last_oe_end:].strip()
        kw_text = re.sub(r'(已读取|Read).*$', '', kw_text, flags=re.DOTALL).strip()
        keywords = [k.strip() for k in kw_text.split() if len(k.strip()) > 1]
        for i, kw in enumerate(keywords[:5]):
            if i < 5:
                result[['AJ', 'AK', 'AL', 'AM', 'AN'][i]] = kw
    return result if result['G'] else None


# 测试 v3 调试文件（包含历史对话）
print("=" * 80)
print("  测试 v3 调试文件（包含历史对话的完整页面文本）")
print("=" * 80)

# v3_row7 对应 OE=22204-24010，但页面含3条回复
# v3_row8 对应 OE=89543-28110，但页面含3条回复
# v3_row9 对应 OE=MR249526，但页面含3条回复
test_cases = [
    (7, '22204-24010', 'debug_v3_row7.txt'),
    (8, '89543-28110', 'debug_v3_row8.txt'),
    (9, 'MR249526', 'debug_v3_row9.txt'),
]

for row, oe_number, filename in test_cases:
    filepath = os.path.join(DEBUG_DIR, filename)
    if not os.path.exists(filepath):
        continue
    print(f"\n{'='*80}")
    print(f"  Row {row}: {oe_number} (file: {filename})")
    print(f"{'='*80}")
    with open(filepath, 'r', encoding='utf-8') as f:
        raw = f.read()
    print(f"  原始文本: {len(raw)} 字符")
    
    # 先测试 extract_last_reply
    extracted = extract_last_reply(raw, oe_number)
    print(f"  extract_last_reply: {len(extracted)} 字符")
    print(f"  提取内容开头: {extracted[:120]}...")
    
    result = parse_deepseek_fixed(raw, oe_number)
    if result:
        print(f"\n  SUCCESS!")
        for col in ['A', 'G', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN']:
            val = result.get(col, '')
            if val:
                d = val[:100] + ('...' if len(val) > 100 else '')
                print(f"    {col:4} ({len(val):4}c) -> {d}")
            else:
                print(f"    {col:4} -> (empty)")
    else:
        print(f"\n  FAILED")