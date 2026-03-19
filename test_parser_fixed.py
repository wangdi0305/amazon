# -*- coding: utf-8 -*-
import re
import os

DEBUG_DIR = r"C:\Users\Administrator\Desktop\deepseek_listing_tool"


def clean_raw_text(text):
    """清理 DeepSeek 原始输出"""
    if not text:
        return ''
    lines = text.split('\n')
    cleaned_lines = []
    data_started = False
    for line in lines:
        stripped = line.strip()
        if re.match(r'^A\s+G\s+AD\s+AE\s+AF\s+AG\s+AH\s+AI\s+AJ\s+AK\s+AL\s+AM\s+AN\s*$', stripped):
            data_started = True
            continue
        if not data_started:
            continue
        if re.match(r'^(已读取|Read).*网页', stripped, re.IGNORECASE):
            continue
        if re.match(r'^\d+$', stripped):
            continue
        if stripped:
            cleaned_lines.append(stripped)
    merged = ' '.join(cleaned_lines)
    merged = re.sub(r'\)\s+[\d\s]+([,;.])', r')\1', merged)
    merged = re.sub(r'\s+', ' ', merged)
    return merged.strip()


def parse_deepseek_fixed(text, oe_number):
    """修复版解析器 v5 - 最终版"""
    if not text or len(text) < 100:
        return None

    cleaned = clean_raw_text(text)
    if not cleaned or len(cleaned) < 50:
        return None

    print(f"    cleaned: {len(cleaned)} chars")

    result = {
        'A': oe_number, 'G': '', 'AD': '', 'AE': '', 'AF': '', 'AG': '',
        'AH': '', 'AI': '', 'AJ': '', 'AK': '', 'AL': '', 'AM': '', 'AN': '',
    }

    # === 1. 定位关键分隔点 ===
    first_oe_pos = cleaned.find(oe_number)
    if first_oe_pos < 0:
        print("    X OE not found"); return None

    vf_match = re.search(r'Vehicle\s+Fitment:', cleaned, re.IGNORECASE)
    if not vf_match:
        print("    X Vehicle Fitment not found"); return None
    vf_end = vf_match.end()

    # 最后一个 OE 号（关键词区域）
    oe_matches = list(re.finditer(re.escape(oe_number), cleaned))
    last_oe_end = oe_matches[-1].end() if oe_matches else len(cleaned)

    # === 2. G 列（标题）===
    title_raw = cleaned[first_oe_pos + len(oe_number):vf_match.start()].strip()
    title_raw = re.sub(r'^\s*\d{4,5}[-]\d{4,5}\s*', '', title_raw).strip()
    title_raw = re.sub(r'\s+\d{4,5}[-]\d{4,5}\s*$', '', title_raw).strip()
    result['G'] = title_raw
    if not result['G']:
        print("    X title empty"); return None

    # === 3. Bullet 起始标记（在 Vehicle Fitment 之后查找）===
    bullet_patterns = [
        r'Verify\s+vehicle\s+fitment',
        r'Direct\s+fit\s+for',
        r'Replaces\s+OEM',
        r'Provides\s+accurate',
        r'Measures\s+intake',
        r'Precise\s+Signal',
        r'Precise\s+Air',
        r'Plug\s*&?\s*play',
        r'Direct\s+plug',
        r'No\s+splicing',
        r'Factory-Grade',
        r'Meets\s+Toyota',
        r'Meets\s+OE',
        r'Corrosion-resistant',
        r'100%\s+tested',
        r'1-?year\s+warranty',
    ]

    after_vf = cleaned[vf_end:]
    bullet_positions = []
    for pattern in bullet_patterns:
        for m in re.finditer(pattern, after_vf, re.IGNORECASE):
            bullet_positions.append({
                'abs_pos': vf_end + m.start(),
                'end_pos': vf_end + m.end(),
                'match': m.group(0),
            })
    bullet_positions.sort(key=lambda x: x['abs_pos'])

    # 去重
    filtered = []
    for bp in bullet_positions:
        if not filtered or bp['abs_pos'] - filtered[-1]['abs_pos'] > 5:
            filtered.append(bp)

    # 过滤范围
    filtered = [bp for bp in filtered if vf_end < bp['abs_pos'] < last_oe_end]
    print(f"    found {len(filtered)} bullets")

    # === 4. AD 列 = Vehicle Fitment 到第一个 bullet ===
    if filtered:
        ad_end = filtered[0]['abs_pos']
        result['AD'] = cleaned[vf_end:ad_end].strip()
    else:
        result['AD'] = cleaned[vf_end:last_oe_end].strip()

    # 清理 AD 末尾
    result['AD'] = re.sub(r'\s*\.?\s*$', '', result['AD'])

    # === 5. AE-AI 列 = 前 5 个 bullet ===
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

    # === 6. 关键词 AJ-AN ===
    if last_oe_end < len(cleaned):
        kw_text = cleaned[last_oe_end:].strip()
        kw_text = re.sub(r'(已读取|Read).*$', '', kw_text, flags=re.DOTALL).strip()
        keywords = [k.strip() for k in kw_text.split() if len(k.strip()) > 1]
        for i, kw in enumerate(keywords[:5]):
            if i < 5:
                result[['AJ', 'AK', 'AL', 'AM', 'AN'][i]] = kw

    # 验证
    filled = sum(1 for v in result.values() if v)
    print(f"    filled {filled}/13 fields")
    for col in ['G', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ']:
        val = result.get(col, '')
        if val:
            print(f"      {col}: {len(val):4} chars | {val[:65]}...")

    return result if result['G'] else None


def main():
    print("=" * 80)
    print("  Offline Parse Test - Fixed v5 (FINAL)")
    print("=" * 80)

    test_cases = [
        (7, '22204-24010', 'debug_final_row7.txt'),
        (8, '89543-28110', 'debug_final_row8.txt'),
    ]

    for row, oe_number, filename in test_cases:
        filepath = os.path.join(DEBUG_DIR, filename)
        if not os.path.exists(filepath):
            continue

        print(f"\n{'='*80}")
        print(f"  Row {row}: {oe_number}")
        print(f"{'='*80}")

        with open(filepath, 'r', encoding='utf-8') as f:
            raw = f.read()

        result = parse_deepseek_fixed(raw, oe_number)

        if result:
            print(f"\n  SUCCESS!")
            for col in ['A', 'G', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN']:
                val = result.get(col, '')
                if val:
                    print(f"    {col:4} -> {val[:120]}{'...' if len(val) > 120 else ''}")
                else:
                    print(f"    {col:4} -> (empty)")
        else:
            print(f"\n  FAILED")

        print()


if __name__ == "__main__":
    main()