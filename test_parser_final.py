# -*- coding: utf-8 -*-
"""
纯文本测试 - 验证解析逻辑（无需 openpyxl）
"""
import re

# ============================================================
# 解析器
# ============================================================
def parse_deepseek_response(text, oe_number):
    """智能解析 DeepSeek 的回复"""
    if not text or len(text) < 100:
        return None

    result = {
        'oe': oe_number,
        'title': '',
        'description': '',
        'bullet1': '',
        'bullet2': '',
        'bullet3': '',
        'bullet4': '',
        'bullet5': '',
        'keyword1': '',
        'keyword2': '',
        'keyword3': '',
        'keyword4': '',
        'keyword5': '',
    }

    lines = text.split('\n')

    # 定位表格区域
    header_idx = -1
    data_idx = -1

    for i, line in enumerate(lines):
        stripped = line.strip()
        if re.match(r'^A\s+G\s+AD\s+AE\s+AF\s+AG\s+AH\s+AI\s+AJ\s+AK\s+AL\s+AM\s+AN$', stripped):
            header_idx = i
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if next_line.startswith(oe_number):
                    data_idx = i + 1
            break

    if data_idx == -1:
        return None

    # 提取表格数据
    data_line = lines[data_idx].strip()
    fields = data_line.split()

    if len(fields) < 3:
        return None

    # 提取标题
    title_end_idx = -1
    for i, field in enumerate(fields):
        if field.lower() in ['compatible', 'for', 'fits']:
            title_end_idx = i
            break

    if title_end_idx == -1 or title_end_idx < 2:
        return None

    title_fields = fields[2:title_end_idx + 1]
    result['title'] = ' '.join(title_fields)

    # 提取段落内容
    vehicle_section = extract_section(text, 'Vehicle Fitment:',
                                      ['Direct Fitment', 'OEM Cross Reference', 'Precise', 'Plug', 'Durable', '1-Year', '写作'])
    if vehicle_section:
        result['description'] = clean_text(vehicle_section)

    oem_section = extract_section(text, 'OEM Cross Reference:',
                                  ['Precise', 'Plug', 'Durable', '1-Year', '写作'])
    if oem_section:
        result['bullet2'] = clean_text(oem_section)

    precise_section = extract_section(text, 'Precise Signal Output:',
                                      ['Plug', 'Durable', '1-Year', '写作'])
    if precise_section:
        result['bullet3'] = clean_text(precise_section)

    plug_section = extract_section(text, ['Plug & Play Installation:', 'Plug Play Installation:'],
                                   ['Durable', '1-Year', '写作'])
    if plug_section:
        result['bullet4'] = clean_text(plug_section)

    durable_section = extract_section(text, 'Durable Construction:',
                                      ['1-Year', '写作'])
    if durable_section:
        durable_text = clean_text(durable_section)
        if result['bullet5']:
            result['bullet5'] = result['bullet5'] + ' ' + durable_text
        else:
            result['bullet5'] = durable_text

    warranty_section = extract_section(text, '1-Year Warranty:',
                                      ['写作', '关键词'])
    if warranty_section:
        warranty_text = clean_text(warranty_section)
        if result['bullet5']:
            result['bullet5'] = result['bullet5'] + ' ' + warranty_text
        else:
            result['bullet5'] = warranty_text

    # 提取适配车型
    if result['description']:
        sentences = result['description'].split('.')
        if sentences:
            first_sentence = sentences[0].strip()
            if 'compatible' in first_sentence.lower():
                result['bullet1'] = first_sentence
                result['description'] = '. '.join(sentences[1:]).strip()

    # 提取关键词
    keywords_pattern = rf'{re.escape(oe_number)}\s+([\w\s\-\(\)]+?)\s*$'
    keywords_match = re.search(keywords_pattern, text, re.MULTILINE)
    if keywords_match:
        keywords_text = keywords_match.group(1).strip()
        keyword_list = [k.strip() for k in keywords_text.split() if k.strip()]
        for i, kw in enumerate(keyword_list[:5]):
            result[f'keyword{i+1}'] = kw

    if result['title'] and (result['description'] or result['bullet1']):
        return result
    else:
        return None


def extract_section(text, keywords, stop_keywords):
    """从文本中提取段落"""
    if isinstance(keywords, str):
        keywords = [keywords]

    escaped_keywords = [re.escape(k) for k in keywords]
    pattern = '|'.join(escaped_keywords)

    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return None

    start_pos = match.end()
    stop_pos = len(text)

    for stop_kw in stop_keywords:
        stop_match = re.search(rf'\n\s*{re.escape(stop_kw)}', text[start_pos:], re.IGNORECASE)
        if stop_match:
            stop_pos = start_pos + stop_match.start()
            break

    content = text[start_pos:stop_pos].strip()
    content = re.sub(r'\s*\d+\s*', ' ', content)
    content = re.sub(r'\s+', ' ', content)
    return content


def clean_text(text):
    """清理文本"""
    if not text:
        return ''
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# ============================================================
# 测试
# ============================================================
def main():
    print("=" * 80)
    print("  DeepSeek 响应解析测试 (纯文本版本)")
    print("=" * 80)

    test_cases = [
        ('debug_reply_row7.txt', 'MR249526'),
        ('debug_reply_row8.txt', '22204-24010'),
        ('debug_reply_row9.txt', '89543-28110'),
    ]

    for filename, oe_number in test_cases:
        filepath = f"/mnt/c/Users/Administrator/Desktop/{filename}"

        print(f"\n{'='*80}")
        print(f"文件: {filename}")
        print(f"OE号: {oe_number}")
        print(f"{'='*80}")

        # 读取文件
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
        except FileNotFoundError:
            print(f"✗ 文件不存在")
            continue

        # 解析
        result = parse_deepseek_response(content, oe_number)

        if result:
            print(f"\n✓ 解析成功!")
            print(f"\n提取结果:")
            print(f"  OE:           {result['oe']}")
            print(f"  标题:         {result['title']}")
            print(f"  描述:         {result['description'][:200]}...")
            print(f"  卖点1:        {result['bullet1'][:150]}...")
            print(f"  卖点2:        {result['bullet2'][:150]}...")
            print(f"  卖点3:        {result['bullet3'][:150]}...")
            print(f"  卖点4:        {result['bullet4'][:150]}...")
            print(f"  卖点5:        {result['bullet5'][:150]}...")
            print(f"  关键词:        {result['keyword1']}, {result['keyword2']}, {result['keyword3']}, {result['keyword4']}, {result['keyword5']}")
        else:
            print(f"\n✗ 解析失败")


if __name__ == "__main__":
    main()
