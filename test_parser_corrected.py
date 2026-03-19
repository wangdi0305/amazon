# -*- coding: utf-8 -*-
"""
修正后的解析器 - 根据 DeepSeek 写作思路说明
"""
import re

def parse_deepseek_response(text, oe_number):
    """
    根据 DeepSeek 写作思路说明解析回复。

    表格映射：
    A 列 (OE号): OE号
    G 列 (标题): [OE号] + [产品名] + compatible with + [品牌] + [车型] + [年份] + 兼容型号
    AD 列 (产品描述): Vehicle Fitment: 段落
    AE 列 (卖点1): Direct Fitment: 段落
    AF 列 (卖点2): OEM Cross Reference: 段落
    AG 列 (卖点3): Precise [功能]: 段落
    AH 列 (卖点4): Plug & Play Installation: 段落
    AI 列 (卖点5): Durable Construction: + 1-Year Warranty: 段落
    AJ-AN 列 (关键词): 关键词
    """
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

    # ==================== 步骤1: 定位表格区域 ====================
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
        print(f"    ✗ 未找到表格数据行")
        return None

    # ==================== 步骤2: 提取标题（G列）====================
    data_line = lines[data_idx].strip()
    fields = data_line.split()

    if len(fields) < 2:
        print(f"    ✗ 数据字段不足")
        return None

    # 根据 DeepSeek 说明：
    # A 列: 字段 0 (OE号)
    # G 列: 字段 1-? (标题)
    # 标题包含：OE号(重复) + 产品名 + compatible with + 品牌 + 车型 + 年份 + 兼容型号

    # 简单策略：从字段 1 开始，提取到表格行的结尾（但排除"Vehicle Fitment:"段落）
    title_fields = fields[1:]

    # 查找"Vehicle Fitment:"的位置，标题在它之前
    title_end_idx = -1
    for i, field in enumerate(title_fields):
        if field == 'Vehicle' and i + 1 < len(title_fields) and title_fields[i+1] == 'Fitment:':
            title_end_idx = i
            break

    if title_end_idx > 0:
        title_fields = title_fields[:title_end_idx]

    result['title'] = ' '.join(title_fields)

    # ==================== 步骤3: 从段落中提取详细内容 ====================
    # AD 列: Vehicle Fitment: -> 描述
    vehicle_section = extract_section(text, 'Vehicle Fitment:',
                                      ['Direct Fitment', 'OEM Cross Reference', 'Precise', 'Plug', 'Durable', '1-Year', '写作'])
    if vehicle_section:
        result['description'] = clean_text(vehicle_section)

    # AE 列: Direct Fitment: -> 卖点1
    direct_section = extract_section(text, 'Direct Fitment:',
                                    ['OEM Cross Reference', 'Precise', 'Plug', 'Durable', '1-Year', '写作'])
    if direct_section:
        result['bullet1'] = clean_text(direct_section)

    # AF 列: OEM Cross Reference: -> 卖点2
    oem_section = extract_section(text, 'OEM Cross Reference:',
                                  ['Precise', 'Plug', 'Durable', '1-Year', '写作'])
    if oem_section:
        result['bullet2'] = clean_text(oem_section)

    # AG 列: Precise ...: -> 卖点3
    # 可能是 "Precise Signal Output:" 或 "Precise Air Volume Measurement:" 等
    precise_patterns = ['Precise Signal Output:', 'Precise Air Volume Measurement:']
    precise_section = extract_section(text, precise_patterns,
                                      ['Plug', 'Durable', '1-Year', '写作'])
    if precise_section:
        result['bullet3'] = clean_text(precise_section)

    # AH 列: Plug & Play Installation: -> 卖点4
    plug_section = extract_section(text, ['Plug & Play Installation:', 'Plug Play Installation:'],
                                   ['Durable', '1-Year', '写作'])
    if plug_section:
        result['bullet4'] = clean_text(plug_section)

    # AI 列: Durable Construction: + 1-Year Warranty: -> 卖点5
    durable_section = extract_section(text, 'Durable Construction:',
                                      ['1-Year', '写作'])
    if durable_section:
        durable_text = clean_text(durable_section)
        result['bullet5'] = durable_text

    warranty_section = extract_section(text, '1-Year Warranty:',
                                      ['写作', '关键词'])
    if warranty_section:
        warranty_text = clean_text(warranty_section)
        if result['bullet5']:
            result['bullet5'] = result['bullet5'] + ' ' + warranty_text
        else:
            result['bullet5'] = warranty_text

    # AJ-AN 列: 关键词
    keywords_pattern = rf'{re.escape(oe_number)}\s+([\w\s\-\(\)]+?)\s*$'
    keywords_match = re.search(keywords_pattern, text, re.MULTILINE)
    if keywords_match:
        keywords_text = keywords_match.group(1).strip()
        keyword_list = [k.strip() for k in keywords_text.split() if k.strip()]
        for i, kw in enumerate(keyword_list[:5]):
            result[f'keyword{i+1}'] = kw

    # ==================== 验证 ====================
    if result['title'] and (result['description'] or result['bullet1']):
        print(f"    ✓ 解析成功!")
        return result
    else:
        print(f"    ✗ 解析失败: 缺少关键字段")
        print(f"       Title: {result['title'][:50] if result['title'] else 'None'}")
        print(f"       Description: {result['description'][:50] if result['description'] else 'None'}")
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


# 测试
def main():
    print("=" * 80)
    print("  修正后的解析器测试")
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

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
        except FileNotFoundError:
            print(f"✗ 文件不存在")
            continue

        result = parse_deepseek_response(content, oe_number)

        if result:
            print(f"\n✓ 解析成功!")
            print(f"\n提取结果:")
            print(f"  OE (A列):     {result['oe']}")
            print(f"  标题 (G列):   {result['title'][:100]}...")
            print(f"  描述 (AD列):  {result['description'][:100]}...")
            print(f"  卖点1 (AE列): {result['bullet1'][:100]}...")
            print(f"  卖点2 (AF列): {result['bullet2'][:100]}...")
            print(f"  卖点3 (AG列): {result['bullet3'][:100]}...")
            print(f"  卖点4 (AH列): {result['bullet4'][:100]}...")
            print(f"  卖点5 (AI列): {result['bullet5'][:100]}...")
            print(f"  关键词:        {result['keyword1']}, {result['keyword2']}, {result['keyword3']}")
        else:
            print(f"\n✗ 解析失败")


if __name__ == "__main__":
    main()
