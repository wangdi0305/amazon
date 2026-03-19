# -*- coding: utf-8 -*-
"""
测试 DeepSeek 回复解析器
"""
import re


def clean_paragraph(text):
    """清理段落文本"""
    if not text:
        return ''
    # 移除多余的空白和换行
    text = re.sub(r'\s+', ' ', text)
    # 移除引用标记
    text = re.sub(r'^\d+\s*', '', text)
    # 移除句尾的数字引用
    text = re.sub(r'\s*\d+\s*$', '', text)
    return text.strip()


def parse_deepseek_response(text, oe_number):
    """
    解析 DeepSeek 的回复，支持多种格式。
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

    # ==================== 策略1: 解析 tab 分隔的表格行 ====================
    lines = text.split('\n')
    table_found = False

    for i, line in enumerate(lines):
        if '\t' in line:
            # 尝试提取表头
            parts = [p.strip().upper() for p in line.split('\t')]
            if len(parts) >= 13 and 'A' in parts and 'G' in parts and 'AD' in parts:
                # 找到表头，下一行应该是数据
                if i + 1 < len(lines):
                    data_line = lines[i + 1]
                    if '\t' in data_line:
                        fields = [f.strip() for f in data_line.split('\t')]
                        if len(fields) >= 13:
                            result['oe'] = fields[0]
                            result['title'] = fields[1]
                            result['description'] = fields[2]
                            result['bullet1'] = fields[3]
                            result['bullet2'] = fields[4]
                            result['bullet3'] = fields[5]
                            result['bullet4'] = fields[6]
                            result['bullet5'] = fields[7]
                            result['keyword1'] = fields[8]
                            result['keyword2'] = fields[9]
                            result['keyword3'] = fields[10]
                            result['keyword4'] = fields[11]
                            result['keyword5'] = fields[12]
                            table_found = True
                            break

    # 如果表格解析成功，返回
    if table_found and result['title']:
        print(f"    ✓ 通过表格格式解析成功")
        return result

    # ==================== 策略2: 解析段落格式 ====================
    print(f"    尝试段落格式解析...")

    # 提取标题
    # 格式: [OE号] [Product Name] compatible with ...
    title_patterns = [
        rf'{re.escape(oe_number)}\s+(.+?)\s+compatible\s+with',
        rf'{re.escape(oe_number)}\s+(.+?)(?:compatible|for|fits)',
    ]
    for pattern in title_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            title_line = match.group(0)
            # 查找这个 OE 号在文本中的位置
            oe_pos = text.find(title_line)
            # 从该位置开始，查找第一行完整的内容
            after_oe = text[oe_pos:]
            first_newline = after_oe.find('\n')
            if first_newline > 0:
                # 取第一行作为标题，限制长度
                title = after_oe[:first_newline]
                # 移除 OE 号（可能重复）
                title = title.replace(oe_number, '').strip()
                if title.startswith('compatible'):
                    title = oe_number + ' ' + title
                else:
                    title = oe_number + ' ' + title
                result['title'] = title[:200]
                break

    # 提取各个段落
    # Vehicle Fitment: ... (描述)
    vehicle_fitment_match = re.search(r'Vehicle\s+Fitment:\s*(.*?)(?=\n\s*(?:Direct\s+Fitment|OEM\s+Cross\s+Reference|Precise|Plug|Durable|1-Year|关键词|$))',
                                       text, re.IGNORECASE | re.DOTALL)
    if vehicle_fitment_match:
        result['description'] = clean_paragraph(vehicle_fitment_match.group(1))

    # Precise Signal Output: ... (卖点3)
    precise_match = re.search(r'Precise\s+Signal\s+Output:\s*(.*?)(?=\n\s*(?:Plug|Durable|1-Year|关键词|$))',
                              text, re.IGNORECASE | re.DOTALL)
    if precise_match:
        result['bullet3'] = clean_paragraph(precise_match.group(1))

    # Plug & Play Installation: ... (卖点4)
    plug_match = re.search(r'Plug\s*&?\s*Play\s+Installation:\s*(.*?)(?=\n\s*(?:Durable|1-Year|关键词|$))',
                          text, re.IGNORECASE | re.DOTALL)
    if plug_match:
        result['bullet4'] = clean_paragraph(plug_match.group(1))

    # Durable Construction: ... (卖点5的一部分)
    durable_match = re.search(r'Durable\s+Construction:\s*(.*?)(?=\n\s*(?:1-Year|关键词|$))',
                             text, re.IGNORECASE | re.DOTALL)
    if durable_match:
        durable_text = clean_paragraph(durable_match.group(1))
        result['bullet5'] = durable_text

    # 1-Year Warranty: ... (卖点5的结尾)
    warranty_match = re.search(r'1-Year\s+Warranty:\s*(.*?)(?=\n\s*(?:关键词|写作思路|$))',
                              text, re.IGNORECASE | re.DOTALL)
    if warranty_match:
        warranty_text = clean_paragraph(warranty_match.group(1))
        if result['bullet5']:
            result['bullet5'] = result['bullet5'] + ' ' + warranty_text
        else:
            result['bullet5'] = warranty_text

    # OEM Cross Reference: ... (卖点2)
    oem_match = re.search(r'OEM\s+Cross\s+Reference:\s*(.*?)(?=\n\s*(?:Precise|Plug|Durable|1-Year))',
                         text, re.IGNORECASE | re.DOTALL)
    if oem_match:
        oem_text = clean_paragraph(oem_match.group(1))
        result['bullet2'] = oem_text

    # 提取适配车型 (卖点1) - 从描述中提取第一句
    if result['description']:
        first_sentence = result['description'].split('.')[0]
        if 'compatible with' in first_sentence.lower() or 'compatible' in first_sentence.lower():
            result['bullet1'] = first_sentence.strip()
            # 从描述中移除
            result['description'] = result['description'][len(first_sentence):].strip()

    # 提取关键词
    # 通常在最后
    keywords_pattern = rf'{re.escape(oe_number)}\s+([\w\s\-\(\)]+?)\s*$'
    keywords_match = re.search(keywords_pattern, text, re.MULTILINE)
    if keywords_match:
        keywords_text = keywords_match.group(1).strip()
        keyword_list = [k.strip() for k in keywords_text.split() if k.strip()]
        for i, kw in enumerate(keyword_list[:5]):
            result[f'keyword{i+1}'] = kw

    # 验证是否提取到足够内容
    if result['title'] and (result['description'] or result['bullet1']):
        print(f"    ✓ 通过段落格式解析成功")
        return result
    else:
        print(f"    ✗ 段落格式解析失败")
        return None


# 测试
def test_parser():
    import os

    # 测试文件列表
    test_files = [
        ('debug_reply_row7.txt', 'MR249526'),
        ('debug_reply_row8.txt', '22204-24010'),
        ('debug_reply_row9.txt', '89543-28110'),
    ]

    for filename, oe_number in test_files:
        filepath = f"/mnt/c/Users/Administrator/Desktop/{filename}"

        if not os.path.exists(filepath):
            print(f"文件不存在: {filepath}")
            continue

        print(f"\n{'='*80}")
        print(f"测试文件: {filename}")
        print(f"OE号: {oe_number}")
        print(f"{'='*80}")

        # 读取文件
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        # 解析
        result = parse_deepseek_response(content, oe_number)

        if result:
            print(f"\n✓ 解析成功!")
            print(f"\n字段提取结果:")
            print(f"  OE: {result['oe']}")
            print(f"  标题: {result['title'][:100]}...")
            print(f"  描述: {result['description'][:100]}...")
            print(f"  卖点1: {result['bullet1'][:80]}...")
            print(f"  卖点2: {result['bullet2'][:80]}...")
            print(f"  卖点3: {result['bullet3'][:80]}...")
            print(f"  卖点4: {result['bullet4'][:80]}...")
            print(f"  卖点5: {result['bullet5'][:80]}...")
            print(f"  关键词: {result['keyword1']}, {result['keyword2']}, {result['keyword3']}")
        else:
            print(f"\n✗ 解析失败")


if __name__ == "__main__":
    test_parser()
