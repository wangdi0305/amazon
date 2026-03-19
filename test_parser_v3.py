# -*- coding: utf-8 -*-
"""
测试 DeepSeek 回复解析器 v3 - 智能版

核心策略：
1. 定位表格区域
2. 使用关键模式提取各个字段
3. 结合表格行和段落内容
"""
import re


def parse_deepseek_response(text, oe_number):
    """
    智能解析 DeepSeek 的回复。

    DeepSeek 输出结构：
    1. 表格行（空格分隔）: A G AD AE AF AG AH AI AJ AK AL AM AN
    2. 数据行（空格分隔）：OE号, Title, Description, Bullet1-5, Keyword1-5
    3. 详细段落（带关键词）：Vehicle Fitment: Direct Fitment: OEM Cross Reference: 等
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
    # 找到表头行和数据行
    header_idx = -1
    data_idx = -1

    for i, line in enumerate(lines):
        stripped = line.strip()
        # 表头行特征：包含 A G AD AE 等列标识
        if re.match(r'^A\s+G\s+AD\s+AE\s+AF\s+AG\s+AH\s+AI\s+AJ\s+AK\s+AL\s+AM\s+AN$', stripped):
            header_idx = i
            # 检查下一行是否是数据
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                # 数据行应该以 OE 号开头
                if next_line.startswith(oe_number):
                    data_idx = i + 1
            break

    if data_idx == -1:
        print(f"    ✗ 未找到表格数据行")
        return None

    # ==================== 步骤2: 提取表格行数据 ====================
    data_line = lines[data_idx].strip()

    # 提取 OE 号和重复的 OE 号
    # 格式: MR249526 MR249526 Title Description Bullet1 Bullet2...
    fields = data_line.split()
    if len(fields) < 3:
        print(f"    ✗ 数据字段不足")
        return None

    # 字段 0: OE 号
    # 字段 1: OE 号（重复）
    # 字段 2-?: Title（到 "compatible with" 之间）

    # 找到 "compatible" 或 "for" 或 "fits" 的位置
    title_end_idx = -1
    for i, field in enumerate(fields):
        if field.lower() in ['compatible', 'for', 'fits']:
            title_end_idx = i
            break

    if title_end_idx == -1:
        print(f"    ✗ 无法定位标题结束位置")
        return None

    # 提取标题（从字段 2 到 title_end_idx + 1）
    # 注意：第一个字段（OE号）重复了一次，所以从 index 2 开始
    if len(fields) < 3:
        print(f"    ✗ 字段数量不足")
        return None

    # Title 部分：fields[2] 到 fields[title_end_idx]
    # 但是实际上表格中可能还有其他结构
    # 让我重新分析：
    # 从调试文件看：MR249526 MR249526 ABS Wheel Speed Sensor Rear Left compatible with Mitsubishi...
    # 实际上 OE 号重复了，然后是标题

    # 重新解析：
    title_start_idx = 2  # 跳过前两个 OE 号
    title_fields = fields[title_start_idx:title_end_idx + 1]
    result['title'] = ' '.join(title_fields)

    # 提取年份和其他信息
    # 继续查找年份（4位数字）
    year_match = None
    for i in range(title_end_idx, min(title_end_idx + 10, len(fields))):
        if re.match(r'^\d{4}$', fields[i]):
            year_match = fields[i]
            break

    # ==================== 步骤3: 从段落中提取详细内容 ====================
    # 查找 "Vehicle Fitment:" 开始的段落
    vehicle_section = extract_section(text, 'Vehicle Fitment:', ['Direct Fitment', 'OEM Cross Reference', 'Precise', 'Plug', 'Durable', '1-Year', '写作'])
    if vehicle_section:
        result['description'] = clean_text(vehicle_section)

    # 提取 "OEM Cross Reference:"
    oem_section = extract_section(text, 'OEM Cross Reference:', ['Precise', 'Plug', 'Durable', '1-Year', '写作'])
    if oem_section:
        result['bullet2'] = clean_text(oem_section)

    # 提取 "Precise Signal Output:"
    precise_section = extract_section(text, 'Precise Signal Output:', ['Plug', 'Durable', '1-Year', '写作'])
    if precise_section:
        result['bullet3'] = clean_text(precise_section)

    # 提取 "Plug & Play Installation:"
    plug_section = extract_section(text, ['Plug & Play Installation:', 'Plug Play Installation:'], ['Durable', '1-Year', '写作'])
    if plug_section:
        result['bullet4'] = clean_text(plug_section)

    # 提取 "Durable Construction:"
    durable_section = extract_section(text, 'Durable Construction:', ['1-Year', '写作'])
    if durable_section:
        durable_text = clean_text(durable_section)
        # 如果已经有 bullet5，合并；否则设置为 bullet5
        if result['bullet5']:
            result['bullet5'] = result['bullet5'] + ' ' + durable_text
        else:
            result['bullet5'] = durable_text

    # 提取 "1-Year Warranty:"
    warranty_section = extract_section(text, '1-Year Warranty:', ['写作', '关键词'])
    if warranty_section:
        warranty_text = clean_text(warranty_section)
        if result['bullet5']:
            result['bullet5'] = result['bullet5'] + ' ' + warranty_text
        else:
            result['bullet5'] = warranty_text

    # 提取适配车型（Bullet 1）
    if result['description']:
        # 第一句通常是适配车型总结
        sentences = result['description'].split('.')
        if sentences:
            first_sentence = sentences[0].strip()
            # 检查是否包含 "compatible with"
            if 'compatible' in first_sentence.lower():
                result['bullet1'] = first_sentence
                # 从描述中移除
                result['description'] = '. '.join(sentences[1:]).strip()

    # 提取关键词（表格最后一列）
    # 查找 OE 号和关键词的组合
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
    """
    从文本中提取指定关键词开始的段落。

    Args:
        text: 完整文本
        keywords: 开始关键词（可以是字符串或列表）
        stop_keywords: 停止关键词列表

    Returns:
        提取的段落文本
    """
    if isinstance(keywords, str):
        keywords = [keywords]

    # 构建正则模式
    escaped_keywords = [re.escape(k) for k in keywords]
    pattern = '|'.join(escaped_keywords)

    # 查找开始位置
    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return None

    start_pos = match.end()

    # 查找结束位置
    stop_pos = len(text)
    for stop_kw in stop_keywords:
        stop_match = re.search(rf'\n\s*{re.escape(stop_kw)}', text[start_pos:], re.IGNORECASE)
        if stop_match:
            stop_pos = start_pos + stop_match.start()
            break

    # 提取内容
    content = text[start_pos:stop_pos].strip()

    # 移除引用标记
    content = re.sub(r'\s*\d+\s*', ' ', content)
    content = re.sub(r'\s+', ' ', content)

    return content


def clean_text(text):
    """清理文本"""
    if not text:
        return ''
    # 移除多余的空格和换行
    text = re.sub(r'\s+', ' ', text)
    # 移除首尾标点和空格
    text = text.strip()
    return text


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
            print(f"  OE:           {result['oe']}")
            print(f"  标题:         {result['title'][:80]}...")
            print(f"  描述:         {result['description'][:80]}...")
            print(f"  卖点1:        {result['bullet1'][:80]}...")
            print(f"  卖点2:        {result['bullet2'][:80]}...")
            print(f"  卖点3:        {result['bullet3'][:80]}...")
            print(f"  卖点4:        {result['bullet4'][:80]}...")
            print(f"  卖点5:        {result['bullet5'][:80]}...")
            print(f"  关键词1-3:     {result['keyword1']}, {result['keyword2']}, {result['keyword3']}")
        else:
            print(f"\n✗ 解析失败")


if __name__ == "__main__":
    test_parser()
