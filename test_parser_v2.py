# -*- coding: utf-8 -*-
"""
测试 DeepSeek 回复解析器 v2 - 改进版
"""
import re


def parse_deepseek_response(text, oe_number):
    """
    解析 DeepSeek 的回复，支持多种格式。

    关键发现：DeepSeek 的表格输出可能是：
    1. 空格分隔：A G AD AE AF AG AH AI AJ AK AL AM AN
    2. Tab 分隔
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

    # ==================== 改进的表格解析策略 ====================
    # DeepSeek 的表格结构：
    # 1. 表头行：A G AD AE AF AG AH AI AJ AK AL AM AN
    # 2. 数据行：OE号, Title, Description, Bullet1-5, Keyword1-5

    # 策略A: 查找表头行 - 包含 A G AD AE 等列标识
    for i, line in enumerate(lines):
        # 尝试 Tab 分隔
        if '\t' in line:
            parts = [p.strip().upper() for p in line.split('\t')]
            if len(parts) >= 13 and 'A' in parts and 'G' in parts and 'AD' in parts:
                # 找到表头，解析下一行
                if i + 1 < len(lines):
                    data_line = lines[i + 1]
                    fields = [f.strip() for f in data_line.split('\t')]
                    if len(fields) >= 13 and fields[0] == oe_number:
                        return extract_fields_from_table(fields, result)

        # 尝试空格分隔（原始格式）
        # 表头行通常看起来像：A G AD AE AF AG AH AI AJ AK AL AM AN
        stripped_line = line.strip()
        # 检查是否是表头行
        header_pattern = r'^A\s+G\s+AD\s+AE\s+AF\s+AG\s+AH\s+AI\s+AJ\s+AK\s+AL\s+AM\s+AN$'
        if re.match(header_pattern, stripped_line):
            # 找到表头，解析下一行
            if i + 1 < len(lines):
                data_line = lines[i + 1].strip()
                # 数据行应该是：OE号, Title, Description, Bullet1-5, Keyword1-5
                # 用空格分割，但要注意字段内部可能包含空格
                fields = extract_fields_from_space_separated(data_line, oe_number)
                if fields and len(fields) >= 13:
                    return extract_fields_from_table(fields, result)

    # ==================== 段落解析策略 ====================
    print(f"    尝试段落格式解析...")

    # 找到包含 OE 号和表格的那一部分
    # 关键词："以下是为您生成的Excel横向表格内容"
    table_section_match = re.search(
        r'以下是为您生成的Excel横向表格内容.*?写作思路说明',
        text, re.DOTALL
    )

    if table_section_match:
        table_section = table_section_match.group(0)
        # 从中提取数据
        return parse_table_section(table_section, oe_number, result)

    # ==================== 最后尝试：从文本中逐行提取 ====================
    return extract_from_full_text(text, oe_number, result)


def extract_fields_from_space_separated(line, oe_number):
    """
    从空格分隔的行中提取字段。
    挑战：Title、Description 等字段本身可能包含空格
    解决方案：根据已知模式推断字段边界
    """
    # 先简单按空格分割
    parts = line.split()

    if not parts or parts[0] != oe_number:
        return None

    # 基本结构：
    # [0] OE号
    # [1] OE号（重复）
    # [2...] Title（可能多个词）
    # Description（长段落）
    # Bullet1（中等）
    # Bullet2（中等）
    # ...
    # Keyword1-5（单个词或短语）

    # 这个很复杂，我们换个策略
    # 查找关键词模式来定位字段
    return None


def extract_fields_from_table(fields, result):
    """从解析的字段数组中提取数据"""
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
    return result


def parse_table_section(section, oe_number, result):
    """
    从表格段落中解析数据
    """
    lines = section.split('\n')

    # 查找表头
    header_idx = -1
    for i, line in enumerate(lines):
        if 'A G AD AE AF AG AH AI AJ AK AL AM AN' in line or \
           re.match(r'^A\s+G\s+AD\s+AE', line.strip()):
            header_idx = i
            break

    if header_idx == -1:
        return None

    # 下一行应该是数据
    if header_idx + 1 < len(lines):
        data_line = lines[header_idx + 1]
        # 尝试多种分隔符
        for sep in ['\t', '  ']:  # Tab 或双空格
            if sep in data_line:
                fields = [f.strip() for f in data_line.split(sep)]
                if len(fields) >= 13 and fields[0] == oe_number:
                    print(f"    ✓ 找到表格并解析成功 (分隔符: {'Tab' if sep == '\t' else 'Space'})")
                    return extract_fields_from_table(fields, result)

        # 如果都不行，尝试手动解析
        # 这是一个复杂的任务，先简化处理
        return None

    return None


def extract_from_full_text(text, oe_number, result):
    """
    从完整文本中提取 - 使用正则表达式模式匹配
    """
    # 查找 OE 号后面跟着 "compatible with" 的标题
    title_match = re.search(
        rf'{re.escape(oe_number)}\s+(.+?compatible\s+with.+?)$',
        text, re.MULTILINE
    )
    if title_match:
        result['title'] = title_match.group(1).strip()[:200]

    # 查找各段落
    # Vehicle Fitment
    vehicle_match = re.search(
        r'Vehicle\s+Fitment:\s*(.*?)(?=\n\s*(?:Direct\s+Fitment|OEM|Precise|Plug|Durable|1-Year|写作|$))',
        text, re.DOTALL | re.IGNORECASE
    )
    if vehicle_match:
        result['description'] = vehicle_match.group(1).strip()[:2000]

    # Precise Signal Output
    precise_match = re.search(
        r'Precise\s+Signal\s+Output:\s*(.*?)(?=\n\s*(?:Plug|Durable|1-Year|写作|$))',
        text, re.DOTALL | re.IGNORECASE
    )
    if precise_match:
        result['bullet3'] = precise_match.group(1).strip()

    # Plug & Play
    plug_match = re.search(
        r'Plug\s*&?\s*Play\s+Installation:\s*(.*?)(?=\n\s*(?:Durable|1-Year|写作|$))',
        text, re.DOTALL | re.IGNORECASE
    )
    if plug_match:
        result['bullet4'] = plug_match.group(1).strip()

    # 1-Year Warranty
    warranty_match = re.search(
        r'1-Year\s+Warranty:\s*(.*?)(?=\n\s*(?:写作|$))',
        text, re.DOTALL | re.IGNORECASE
    )
    if warranty_match:
        result['bullet5'] = warranty_match.group(1).strip()

    # 验证
    if result['title'] or result['description']:
        print(f"    ✓ 通过文本模式解析成功")
        return result
    else:
        print(f"    ✗ 所有解析策略都失败")
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

        # 输出文件前50行用于调试
        print(f"\n文件前50行预览:")
        lines = content.split('\n')
        for i, line in enumerate(lines[:50]):
            print(f"{i:3}: {line[:120]}")

        # 解析
        result = parse_deepseek_response(content, oe_number)

        if result:
            print(f"\n✓ 解析成功!")
            print(f"\n字段提取结果:")
            for key, val in result.items():
                if val:
                    print(f"  {key:15}: {val[:100]}...")
        else:
            print(f"\n✗ 解析失败")


if __name__ == "__main__":
    test_parser()
