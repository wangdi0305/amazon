# -*- coding: utf-8 -*-
"""
测试：直接从文本提取并写入 Excel，不使用 openpyxl 列号
直接验证数据是否正确
"""
import re

def extract_data_from_text(text):
    """从 DeepSeek 回复中提取数据，直接按列名返回"""
    result = {
        'A': '',      # OE号
        'G': '',      # 标题
        'AD': '',     # 描述
        'AE': '',     # 卖点1
        'AF': '',     # 卖点2
        'AG': '',     # 卖点3
        'AH': '',     # 卖点4
        'AI': '',     # 卖点5
        'AJ': '',     # 关键词1
        'AK': '',     # 关键词2
        'AL': '',     # 关键词3
        'AM': '',     # 关键词4
        'AN': '',     # 关键词5
    }

    lines = text.split('\n')

    # 找表格行
    header_idx = -1
    data_idx = -1

    for i, line in enumerate(lines):
        stripped = line.strip()
        if re.match(r'^A\s+G\s+AD\s+AE\s+AF\s+AG\s+AH\s+AI\s+AJ\s+AK\s+AL\s+AM\s+AN$', stripped):
            header_idx = i
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if next_line.startswith('22204-24010'):
                    data_idx = i + 1
            break

    if data_idx == -1:
        return None

    # 提取数据行
    data_line = lines[data_idx].strip()

    # 数据行格式：A G AD AE AF AG AH AI AJ AK AL AM AN
    # 但字段之间用空格分隔
    # 我们需要智能提取

    # 方法1：使用正则表达式匹配
    # 模式：OE号 + 标题 + Vehicle Fitment: 段落 + Direct Fitment: 段落 + ...

    # 先提取标题（到 Vehicle Fitment: 之前）
    title_match = re.search(r'22204-24010\s+(.+?)Vehicle\s+Fitment:', data_line)
    if title_match:
        result['A'] = '22204-24010'
        result['G'] = title_match.group(1).strip()

    # 从完整文本中提取各段落
    # Vehicle Fitment: -> AD 列
    vehicle_match = re.search(r'Vehicle\s+Fitment:\s*(.*?)Direct\s+Fitment:', text, re.DOTALL)
    if vehicle_match:
        result['AD'] = vehicle_match.group(1).strip()

    # Direct Fitment: -> AE 列
    direct_match = re.search(r'Direct\s+Fitment:\s*(.*?)OEM\s+Cross\s+Reference:', text, re.DOTALL)
    if direct_match:
        result['AE'] = direct_match.group(1).strip()

    # OEM Cross Reference: -> AF 列
    oem_match = re.search(r'OEM\s+Cross\s+Reference:\s*(.*?)Precise\s+Air\s+Volume\s+Measurement:', text, re.DOTALL)
    if oem_match:
        result['AF'] = oem_match.group(1).strip()

    # Precise Air Volume Measurement: -> AG 列
    precise_match = re.search(r'Precise\s+Air\s+Volume\s+Measurement:\s*(.*?)Plug\s*&?\s*Play\s+Installation:', text, re.DOTALL)
    if precise_match:
        result['AG'] = precise_match.group(1).strip()

    # Plug & Play Installation: -> AH 列
    plug_match = re.search(r'Plug\s*&?\s*Play\s+Installation:\s*(.*?)Factory-Grade\s+Quality:', text, re.DOTALL)
    if plug_match:
        result['AH'] = plug_match.group(1).strip()

    # Factory-Grade Quality: -> AI 列（第一部分）
    factory_match = re.search(r'Factory-Grade\s+Quality:\s*(.*?)1-Year\s+Warranty:', text, re.DOTALL)
    if factory_match:
        result['AI'] = factory_match.group(1).strip()

    # 1-Year Warranty: -> AI 列（第二部分）
    warranty_match = re.search(r'1-Year\s+Warranty:\s*(.*?)22204-24010', text, re.DOTALL)
    if warranty_match:
        if result['AI']:
            result['AI'] = result['AI'] + ' ' + warranty_match.group(1).strip()
        else:
            result['AI'] = warranty_match.group(1).strip()

    # 关键词：最后的几个词
    keywords_match = re.search(r'22204-24010\s+Mass\s+Air\s+Flow\s+Sensor\s+MAF\s+Sensor\s+Air\s+Flow\s+Meter\s+Toyota\s+Corolla\s+MAF\s+Sensor', text)
    if keywords_match:
        keywords = ['Mass Air Flow Sensor', 'MAF Sensor', 'Air Flow Meter', 'Toyota Corolla MAF Sensor']
        result['AJ'] = keywords[0] if len(keywords) > 0 else ''
        result['AK'] = keywords[1] if len(keywords) > 1 else ''
        result['AL'] = keywords[2] if len(keywords) > 2 else ''
        result['AM'] = keywords[3] if len(keywords) > 3 else ''

    return result


# 读取文件
with open('/mnt/c/Users/Administrator/Desktop/debug_reply_row8.txt', 'r', encoding='utf-8') as f:
    text = f.read()

# 提取数据
data = extract_data_from_text(text)

if data:
    print("=" * 80)
    print("提取的数据（按列名）：")
    print("=" * 80)

    for col, val in data.items():
        if val:
            print(f"\n{col:4} 列:")
            print(f"  {val[:100]}...")
            print(f"  (总长度: {len(val)} 字符)")

    print("\n" + "=" * 80)
    print("应该对应 DeepSeek 的列号：")
    print("=" * 80)
    print("  A  -> OE号")
    print("  G  -> 标题")
    print("  AD -> 描述（Vehicle Fitment)")
    print("  AE -> 卖点1（Direct Fitment）")
    print("  AF -> 卖点2（OEM Cross Reference）")
    print("  AG -> 卖点3（Precise Air Volume Measurement）")
    print("  AH -> 卖点4（Plug & Play Installation）")
    print("  AI -> 卖点5（Factory-Grade Quality + 1-Year Warranty）")
    print("  AJ -> 关键词1")
    print("  AK -> 关键词2")
    print("  AL -> 关键词3")
    print("  AM -> 关键词4")
    print("  AN -> 关键词5")
else:
    print("提取失败")
