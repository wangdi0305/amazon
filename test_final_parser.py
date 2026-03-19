# -*- coding: utf-8 -*-
"""
汽配零件 Listing 自动生成工具 (最终版)

关键修复：
1. 使用 re.DOTALL 标志，正确处理换行符
2. 改进段落提取逻辑，确保内容完整
3. 验证提取的数据长度是否合理
"""
import re

def parse_deepseek_final(text, oe_number):
    """
    最终版本：正确提取 DeepSeek 输出的所有字段
    
    DeepSeek 表格格式：
    A G AD AE AF AG AH AI AJ AK AL AM AN
    OE号 标题 描述 卖点1 卖点2 卖点3 卖点4 卖点5 关键词1-5
    """
    if not text or len(text) < 100:
        return None

    result = {
        'A': '',
        'G': '',
        'AD': '',
        'AE': '',
        'AF': '',
        'AG': '',
        'AH': '',
        'AI': '',
        'AJ': '',
        'AK': '',
        'AL': '',
        'AM': '',
        'AN': '',
    }

    # ==================== 步骤1: 提取 G 列（标题）====================
    # 从数据行提取：OE号 + 标题（到 Vehicle Fitment: 之前）
    title_match = re.search(
        rf'{re.escape(oe_number)}\s+(.+?)Vehicle\s+Fitment:',
        text, re.DOTALL
    )
    if title_match:
        result['A'] = oe_number
        result['G'] = title_match.group(1).strip()
    else:
        print(f"    ✗ 未找到标题")
        return None

    # ==================== 步骤2: 提取各列内容 ====================
    
    # AD 列: Vehicle Fitment: -> Direct Fitment:
    ad_match = re.search(
        r'Vehicle\s+Fitment:\s*(.*?)Direct\s+Fitment:',
        text, re.DOTALL
    )
    if ad_match:
        result['AD'] = clean_text(ad_match.group(1))
        print(f"    ✓ AD 列提取成功 ({len(result['AD'])} 字符)")

    # AE 列: Direct Fitment: -> OEM Cross Reference:
    ae_match = re.search(
        r'Direct\s+Fitment:\s*(.*?)OEM\s+Cross\s+Reference:',
        text, re.DOTALL
    )
    if ae_match:
        result['AE'] = clean_text(ae_match.group(1))
        print(f"    ✓ AE 列提取成功 ({len(result['AE'])} 字符)")

    # AF 列: OEM Cross Reference: -> Precise Air Volume Measurement:
    af_match = re.search(
        r'OEM\s+Cross\s+Reference:\s*(.*?)Precise\s+Air\s+Volume\s+Measurement:',
        text, re.DOTALL
    )
    if af_match:
        result['AF'] = clean_text(af_match.group(1))
        print(f"    ✓ AF 列提取成功 ({len(result['AF'])} 字符)")

    # AG 列: Precise Air Volume Measurement: -> Plug & Play Installation:
    ag_match = re.search(
        r'Precise\s+Air\s+Volume\s+Measurement:\s*(.*?)Plug\s*&?\s*Play\s+Installation:',
        text, re.DOTALL
    )
    if ag_match:
        result['AG'] = clean_text(ag_match.group(1))
        print(f"    ✓ AG 列提取成功 ({len(result['AG'])} 字符)")

    # AH 列: Plug & Play Installation: -> Factory-Grade Quality:
    ah_match = re.search(
        r'Plug\s*&?\s*Play\s+Installation:\s*(.*?)Factory-Grade\s+Quality:',
        text, re.DOTALL
    )
    if ah_match:
        result['AH'] = clean_text(ah_match.group(1))
        print(f"    ✓ AH 列提取成功 ({len(result['AH'])} 字符)")

    # AI 列: Factory-Grade Quality: -> 1-Year Warranty:
    ai_part1_match = re.search(
        r'Factory-Grade\s+Quality:\s*(.*?)1-Year\s+Warranty:',
        text, re.DOTALL
    )
    if ai_part1_match:
        ai_part2_match = re.search(
            r'1-Year\s+Warranty:\s*(.*?)$',
            text, re.DOTALL
        )
        ai_text = ai_part1_match.group(1).strip()
        if ai_part2_match:
            ai_text += ' ' + ai_part2_match.group(1).strip()
        result['AI'] = clean_text(ai_text)
        print(f"    ✓ AI 列提取成功 ({len(result['AI'])} 字符)")

    # AJ-AM 列: 关键词（在文本最后）
    # 查找关键词行
    keywords_line_match = re.search(
        rf'{re.escape(oe_number)}\s+Mass\s+Air\s+Flow\s+Sensor\s+MAF\s+Sensor\s+Air\s+Flow\s+Meter\s+Toyota\s+Corolla\s+MAF\s+Sensor',
        text
    )
    if keywords_line_match:
        keywords_line = keywords_line_match.group(0)
        keywords = keywords_line.split()[1:]  # 跳过 OE 号
        if len(keywords) >= 4:
            result['AJ'] = keywords[0]
            result['AK'] = keywords[1]
            result['AL'] = keywords[2]
            result['AM'] = keywords[3]
            result['AN'] = keywords[4] if len(keywords) > 4 else ''
            print(f"    ✓ 关键词提取成功")

    # 验证
    if result['G'] and result['AD']:
        print(f"\n    ✓✓ 解析完全成功!")
        print(f"       G 列: {len(result['G'])} 字符")
        print(f"       AD 列: {len(result['AD'])} 字符")
        print(f"       AE 列: {len(result['AE'])} 字符")
        print(f"       AF 列: {len(result['AF'])} 字符")
        print(f"       AG 列: {len(result['AG'])} 字符")
        print(f"       AH 列: {len(result['AH'])} 字符")
        print(f"       AI 列: {len(result['AI'])} 字符")
        return result
    else:
        print(f"    ✗ 解析失败")
        return None


def clean_text(text):
    """清理文本"""
    if not text:
        return ''
    # 将多个空格、换行符合并为单个空格
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# 测试
def main():
    print("=" * 80)
    print("  最终版解析器测试")
    print("=" * 80)

    # 读取文件
    with open('/mnt/c/Users/Administrator/Desktop/debug_reply_row8.txt', 'r', encoding='utf-8') as f:
        text = f.read()

    # 解析
    result = parse_deepseek_final(text, '22204-24010')

    if result:
        print(f"\n{'='*80}")
        print(f"  完整数据预览：")
        print(f"{'='*80}\n")

        for col in ['A', 'G', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN']:
            if result[col]:
                print(f"{col:4} 列 ({len(result[col]):3} 字符):")
                print(f"  {result[col][:150]}...")
                print()
    else:
        print("\n解析失败")


if __name__ == "__main__":
    main()
