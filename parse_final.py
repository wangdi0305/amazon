# -*- coding: utf-8 -*-
"""
汽配零件 Listing 自动生成工具 (最终版 - 专门处理新格式)

专门处理最新的 DeepSeek 输出格式
"""
import re

def parse_deepseek_new_format(text, oe_number):
    """处理新的 DeepSeek 输出格式"""
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

    # 提取 G 列（标题）
    # 模式：OE号 + 标题（到 Vehicle Fitment: 之前）
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

    # 提取 AD 列（Vehicle Fitment: 到 Direct fit）
    # 修正：使用更宽松的模式
    ad_match = re.search(
        r'Vehicle\s+Fitment:\s*(.*?)Direct\s+fit\s+for\s+Toyota/Lexus',
        text, re.DOTALL
    )
    if ad_match:
        result['AD'] = clean_text(ad_match.group(1))

    # 提取 AE 列（Direct fit 到 Replaces OEM）
    ae_match = re.search(
        r'Direct\s+fit\s+for\s+Toyota/Lexus.*?Replaces\s+OEM\s+Numbers:',
        text, re.DOTALL
    )
    if ae_match:
        result['AE'] = clean_text(ae_match.group(1))

    # 提取 AF 列（Replaces OEM 到 Measures intake）
    af_match = re.search(
        r'Replaces\s+OEM\s+Numbers:\s*(.*?)Measures\s+intake\s+air\s+volume',
        text, re.DOTALL
    )
    if af_match:
        result['AF'] = clean_text(af_match.group(1))

    # 提取 AG 列（Measures intake 到 Plug & play）
    ag_match = re.search(
        r'Measures\s+intake\s+air\s+volume.*?Plug\s+&\s+play',
        text, re.DOTALL
    )
    if ag_match:
        result['AG'] = clean_text(ag_match.group(1))

    # 提取 AH 列（Plug & play 到 Meets Toyota）
    ah_match = re.search(
        r'Plug\s+&\s+play.*?Meets\s+Toyota/Lexus\s+OE\s+specs',
        text, re.DOTALL
    )
    if ah_match:
        result['AH'] = clean_text(ah_match.group(1))

    # 提取 AI 列（Meets Toyota 到 1-year warranty）
    ai_match = re.search(
        r'Meets\s+Toyota/Lexus\s+OE\s+specs.*?1-year\s+warranty',
        text, re.DOTALL
    )
    if ai_match:
        result['AI'] = clean_text(ai_match.group(1))

    # 提取关键词（在文本最后）
    keywords_match = re.search(
        rf'{re.escape(oe_number)}\s+([\w\s]+?)$',
        text, re.MULTILINE
    )
    if keywords_match:
        keywords_text = keywords_match.group(1).strip()
        keywords = keywords_text.split()
        for i, kw in enumerate(keywords[:5]):
            col_name = ['AJ', 'AK', 'AL', 'AM', 'AN'][i]
            result[col_name] = kw

    # 验证
    if result['G'] and result['AD'] and result['AE']:
        print(f"    ✓✓✓ 解析完全成功!")
        print(f"       G: {len(result['G'])} 字符")
        print(f"       AD: {len(result['AD'])} 字符")
        print(f"       AE: {len(result['AE'])} 字符")
        print(f"       AF: {len(result['AF'])} 字符")
        print(f"       AG: {len(result['AG'])} 字符")
        print(f"       AH: {len(result['AH'])} 字符")
        print(f"       AI: {len(result['AI'])} 字符")
        return result
    else:
        print(f"    ✗ 解析失败")
        return None


def clean_text(text):
    """清理文本"""
    if not text:
        return ''
    # 移除数字标记和多余空格
    text = re.sub(r'\d+\s*', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# 测试
def main():
    print("=" * 80)
    print("  最终版解析器测试")
    print("=" * 80)

    # 读取文件
    with open('/mnt/c/Users/Administrator/Desktop/debug_final_row7.txt', 'r', encoding='utf-8') as f:
        text = f.read()

    # 解析
    result = parse_deepseek_new_format(text, '22204-24010')

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