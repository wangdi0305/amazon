# -*- coding: utf-8 -*-
"""
简单测试：直接从文本中提取各部分
"""
import re

def extract_parts(text):
    """直接从文本中提取各部分，不使用正则"""
    parts = {
        'title': '',
        'ad': '',
        'ae': '',
        'af': '',
        'ag': '',
        'ah': '',
        'ai': '',
        'keywords': [],
    }

    # 按行分割
    lines = text.split('\n')
    
    # 查找标题行
    for line in lines:
        if line.startswith('22204-24010') and 'Mass Air Flow Sensor' in line:
            # 标题在 "Vehicle Fitment:" 之前
            parts['title'] = line.split('Vehicle Fitment:')[0].strip()
            break

    # 查找 AD 部分（Vehicle Fitment:）
    ad_start = False
    ad_content = []
    for line in lines:
        if 'Vehicle Fitment:' in line:
            ad_start = True
            continue
        if ad_start and 'Direct fit for Toyota/Lexus' in line:
            break
        if ad_start:
            ad_content.append(line)
    
    parts['ad'] = ' '.join(ad_content).strip()

    # 查找 AE 部分（Direct fit for Toyota/Lexus）
    ae_start = False
    ae_content = []
    for line in lines:
        if 'Direct fit for Toyota/Lexus' in line:
            ae_start = True
            continue
        if ae_start and 'Replaces OEM Numbers:' in line:
            break
        if ae_start:
            ae_content.append(line)
    
    parts['ae'] = ' '.join(ae_content).strip()

    # 查找 AF 部分（Replaces OEM Numbers:）
    af_start = False
    af_content = []
    for line in lines:
        if 'Replaces OEM Numbers:' in line:
            af_start = True
            continue
        if af_start and 'Measures intake air volume' in line:
            break
        if af_start:
            af_content.append(line)
    
    parts['af'] = ' '.join(af_content).strip()

    # 查找 AG 部分（Measures intake air volume）
    ag_start = False
    ag_content = []
    for line in lines:
        if 'Measures intake air volume' in line:
            ag_start = True
            continue
        if ag_start and 'Plug & play' in line:
            break
        if ag_start:
            ag_content.append(line)
    
    parts['ag'] = ' '.join(ag_content).strip()

    # 查找 AH 部分（Plug & play）
    ah_start = False
    ah_content = []
    for line in lines:
        if 'Plug & play' in line:
            ah_start = True
            continue
        if ah_start and 'Meets Toyota/Lexus OE specs' in line:
            break
        if ah_start:
            ah_content.append(line)
    
    parts['ah'] = ' '.join(ah_content).strip()

    # 查找 AI 部分（Meets Toyota/Lexus OE specs）
    ai_start = False
    ai_content = []
    for line in lines:
        if 'Meets Toyota/Lexus OE specs' in line:
            ai_start = True
            continue
        if ai_start and '1-year warranty' in line:
            break
        if ai_start:
            ai_content.append(line)
    
    parts['ai'] = ' '.join(ai_content).strip()

    # 查找关键词
    for line in lines:
        if line.startswith('22204-24010 Mass Air Flow Sensor'):
            keywords_line = line
            parts['keywords'] = keywords_line.split()[1:]  # 跳过 OE 号
            break

    return parts


# 测试
def main():
    print("=" * 80)
    print("  简单提取测试")
    print("=" * 80)

    # 读取文件
    with open('/mnt/c/Users/Administrator/Desktop/debug_final_row7.txt', 'r', encoding='utf-8') as f:
        text = f.read()

    # 提取
    parts = extract_parts(text)

    print(f"\n提取结果：")
    print(f"  标题: {parts['title'][:100]}...")
    print(f"  AD: {parts['ad'][:100]}...")
    print(f"  AE: {parts['ae'][:100]}...")
    print(f"  AF: {parts['af'][:100]}...")
    print(f"  AG: {parts['ag'][:100]}...")
    print(f"  AH: {parts['ah'][:100]}...")
    print(f"  AI: {parts['ai'][:100]}...")
    print(f"  关键词: {parts['keywords']}")


if __name__ == "__main__":
    main()