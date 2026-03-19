# -*- coding: utf-8 -*-
"""
测试标题提取的不同方法
"""
import re

# 从调试文件中提取的数据行
data_line = "22204-24010 22204-24010 Mass Air Flow Sensor MAF Meter compatible with Toyota Lexus Corolla Prius Tundra Sequoia Land Cruiser 2022-2025 22204-22010 22204-0A010"

fields = data_line.split()

print(f"字段数量: {len(fields)}")
print(f"\n所有字段:")
for i, f in enumerate(fields):
    print(f"  {i}: {f}")

print(f"\n" + "="*80)
print("测试不同的标题提取方法:")
print("="*80)

# 方法1: 提取 OE 号之后到 "compatible" 之间的内容
compatible_idx = -1
for i, field in enumerate(fields):
    if field.lower() == 'compatible':
        compatible_idx = i
        break

if compatible_idx > 1:
    title_fields = fields[2:compatible_idx]
    title = ' '.join(title_fields)
    print(f"\n方法1: OE号之后 -> 'compatible' 之前")
    print(f"  标题: {title}")

# 方法2: 提取第二个 OE 号之后到第一个年份之前
year_idx = -1
for i, field in enumerate(fields):
    if re.match(r'^\d{4}$', field):
        year_idx = i
        break

if year_idx > 2:
    title_fields = fields[2:year_idx]
    title = ' '.join(title_fields)
    print(f"\n方法2: 第二个OE号之后 -> 第一个年份之前")
    print(f"  标题: {title}")

# 方法3: 根据 DeepSeek 写作思路
# 标题格式: [OE号] + [产品名] + compatible with + [品牌] + [车型] + [年份] + 兼容型号
# 但表格中已经以 OE 号开头了

# 尝试：从第二个 OE 号开始，提取到年份，包含 "compatible with"
if compatible_idx > 2 and year_idx > compatible_idx:
    title_fields = fields[2:year_idx]
    title = ' '.join(title_fields)
    print(f"\n方法3: 第二个OE号之后 -> 年份之前（包含 compatible with）")
    print(f"  标题: {title}")

# 方法4: 从 OE 号开始，提取到第一个兼容OE号之前
# 兼容OE号通常在最后
oe_pattern = r'\d{4}-[A-Z0-9]+'
compat_oe_idx = -1
for i, field in enumerate(fields):
    if i > 2 and re.match(oe_pattern, field):
        # 不是第一个 OE 号
        compat_oe_idx = i
        break

if compat_oe_idx > 0:
    title_fields = fields[2:compat_oe_idx]
    title = ' '.join(title_fields)
    print(f"\n方法4: 第二个OE号之后 -> 兼容OE号之前")
    print(f"  标题: {title}")

print(f"\n" + "="*80)
print("期望的标题（根据写作思路说明）:")
print("="*80)
expected_title = "22204-24010 Mass Air Flow Sensor MAF Meter compatible with Toyota Lexus Corolla Prius Tundra Sequoia Land Cruiser 2022-2025 22204-22010 22204-0A010"
print(f"\n  {expected_title}")
