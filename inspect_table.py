# -*- coding: utf-8 -*-
"""
仅检查页面 DOM 结构 - 找 DeepSeek 回复中的 table 元素
在已有对话（有回复）的 DeepSeek 页面上运行
"""
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
import time, re

opts = Options()
opts.add_argument('--start-maximized')
opts.add_experimental_option('detach', True)
opts.add_experimental_option('excludeSwitches', ['enable-automation'])

driver = webdriver.Edge(options=opts)
driver.get("https://chat.deepseek.com")
time.sleep(5)

input(">>> 确保 DeepSeek 已有回复。按 Enter 检查...")

# 1. 找所有 table
tables = driver.find_elements(By.TAG_NAME, 'table')
print(f"\n=== 页面共 {len(tables)} 个 <table> ===")

for t_idx, table in enumerate(tables):
    rows = table.find_elements(By.TAG_NAME, 'tr')
    print(f"\n--- Table {t_idx} ({len(rows)} rows) ---")
    
    for r_idx, row in enumerate(rows[:3]):
        cells = row.find_elements(By.TAG_NAME, 'td')
        ths = row.find_elements(By.TAG_NAME, 'th')
        all_cells = ths + cells
        
        row_data = []
        for c_idx, cell in enumerate(all_cells):
            text = cell.text.strip().replace('\n', ' ')
            if len(text) > 100:
                text = text[:100] + '...'
            row_data.append(f"[{c_idx}]{text}")
        
        if row_data:
            print(f"  Row {r_idx} ({len(all_cells)} cells): {' | '.join(row_data)}")

# 2. 找最近消息容器（不同 DeepSeek 版本的 DOM 结构不同）
print(f"\n=== 消息容器检查 ===")
msg_selectors = [
    '[class*="message"]',
    '[class*="response"]',
    '[class*="assistant"]',
    '[class*="content-wrap"]',
    '[class*="markdown"]',
    '[class*="reply"]',
    '[class*="chat-item"]',
    '[class*="dialogue"]',
    '[class*="bubble"]',
]
for sel in msg_selectors:
    try:
        els = driver.find_elements(By.CSS_SELECTOR, sel)
        if els and len(els) >= 1:
            last = els[-1]
            tables_in = last.find_elements(By.TAG_NAME, 'table')
            text = last.text.strip().replace('\n', ' ')[:200]
            if tables_in:
                print(f"  {sel}: {len(els)}个, 最后1个含 {len(tables_in)} table")
                print(f"    text: {text}")
    except:
        pass

# 3. 检查 DeepSeek 特有的 DOM 结构
print(f"\n=== DeepSeek DOM 结构探测 ===")
# 打印 body 的直接子元素结构
body_children = driver.execute_script("""
    var children = document.body.children;
    var result = [];
    for (var i = 0; i < Math.min(children.length, 30); i++) {
        var el = children[i];
        var tag = el.tagName.toLowerCase();
        var id = el.id ? '#' + el.id : '';
        var cls = el.className ? '.' + el.className.toString().split(' ').slice(0,2).join('.') : '';
        result.push(tag + id + cls);
    }
    return result;
""")
print(f"  body 直接子元素: {body_children[:15]}")

# 检查是否有 DSW* 开头的 class（DeepSeek 特有）
dsw_els = driver.execute_script("""
    var all = document.querySelectorAll('[class*="dsw"]');
    var tags = {};
    for (var i = 0; i < all.length; i++) {
        var tag = all[i].tagName + '.' + all[i].className.toString().split(' ')[0];
        tags[tag] = (tags[tag] || 0) + 1;
    }
    return Object.entries(tags).sort(function(a,b){return b[1]-a[1]}).slice(0,20);
""")
print(f"  DeepSeek 组件 (dsw*): {dsw_els[:15]}")

# 4. 找最后一个 table 并保存其完整 HTML
if tables:
    last_table = tables[-1]
    html = last_table.get_attribute('outerHTML')
    with open(r"C:\Users\Administrator\Desktop\deepseek_listing_tool\last_table.html", 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n=== 最后一个 table HTML 已保存到 last_table.html ({len(html)} bytes) ===")
    
    # 提取 td 内容
    trs = last_table.find_elements(By.TAG_NAME, 'tr')
    if trs:
        last_tr = trs[-1] if len(trs) > 1 else trs[0]
        tds = last_tr.find_elements(By.TAG_NAME, 'td')
        print(f"\n最后一个数据行 ({len(tds)} 个 td):")
        for i, td in enumerate(tds):
            text = td.text.strip().replace('\n', ' ')
            print(f"  [{i}] ({len(text)}c) {text[:120]}{'...' if len(text)>120 else ''}")

input("\n按 Enter 关闭...")