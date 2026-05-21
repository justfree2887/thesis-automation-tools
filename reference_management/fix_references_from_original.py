#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从原始文件重新处理参考文献：
1. 扫描正文，按首次出现顺序确定引用编号映射
2. 用占位符算法替换正文中的引用标注（避免级联）
3. 重排参考文献列表
"""

import re
import copy
import json
from docx import Document
from docx.oxml.ns import qn

INPUT_FILE = '/Users/shiberlin/Desktop/毕业论文1.2.docx'
OUTPUT_FILE = '/Users/shiberlin/Desktop/毕业论文/师江柏毕业论文初稿.docx'
BACKUP_FILE = '/Users/shiberlin/Desktop/毕业论文/师江柏毕业论文1.2_backup.docx'

# ============================================================
# Step 1: 读取原始文件，扫描正文引用
# ============================================================
print("=" * 60)
print("Step 1: 扫描正文引用标注")
print("=" * 60)

doc = Document(INPUT_FILE)

# 找参考文献标题位置
ref_title_idx = None
for i, p in enumerate(doc.paragraphs):
    t = p.text.strip()
    if t == '参考文献':
        ref_title_idx = i
        break

print(f"参考文献标题在第 {ref_title_idx} 段")

# 提取参考文献列表的条目（ref_title_idx 之后的段落）
ref_entries = {}  # old_num -> (para_index, paragraph_element)
for i in range(ref_title_idx + 1, len(doc.paragraphs)):
    t = doc.paragraphs[i].text.strip()
    m = re.match(r'\[(\d+)\]', t)
    if m:
        old_num = int(m.group(1))
        ref_entries[old_num] = i
    elif t and not t.startswith('['):
        break  # 参考文献结束

print(f"参考文献条目数: {len(ref_entries)}")
print(f"参考文献编号: {sorted(ref_entries.keys())}")

# 扫描正文中的引用标注（参考文献标题之前的段落）
citations_in_order = []  # 按首次出现顺序的原始编号
seen = set()
citation_pattern = re.compile(r'\[(\d+)\]')

for i in range(ref_title_idx):
    p = doc.paragraphs[i]
    text = p.text
    for m in citation_pattern.finditer(text):
        num = int(m.group(1))
        if num in ref_entries and num not in seen:
            citations_in_order.append(num)
            seen.add(num)

print(f"正文中引用的不同文献数: {len(citations_in_order)}")
print(f"引用顺序: {citations_in_order}")

# 构建旧→新映射
old_to_new = {}
for new_num, old_num in enumerate(citations_in_order, start=1):
    old_to_new[old_num] = new_num

# 未被引用的文献排在末尾
uncited = sorted(set(ref_entries.keys()) - set(citations_in_order))
next_num = len(citations_in_order) + 1
for old_num in uncited:
    old_to_new[old_num] = next_num
    next_num += 1

print(f"\n旧→新编号映射:")
for old_num in sorted(old_to_new.keys()):
    new_num = old_to_new[old_num]
    marker = " (被引用)" if old_num in set(citations_in_order) else " (未引用)"
    print(f"  [{old_num}] -> [{new_num}]{marker}")

# ============================================================
# Step 2: 用占位符算法替换正文引用标注（避免级联）
# ============================================================
print("\n" + "=" * 60)
print("Step 2: 占位符替换正文引用标注")
print("=" * 60)

# Phase 1: [old_N] -> [__REF_N__]（占位符，不可能与任何实际编号冲突）
placeholder_pattern = re.compile(r'\[(\d+)\]')
total_replaced = 0

for i in range(ref_title_idx):
    p = doc.paragraphs[i]
    for run in p.runs:
        text = run.text
        new_text = placeholder_pattern.sub(lambda m: f'[__REF_{m.group(1)}__]', text)
        if new_text != text:
            run.text = new_text
            count = len(placeholder_pattern.findall(text))
            total_replaced += count

print(f"Phase 1: 替换了 {total_replaced} 处引用为占位符")

# Phase 2: [__REF_old_N__] -> [new_N]
placeholder2_pattern = re.compile(r'\[__REF_(\d+)__\]')
total_final = 0

for i in range(ref_title_idx):
    p = doc.paragraphs[i]
    for run in p.runs:
        text = run.text
        new_text = placeholder2_pattern.sub(
            lambda m: f'[{old_to_new.get(int(m.group(1)), int(m.group(1)))}]',
            text
        )
        if new_text != text:
            run.text = new_text
            count = len(placeholder2_pattern.findall(text))
            total_final += count

print(f"Phase 2: 替换了 {total_final} 处占位符为新编号")

# 验证正文中的引用
body_citations = set()
for i in range(ref_title_idx):
    text = doc.paragraphs[i].text
    for m in re.finditer(r'\[(\d+)\]', text):
        num = int(m.group(1))
        body_citations.add(num)

print(f"正文中出现的引用编号: {sorted(body_citations)}")
expected = set(range(1, len(citations_in_order) + 1))
if body_citations == expected:
    print("✓ 正文引用编号验证通过！")
else:
    missing = expected - body_citations
    extra = body_citations - expected
    if missing:
        print(f"✗ 缺少编号: {sorted(missing)}")
    if extra:
        print(f"✗ 多余编号: {sorted(extra)}")

# ============================================================
# Step 3: 重排参考文献列表
# ============================================================
print("\n" + "=" * 60)
print("Step 3: 重排参考文献列表")
print("=" * 60)

# 构建新顺序：先是被引用的（按首次出现顺序），再是未被引用的（按原编号排序）
new_order = list(citations_in_order) + uncited
print(f"新参考文献顺序（前10个）: {new_order[:10]}...")

# 获取参考文献段落的 XML 元素
ref_para_elements = []
ref_para_texts = {}
for i in range(ref_title_idx + 1, ref_title_idx + 1 + len(ref_entries)):
    p = doc.paragraphs[i]
    ref_para_elements.append(p._element)
    # 也提取后续可能不属于参考文献的段落
    t = p.text.strip()
    m = re.match(r'\[(\d+)\]', t)
    if m:
        ref_para_texts[int(m.group(1))] = p._element

print(f"收集到 {len(ref_para_texts)} 个参考文献段落的XML元素")

# 为每个旧编号的新条目创建段落文本
def make_ref_text(new_num, old_num):
    """根据旧编号的参考文献内容，生成新编号的条目文本"""
    old_para = doc.paragraphs[ref_entries[old_num]]
    old_text = old_para.text.strip()
    # 替换编号 [old_num] -> [new_num]
    new_text = re.sub(r'^\[\d+\]', f'[{new_num}]', old_text)
    return new_text

# 清空原有参考文献段落，按新顺序重新写入
# 先把所有参考文献元素从文档中移除
ref_parent = ref_para_elements[0].getparent()
ref_elements_to_remove = []
for elem in ref_para_elements:
    ref_elements_to_remove.append(elem)

for elem in ref_elements_to_remove:
    ref_parent.remove(elem)

# 找到参考文献标题元素
ref_title_elem = doc.paragraphs[ref_title_idx]._element

# 按新顺序创建新的参考文献段落
from docx.oxml import OxmlElement
from lxml import etree

for new_num, old_num in enumerate(new_order, start=1):
    new_text = make_ref_text(new_num, old_num)
    # 创建新的段落元素
    new_p = OxmlElement('w:p')
    new_r = OxmlElement('w:r')
    new_t = OxmlElement('w:t')
    new_t.text = new_text
    new_t.set(qn('xml:space'), 'preserve')
    new_r.append(new_t)
    new_p.append(new_r)
    # 插入到参考文献标题之后
    ref_title_elem.addnext(new_p)
    ref_title_elem = new_p  # 下一个插在刚插入的之后

print("参考文献列表已重排")

# ============================================================
# Step 4: 保存结果
# ============================================================
print("\n" + "=" * 60)
print("Step 4: 保存结果")
print("=" * 60)

# 先备份
import shutil
shutil.copy2(INPUT_FILE, BACKUP_FILE)
print(f"原始文件已备份到: {BACKUP_FILE}")

# 保存修改后的文件
doc.save(OUTPUT_FILE)
print(f"修改后文件已保存到: {OUTPUT_FILE}")

# ============================================================
# Step 5: 最终验证
# ============================================================
print("\n" + "=" * 60)
print("Step 5: 最终验证")
print("=" * 60)

doc2 = Document(OUTPUT_FILE)
total_paras = len(doc2.paragraphs)
print(f"总段落数: {total_paras}")

# 找参考文献标题
ref_title2 = None
for i, p in enumerate(doc2.paragraphs):
    if p.text.strip() == '参考文献':
        ref_title2 = i
        break

# 检查正文引用
body_cits = set()
body_cit_count = 0
for i in range(ref_title2):
    for m in re.finditer(r'\[(\d+)\]', doc2.paragraphs[i].text):
        body_cits.add(int(m.group(1)))
        body_cit_count += 1

print(f"正文中引用标注总数: {body_cit_count}")
print(f"正文中不同引用编号: {sorted(body_cits)}")
print(f"引用文献数: {len(body_cits)}")

# 检查参考文献列表
ref_nums = []
for i in range(ref_title2 + 1, len(doc2.paragraphs)):
    t = doc2.paragraphs[i].text.strip()
    m = re.match(r'\[(\d+)\]', t)
    if m:
        ref_nums.append(int(m.group(1)))
    elif t:
        break

print(f"参考文献列表条目数: {len(ref_nums)}")
print(f"参考文献编号: {ref_nums}")

# 检查编号连续性
expected_ref = list(range(1, len(ref_nums) + 1))
if ref_nums == expected_ref:
    print("✓ 参考文献编号连续 [1]~[{}]".format(len(ref_nums)))
else:
    print("✗ 参考文献编号不连续!")
    print(f"  期望: {expected_ref}")
    print(f"  实际: {ref_nums}")

# 检查正文引用是否都是有效编号
invalid = body_cits - set(ref_nums)
if not invalid:
    print("✓ 正文中的所有引用编号在参考文献列表中都有对应条目")
else:
    print(f"✗ 正文中有无效引用编号: {sorted(invalid)}")

# 保存映射数据供参考
mapping_data = {
    'citations_in_order': citations_in_order,
    'uncited': uncited,
    'old_to_new': {str(k): v for k, v in old_to_new.items()},
    'new_order': new_order,
    'body_citation_count': body_cit_count,
    'ref_count': len(ref_nums)
}
with open('/tmp/ref_mapping_final.json', 'w') as f:
    json.dump(mapping_data, f, ensure_ascii=False, indent=2)
print(f"\n映射数据已保存到 /tmp/ref_mapping_final.json")
