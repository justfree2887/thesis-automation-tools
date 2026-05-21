#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从原始文件重新处理参考文献（改进版）：
1. 扫描正文，按首次出现顺序确定引用编号映射
2. 处理跨 run 的引用标注，用占位符算法替换
3. 重排参考文献列表
"""

import re
import copy
import json
import shutil
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

INPUT_FILE = '/Users/shiberlin/Desktop/毕业论文1.2.docx'
OUTPUT_FILE = '/Users/shiberlin/Desktop/毕业论文/师江柏毕业论文初稿.docx'
BACKUP_FILE = '/Users/shiberlin/Desktop/毕业论文/师江柏毕业论文1.2_backup.docx'

doc = Document(INPUT_FILE)

# ============================================================
# Step 1: 找到参考文献标题位置
# ============================================================
ref_title_idx = None
for i, p in enumerate(doc.paragraphs):
    if p.text.strip() == '参考文献':
        ref_title_idx = i
        break

print(f"参考文献标题在第 {ref_title_idx} 段")

# ============================================================
# Step 2: 提取参考文献条目（先保存文本，避免后续操作影响）
# ============================================================
ref_entries = {}  # old_num -> para_index
ref_texts = {}    # old_num -> full text
for i in range(ref_title_idx + 1, len(doc.paragraphs)):
    t = doc.paragraphs[i].text.strip()
    m = re.match(r'\[(\d+)\]', t)
    if m:
        old_num = int(m.group(1))
        ref_entries[old_num] = i
        ref_texts[old_num] = t
    elif t and not t.startswith('['):
        break

print(f"参考文献条目数: {len(ref_entries)}, 编号: {sorted(ref_entries.keys())}")

# ============================================================
# Step 3: 扫描正文引用（处理跨 run 情况）
# ============================================================
def get_para_full_text(para):
    """获取段落完整文本"""
    return ''.join(run.text for run in para.runs)

def find_all_citations_in_para(para):
    """在段落中查找所有 [N] 引用，返回 (start, end, num) 列表"""
    full_text = get_para_full_text(para)
    results = []
    for m in re.finditer(r'\[(\d+)\]', full_text):
        num = int(m.group(1))
        results.append((m.start(), m.end(), num, full_text[m.start():m.end()]))
    return results, full_text

# 扫描正文，记录首次出现顺序
citations_in_order = []
seen = set()

for i in range(ref_title_idx):
    p = doc.paragraphs[i]
    cits, _ = find_all_citations_in_para(p)
    for _, _, num, _ in cits:
        if num in ref_entries and num not in seen:
            citations_in_order.append(num)
            seen.add(num)

print(f"正文引用的不同文献数: {len(citations_in_order)}")
print(f"引用顺序: {citations_in_order}")

# 构建旧→新映射
old_to_new = {}
for new_num, old_num in enumerate(citations_in_order, start=1):
    old_to_new[old_num] = new_num

uncited = sorted(set(ref_entries.keys()) - set(citations_in_order))
next_num = len(citations_in_order) + 1
for old_num in uncited:
    old_to_new[old_num] = next_num
    next_num += 1

print(f"\n旧→新映射 (被引用的):")
for old_num in citations_in_order:
    print(f"  [{old_num}] -> [{old_to_new[old_num]}]")
print(f"未被引用的文献 ({len(uncited)}个): {[old_to_new[n] for n in uncited]}")

# ============================================================
# Step 4: 用占位符算法替换正文引用标注
# ============================================================
print("\n--- Phase 1: [N] -> [__REF_N__] 占位符 ---")

# 对每个正文段落：
# 1. 获取完整文本
# 2. 替换所有 [N] 为占位符
# 3. 将替换后的文本重新分配到 runs 中
def replace_citations_in_para(para, replace_func):
    """
    在段落中替换引用标注，正确处理跨 run 的情况。
    replace_func: (num) -> new_text
    """
    if not para.runs:
        return 0

    full_text = get_para_full_text(para)
    # 执行替换
    new_text = re.sub(r'\[(\d+)\]',
                      lambda m: replace_func(int(m.group(1))),
                      full_text)

    if new_text == full_text:
        return 0

    # 将新文本分配到 runs 中
    # 策略：保留第一个 run 的格式，将所有文本放入第一个 run，清空其余 runs
    if len(para.runs) == 1:
        para.runs[0].text = new_text
    else:
        # 保留第一个 run 的格式，把所有文本放进去
        # 清空其他 runs 的文本
        first_run = para.runs[0]
        first_run.text = new_text
        for run in para.runs[1:]:
            run.text = ''

    # 计算替换了多少处
    old_matches = list(re.finditer(r'\[(\d+)\]', full_text))
    return len(old_matches)

# Phase 1: 用占位符替换
total_ph1 = 0
for i in range(ref_title_idx):
    p = doc.paragraphs[i]
    count = replace_citations_in_para(p, lambda num: f'[__REF_{num}__]')
    total_ph1 += count

print(f"Phase 1: 替换了 {total_ph1} 处引用为占位符")

# Phase 2: 占位符 -> 新编号
print("--- Phase 2: [__REF_N__] -> [new_N] ---")
total_ph2 = 0

for i in range(ref_title_idx):
    p = doc.paragraphs[i]
    full_text = get_para_full_text(p)

    # 检查是否有占位符
    if '[__REF_' not in full_text:
        continue

    new_text = re.sub(r'\[__REF_(\d+)__\]',
                      lambda m: f'[{old_to_new.get(int(m.group(1)), int(m.group(1)))}]',
                      full_text)

    if new_text != full_text:
        if len(p.runs) == 1:
            p.runs[0].text = new_text
        else:
            p.runs[0].text = new_text
            for run in p.runs[1:]:
                run.text = ''
        count = len(re.findall(r'\[__REF_\d+__\]', full_text))
        total_ph2 += count

print(f"Phase 2: 替换了 {total_ph2} 处占位符为新编号")

# ============================================================
# Step 5: 验证正文引用
# ============================================================
print("\n--- 验证正文引用 ---")
body_cits = set()
body_cit_count = 0
for i in range(ref_title_idx):
    p = doc.paragraphs[i]
    full_text = get_para_full_text(p)
    for m in re.finditer(r'\[(\d+)\]', full_text):
        body_cits.add(int(m.group(1)))
        body_cit_count += 1

print(f"正文引用标注总数: {body_cit_count}")
print(f"不同引用编号: {sorted(body_cits)}")
expected_cits = set(range(1, len(citations_in_order) + 1))
missing_cits = expected_cits - body_cits
extra_cits = body_cits - expected_cits
if not missing_cits and not extra_cits:
    print("✓ 正文引用编号完全正确！")
else:
    if missing_cits:
        print(f"✗ 缺少: {sorted(missing_cits)}")
    if extra_cits:
        print(f"✗ 多余: {sorted(extra_cits)}")

# ============================================================
# Step 6: 重排参考文献列表
# ============================================================
print("\n--- 重排参考文献列表 ---")

new_order = list(citations_in_order) + uncited

# 找到参考文献标题的 XML 元素
ref_title_elem = doc.paragraphs[ref_title_idx]._element

# 收集所有参考文献段落元素
ref_elems = []
for old_num in sorted(ref_entries.keys()):
    idx = ref_entries[old_num]
    ref_elems.append((old_num, doc.paragraphs[idx]._element))

# 从文档中移除所有参考文献段落
for _, elem in ref_elems:
    parent = elem.getparent()
    if parent is not None:
        parent.remove(elem)

# 按新顺序创建新的参考文献段落（使用简单段落，保持基本格式）
# 复制参考文献标题的段落格式作为模板
from docx.oxml import parse_xml

insert_after = ref_title_elem
for new_num, old_num in enumerate(new_order, start=1):
    new_text = re.sub(r'^\[\d+\]', f'[{new_num}]', ref_texts[old_num])

    # 创建段落元素
    new_p = OxmlElement('w:p')
    # 添加段落属性（与标题后原段落保持一致的基本格式）
    pPr = OxmlElement('w:pPr')
    new_p.append(pPr)

    # 创建 run
    new_r = OxmlElement('w:r')
    new_t = OxmlElement('w:t')
    new_t.text = new_text
    new_t.set(qn('xml:space'), 'preserve')
    new_r.append(new_t)
    new_p.append(new_r)

    # 插入
    insert_after.addnext(new_p)
    insert_after = new_p

print(f"已重新排列 {len(new_order)} 条参考文献")

# ============================================================
# Step 7: 保存
# ============================================================
print("\n--- 保存文件 ---")
shutil.copy2(INPUT_FILE, BACKUP_FILE)
print(f"原始文件已备份: {BACKUP_FILE}")

doc.save(OUTPUT_FILE)
print(f"修改后文件已保存: {OUTPUT_FILE}")

# ============================================================
# Step 8: 最终验证
# ============================================================
print("\n" + "=" * 50)
print("最终验证")
print("=" * 50)

doc2 = Document(OUTPUT_FILE)

ref_title2 = None
for i, p in enumerate(doc2.paragraphs):
    if p.text.strip() == '参考文献':
        ref_title2 = i
        break

# 检查正文
body_c2 = set()
for i in range(ref_title2):
    for m in re.finditer(r'\[(\d+)\]', doc2.paragraphs[i].text):
        body_c2.add(int(m.group(1)))

print(f"正文中引用编号: {sorted(body_c2)}")
print(f"被引文献数: {len(body_c2)}")

# 检查参考文献列表
ref_nums2 = []
for i in range(ref_title2 + 1, len(doc2.paragraphs)):
    t = doc2.paragraphs[i].text.strip()
    m = re.match(r'\[(\d+)\]', t)
    if m:
        ref_nums2.append(int(m.group(1)))
    elif t:
        break

print(f"参考文献条目数: {len(ref_nums2)}")
print(f"编号范围: [{ref_nums2[0]}]~[{ref_nums2[-1]}]")

expected_refs = list(range(1, len(ref_nums2) + 1))
if ref_nums2 == expected_refs:
    print("✓ 参考文献编号连续！")
else:
    print(f"✗ 参考文献编号不连续: {ref_nums2}")

invalid = body_c2 - set(ref_nums2)
if not invalid:
    print("✓ 所有正文引用在参考文献列表中都有对应条目")
else:
    print(f"✗ 无效引用: {sorted(invalid)}")

# 显示前5条新参考文献
print(f"\n新参考文献列表（前5条）:")
for i in range(min(5, len(doc2.paragraphs) - ref_title2 - 1)):
    p = doc2.paragraphs[ref_title2 + 1 + i]
    t = p.text.strip()
    if t.startswith('['):
        print(f"  {t[:80]}...")

# 保存映射
mapping_data = {
    'citations_in_order': citations_in_order,
    'uncited': uncited,
    'old_to_new': {str(k): v for k, v in old_to_new.items()},
    'new_order': new_order,
    'body_citation_count': body_cit_count,
    'ref_count': len(ref_nums2),
    'verified': len(missing_cits) == 0 and len(extra_cits) == 0
}
with open('/tmp/ref_mapping_final.json', 'w') as f:
    json.dump(mapping_data, f, ensure_ascii=False, indent=2)

print("\n完成！")
