"""
毕业论文1.1.docx 修改脚本：
1. 将所有图片环绕类型从 anchor 改为 inline（嵌入型）
2. 将吸湿试验内容从第八章移动到第七章（7.7之后）
"""

from lxml import etree
from docx import Document
import shutil, os

doc_path = '/Users/shiberlin/Desktop/毕业论文1.1.docx'
backup_path = '/Users/shiberlin/Desktop/毕业论文1.1_backup.docx'

# 备份
shutil.copy2(doc_path, backup_path)
print(f'已备份到: {backup_path}')

doc = Document(doc_path)
body = doc.element.body

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
WP14 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
Q_WP = '{' + WP + '}'
Q_WP14 = '{' + WP14 + '}'

# ============================================================
# Task 1: 将所有 anchor 图片改为 inline（嵌入型）
# ============================================================
anchors = body.findall(f'.//{{{W}}}drawing/{{{WP}}}anchor')
print(f'\n找到 {len(anchors)} 个 anchor 图片，开始转换...')

# 需要删除的 anchor 属性
anchor_only_attrs = {'simplePos', 'relativeHeight', 'behindDoc', 'locked', 'layoutInCell', 'allowOverlap'}
# 需要删除的子元素 tag（本地名）
remove_child_tags = {'simplePos', 'positionH', 'positionV', 'sizeRelH', 'sizeRelV'}

converted = 0
for anchor in anchors:
    # 1. 删除 anchor 特有属性
    for attr in list(anchor.attrib.keys()):
        local = attr.split('}')[-1] if '}' in attr else attr
        if local in anchor_only_attrs:
            del anchor.attrib[attr]

    # 2. 设置 dist 为 0（嵌入型无间距）
    for attr in ['distT', 'distB', 'distL', 'distR']:
        if attr in anchor.attrib:
            anchor.attrib[attr] = '0'

    # 3. 删除不需要的子元素
    for child in list(anchor):
        tag_local = etree.QName(child.tag).localname
        tag_ns = etree.QName(child.tag).namespace
        if tag_local in remove_child_tags:
            anchor.remove(child)
        elif tag_local.startswith('wrap'):
            anchor.remove(child)

    # 4. 重命名标签: anchor → inline
    anchor.tag = Q_WP + 'inline'
    converted += 1

print(f'已转换 {converted} 个 anchor → inline')

# 验证：确保没有遗留的 anchor
remaining_anchors = body.findall(f'.//{{{W}}}drawing/{{{WP}}}anchor')
total_inline = body.findall(f'.//{{{W}}}drawing/{{{WP}}}inline')
print(f'验证: 剩余 anchor={len(remaining_anchors)}, 总 inline={len(total_inline)}')

# ============================================================
# Task 2: 将吸湿试验内容移动到第七章（7.7之后）
# ============================================================
print('\n开始移动吸湿试验内容...')

elements = list(body)

# 吸湿试验内容范围: body[394] 到 body[401]
# body[393] = 第八章 结论与展望 标题（保留不移动）
# body[394] = 吸湿试验方法段落
# body[395] = 吸湿试验结果段落
# body[396] = 吸湿试验机理段落
# body[397] = 表7-6标题
# body[398] = 表7-6 (table)
# body[399] = 图7-11标题
# body[400] = 图7-11英文标题
# body[401] = 图7-11图片

# 插入位置: body[392] 之后（7.7最后一个图片的英文标题body[391]之后的空段落）
to_move = elements[394:402]  # body[394] ~ body[401] 共9个元素
ref_elem = elements[392]  # 在此元素之后插入

print(f'  移动 {len(to_move)} 个元素 (body[394]-body[401])')
print(f'  插入位置: body[392] 之后')

# 验证边界
print(f'  body[392] 文本: {"".join(t.text or "" for t in ref_elem.findall(f".//{{{W}}}t"))[:50] if ref_elem.findall(f".//{{{W}}}t") else "(empty)"}')
print(f'  to_move[0] 文本: {"".join(t.text or "" for t in to_move[0].findall(f".//{{{W}}}t"))[:50]}')
print(f'  to_move[-1] tag: {etree.QName(to_move[-1].tag).localname}')

# 先从原位置移除
for elem in to_move:
    body.remove(elem)

# 找到 ref_elem 的新索引（移除后索引可能变化）
new_ref_idx = list(body).index(ref_elem)

# 按顺序插入到 ref_elem 之后
for i, elem in enumerate(to_move):
    body.insert(new_ref_idx + 1 + i, elem)

print(f'  移动完成!')

# 验证移动后的结构
print('\n验证移动后的结构:')
for idx in range(388, 410):
    child = body[idx]
    tag = etree.QName(child.tag).localname
    texts = child.findall(f'.//{{{W}}}t')
    p_text = ''.join(t.text or '' for t in texts).strip()[:60]
    has_img = child.findall(f'.//{{{W}}}drawing')
    is_tbl = tag == 'tbl'
    extra = f' [IMG:{len(has_img)}]' if has_img else ''
    extra = ' [TABLE]' if is_tbl else extra
    if p_text or extra:
        with open('/tmp/move_verify.txt', 'a') as f:
            f.write(f'  body[{idx}] ({tag}){extra} | {p_text}\n')

with open('/tmp/move_verify.txt', 'a') as f:
    f.write('\n')

print(f'验证信息已写入 /tmp/move_verify.txt')

# ============================================================
# 保存
# ============================================================
doc.save(doc_path)
print(f'\n已保存到: {doc_path}')
print('完成!')
