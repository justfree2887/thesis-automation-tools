#!/usr/bin/env python3
"""
cleanup_residual_content.py
清理毕业论文_修改版.docx中残余的旧章节内容。

残余内容范围：新第六章结尾到结论章之间的旧章节内容（约body[321]-[498]），
包含16个单个曲线图（图X-Y-Z格式）和重复的章节小结文本。
"""

import os, re
from docx import Document

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

BASE_DIR = '/Users/shiberlin/Desktop/毕业论文'
DOCX_PATH = os.path.join(BASE_DIR, '毕业论文_修改版.docx')


def get_text(elem):
    return ''.join(t.text or '' for t in elem.findall(f'.//{{{W}}}t')).strip()


def main():
    print("=" * 60)
    print("清理毕业论文_修改版.docx 中残余的旧章节内容")
    print("=" * 60)

    doc = Document(DOCX_PATH)
    body = doc.element.body
    children = list(body)
    print(f"文档共 {len(children)} 个 body 子元素")

    # ===== 第一步：找到所有单个曲线图标题的位置 =====
    single_curve_positions = []
    for i, child in enumerate(children):
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        if re.match(r'图\d+-\d+-\d+', text):
            single_curve_positions.append(i)

    print(f"找到 {len(single_curve_positions)} 个单个曲线图标题")
    if not single_curve_positions:
        print("无需清理，退出。")
        return

    # ===== 第二步：找到结论章标题（包含"第七章"且不在目录部分） =====
    conclusion_idx = None
    # 从后半部分查找（body[300]之后）
    for i in range(max(300, single_curve_positions[-1] + 1), len(children)):
        child = children[i]
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        if '第七章' in text and len(text) < 40:
            conclusion_idx = i
            print(f"结论章标题: body[{i}]: {text}")
            break

    if conclusion_idx is None:
        print("错误：未找到结论章标题！")
        return

    # ===== 第三步：确定删除范围 =====
    # 从第一个单个曲线图向前找，找到残余内容的真正起始位置
    first_curve = single_curve_positions[0]
    delete_start = first_curve

    # 向前查找，找到旧的章节小结段落的起始
    # 跳过单个曲线图前面的空段落和图片段落
    for i in range(first_curve - 1, max(0, first_curve - 60), -1):
        child = children[i]
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        # 找到 "4.3 力学性能保留率与本章小结" 或类似的小节标题
        if re.match(r'[456]\.3', text):
            delete_start = i
            break
        # 找到 "本章小结" 或 "本章研究了" 段落
        if text.startswith('（') and '纳米' in text:
            delete_start = i
            break

    # 再向前查找 "本章研究了..." 段落
    for i in range(delete_start - 1, max(0, delete_start - 30), -1):
        child = children[i]
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        if '本章' in text and '研究了' in text:
            delete_start = i
            break

    # 再向前检查是否有 "X.3 小结" 标题
    for i in range(delete_start - 1, max(0, delete_start - 30), -1):
        child = children[i]
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        if '小结' in text or re.match(r'[456]\.\d', text):
            if '保留率' in text or '小结' in text:
                delete_start = i
                break

    # 再向前找重复的 "本章小结" 段落块
    for i in range(delete_start - 1, max(0, delete_start - 50), -1):
        child = children[i]
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        if text.startswith('（1）') or text.startswith('（2）') or text.startswith('（3）'):
            delete_start = i
            break

    # 向后：删除到结论章之前
    delete_end = conclusion_idx - 1

    # 向后扩展：包含 conclusion_idx 前面的空段落
    while delete_end > delete_start:
        child = children[delete_end]
        if child.tag == f'{{{W}}}p':
            text = get_text(child)
            has_drawing = child.findall(f'.//{{{W}}}drawing')
            if text == '' and not has_drawing:
                delete_end -= 1
                continue
        break

    print(f"\n删除范围: body[{delete_start}]-[{delete_end}]")
    print(f"删除元素数: {delete_end - delete_start + 1}")

    # 验证删除范围确实包含所有单个曲线图
    curves_in_range = sum(1 for p in single_curve_positions if delete_start <= p <= delete_end)
    print(f"范围内的单个曲线图: {curves_in_range}/{len(single_curve_positions)}")

    if curves_in_range < len(single_curve_positions):
        print("警告：部分单个曲线图不在删除范围内！")
        return

    # 预览删除范围
    print("\n删除范围预览（前5个非空段落）:")
    shown = 0
    for i in range(delete_start, min(delete_end + 1, delete_start + 50)):
        child = children[i]
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        has_img = child.findall(f'.//{{{W}}}drawing')
        if text or has_img:
            label = f'[IMG] {text[:60]}' if has_img else text[:60]
            print(f"  [{i}] {label}")
            shown += 1
            if shown >= 5:
                break

    # ===== 第四步：执行删除 =====
    to_remove = children[delete_start:delete_end + 1]
    for elem in to_remove:
        body.remove(elem)

    # ===== 第五步：保存并验证 =====
    doc.save(DOCX_PATH)
    print(f"\n已保存到: {DOCX_PATH}")

    # 验证
    doc2 = Document(DOCX_PATH)
    body2 = doc2.element.body
    children2 = list(body2)

    remaining_singles = 0
    for child in children2:
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        if re.match(r'图\d+-\d+-\d+', text):
            remaining_singles += 1
            print(f"  残余: {text[:60]}")

    print(f"\n验证结果:")
    print(f"  文档元素总数: {len(children2)} (之前 {len(children)})")
    print(f"  残余单个曲线图标题: {remaining_singles} 个")

    if remaining_singles == 0:
        print("\n✅ 清理完成！所有单个曲线图已成功删除。")
    else:
        print(f"\n⚠️ 仍有 {remaining_singles} 个单个曲线图未删除，需要进一步处理。")

    # 显示最终章节结构
    print("\n最终章节结构:")
    for i, child in enumerate(children2):
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        if ('第' in text and '章' in text and len(text) < 50):
            print(f"  body[{i}]: {text[:60]}")


if __name__ == '__main__':
    main()
