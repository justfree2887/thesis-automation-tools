#!/usr/bin/env python3
"""
refresh_charts_in_doc.py
将毕业论文_修改版.docx中已经嵌入的分组柱状图替换为新生成的版本。
由于docx图片是以二进制嵌入的，需要：
1. 删除旧图片对应的relationship
2. 重新插入新图片
"""

import os, re, glob, json
from docx import Document
from docx.shared import Cm
from PIL import Image
from lxml import etree

BASE_DIR = '/Users/shiberlin/Desktop/毕业论文'
import glob as _glob
_docx_candidates = _glob.glob('/Users/shiberlin/Desktop/*1.2.docx')
DOCX_PATH = _docx_candidates[0] if _docx_candidates else os.path.join(BASE_DIR, '毕业论文1.2.docx')
# 使用glob找到charts/thesis目录
_chart_dirs = _glob.glob('/Users/shiberlin/Desktop/*/charts/thesis')
CHART_DIR = _chart_dirs[0] if _chart_dirs else os.path.join(BASE_DIR, 'charts', 'thesis')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def get_text(elem):
    return ''.join(t.text or '' for t in elem.findall(f'.//{{{W}}}t')).strip()


def find_chart(pattern):
    """按模式查找图表文件"""
    for f in os.listdir(CHART_DIR):
        if pattern in f and f.endswith('.png'):
            path = os.path.join(CHART_DIR, f)
            if os.path.exists(path):
                return path
    # Debug
    import sys
    print(f"    [DEBUG] CHART_DIR={CHART_DIR}", file=sys.stderr)
    print(f"    [DEBUG] CHART_DIR exists={os.path.exists(CHART_DIR)}", file=sys.stderr)
    print(f"    [DEBUG] pattern={repr(pattern)}", file=sys.stderr)
    matches = [f for f in os.listdir(CHART_DIR) if '强度_分组' in f]
    print(f"    [DEBUG] files with 强度_分组: {matches[:5]}", file=sys.stderr)
    return None


def get_image_rel_id(para_elem, doc_part):
    """获取段落中第一个图片的 rId"""
    drawings = para_elem.findall(f'.//{{{W}}}drawing')
    for drawing in drawings:
        blips = drawing.findall(f'.//{{{A}}}blip')
        for blip in blips:
            rId = blip.get(f'{{{REL}}}embed')
            if rId and rId in doc_part.rels:
                return rId
    return None


def main():
    print("=" * 60)
    print("刷新文档中的分组柱状图")
    print("=" * 60)

    doc = Document(DOCX_PATH)
    body = doc.element.body
    children = list(body)
    main_part = doc.part

    # 需要替换的图表映射：图题关键词 → 新图表文件模式
    chart_map = {
        # 第四章分组图
        '图4-2': ('第四章_拉伸_强度_分组柱状图', '不同紫外老化时间下复合材料的拉伸强度'),
        '图4-3': ('第四章_拉伸_延伸率_分组柱状图', '不同紫外老化时间下复合材料的拉伸断裂延伸率'),
        '图4-5': ('第四章_弯曲_强度_分组柱状图', '不同紫外老化时间下复合材料的弯曲强度'),
        '图4-6': ('第四章_弯曲_延伸率_分组柱状图', '不同紫外老化时间下复合材料的弯曲断裂延伸率'),
        '图4-7': ('第四章_拉伸_强度保留率_分组', '不同紫外老化时间下拉伸强度保留率'),
        '图4-8': ('第四章_弯曲_强度保留率_分组', '不同紫外老化时间下弯曲强度保留率'),
        # 第五章分组图
        '图5-2': ('第五章_拉伸_强度_分组柱状图', '不同湿热老化时间下复合材料的拉伸强度'),
        '图5-3': ('第五章_拉伸_延伸率_分组柱状图', '不同湿热老化时间下复合材料的拉伸断裂延伸率'),
        '图5-5': ('第五章_弯曲_强度_分组柱状图', '不同湿热老化时间下复合材料的弯曲强度'),
        '图5-6': ('第五章_弯曲_延伸率_分组柱状图', '不同湿热老化时间下复合材料的弯曲断裂延伸率'),
        '图5-7': ('第五章_拉伸_强度保留率_分组', '不同湿热老化时间下拉伸强度保留率'),
        '图5-8': ('第五章_弯曲_强度保留率_分组', '不同湿热老化时间下弯曲强度保留率'),
    }

    replaced = 0
    skipped = 0

    for i, child in enumerate(children):
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)

        # 检查是否是图题段落（格式："图X-Y  " 图号后有空格，排除正文引用如"图4-2对比了"）
        import re
        for fig_label, (chart_pattern, _) in chart_map.items():
            # 图题格式: "图X-Y  " (图号后跟两个空格)，或者 "图X-Y  xxx"
            # 排除 "图4-2对比了..." 这种正文引用
            if not (text.startswith(fig_label + '  ') or text == fig_label):
                continue

            # 找到前面的图片段落
            chart_path = find_chart(chart_pattern)
            if not chart_path:
                print(f"  跳过 {fig_label}: 找不到图表文件 {chart_pattern}")
                skipped += 1
                continue

            # 从当前图题向前找图片段落（跳过空段落）
            img_para_idx = None
            for j in range(i - 1, max(i - 10, 0), -1):
                prev = children[j]
                if prev.tag != f'{{{W}}}p':
                    continue
                prev_text = get_text(prev)
                has_drawing = prev.findall(f'.//{{{W}}}drawing')
                if has_drawing:
                    img_para_idx = j
                    break
                elif prev_text:  # 有文字但不是图片，停止
                    break

            if img_para_idx is None:
                print(f"  跳过 {fig_label}: 找不到前面的图片段落")
                skipped += 1
                continue

            # 获取旧图片的 rId
            old_rId = get_image_rel_id(children[img_para_idx], main_part)
            if not old_rId:
                print(f"  跳过 {fig_label}: 图片段落无 rId")
                skipped += 1
                continue

            # 获取旧图片的尺寸信息（用于保持大小一致）
            old_ext = None
            drawings = children[img_para_idx].findall(f'.//{{{W}}}drawing')
            for drawing in drawings:
                extents = drawing.findall(f'.//{{{A}}}ext')
                for ext in extents:
                    cx = ext.get('cx')
                    cy = ext.get('cy')
                    if cx and cy:
                        old_ext = (int(cx), int(cy))
                        break

            # 插入新图片（获取新的 rId）
            new_rId, new_image_part = main_part.get_or_add_image(chart_path)

            # 替换所有引用旧 rId 的地方为新 rId
            drawing_elem = children[img_para_idx]
            blips = drawing_elem.findall(f'.//{{{A}}}blip')
            for blip in blips:
                embed_attr = blip.get(f'{{{REL}}}embed')
                if embed_attr == old_rId:
                    blip.set(f'{{{REL}}}embed', new_rId)

            # 更新 docPr name
            cNvPrs = drawing_elem.findall(f'.//{{{A}}}cNvPr')
            for cNvPr in cNvPrs:
                old_name = cNvPr.get('name', '')
                if old_name:
                    # 更新名称
                    base_name = os.path.splitext(os.path.basename(chart_path))[0]
                    cNvPr.set('name', base_name)

            print(f"  ✓ 替换 {fig_label}: {os.path.basename(chart_path)}")
            replaced += 1
            break  # 一个段落只匹配一次

    if replaced > 0:
        doc.save(DOCX_PATH)
        with open('/tmp/refresh_log.json', 'w', encoding='utf-8') as lf:
            json.dump({"replaced": replaced, "skipped": skipped, "saved": True}, lf)
        print(f"\n已保存到: {DOCX_PATH}")
        print(f"替换: {replaced} 张, 跳过: {skipped} 张")
    else:
        with open('/tmp/refresh_log.json', 'w', encoding='utf-8') as lf:
            json.dump({"replaced": 0, "skipped": skipped, "saved": False}, lf)
        print(f"\n未替换任何图表")

    # 验证
    doc2 = Document(DOCX_PATH)
    body2 = doc2.element.body
    WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    inline_count = len(body2.findall(f'.//{{{WP_NS}}}inline'))
    anchor_count = len(body2.findall(f'.//{{{WP_NS}}}anchor'))
    print(f"\n文档图片总数: {inline_count + anchor_count} (inline={inline_count}, anchor={anchor_count})")


if __name__ == '__main__':
    main()
