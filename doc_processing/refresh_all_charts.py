#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
refresh_all_charts.py
替换文档中所有图表为新字体版本（四号黑体+Times New Roman）
覆盖全部40张图：第三章到第七章的曲线图、柱状图、FTIR图
"""

import os, glob, json
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

# ── 路径 ──
BASE_DIR = '/Users/shiberlin/Desktop/毕业论文'
DOCX_PATH = glob.glob('/Users/shiberlin/Desktop/*1.2.docx')[0]
THESIS_DIR = os.path.join(BASE_DIR, 'charts', 'thesis')
CHARTS_ROOT = os.path.join(BASE_DIR, 'charts')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def get_text(elem):
    return ''.join(t.text or '' for t in elem.findall(f'.//{{{W}}}t')).strip()


def find_chart_file(pattern, search_dirs):
    """在指定目录中按关键词查找图表文件"""
    for d in search_dirs:
        if not os.path.exists(d):
            continue
        for f in os.listdir(d):
            if pattern in f and f.endswith('.png'):
                path = os.path.join(d, f)
                if os.path.exists(path):
                    return path
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
    print("替换文档中所有图表（四号字体版本）")
    print("=" * 60)

    doc = Document(DOCX_PATH)
    body = doc.element.body
    children = list(body)
    main_part = doc.part

    # ── 图题 → 图表文件映射 ──
    # 格式: 图题关键词 → (文件名关键词, 搜索目录列表)
    # 对于第四章/第五章，分组柱状图在 thesis/ 目录
    # 对于第三章/第六章，单条件柱状图在 thesis/ 目录
    # FTIR 在 charts/ 根目录

    search_dirs_thesis = [THESIS_DIR]
    search_dirs_ftir = [CHARTS_ROOT]
    search_dirs_all = [THESIS_DIR, CHARTS_ROOT]

    chart_map = {
        # ===== 第三章 =====
        '图3-1': '第三章_拉伸_曲线_全',
        '图3-2': '第三章_拉伸_强度柱状图',
        '图3-3': '第三章_拉伸_位移柱状图',    # 文件名是位移，图题是延伸率
        '图3-4': '第三章_弯曲_曲线_全',
        '图3-5': '第三章_弯曲_强度柱状图',
        '图3-6': '第三章_弯曲_位移柱状图',    # 文件名是位移，图题是延伸率
        '图3-7': '第七章_初始_红外光谱',      # FTIR（初始）

        # ===== 第四章 =====
        '图4-1': '第四章_拉伸_曲线_全',
        '图4-2': '第四章_拉伸_强度_分组柱状图',
        '图4-3': '第四章_拉伸_延伸率_分组柱状图',
        '图4-4': '第四章_弯曲_曲线_全',
        '图4-5': '第四章_弯曲_强度_分组柱状图',
        '图4-6': '第四章_弯曲_延伸率_分组柱状图',
        '图4-7': '第四章_拉伸_强度保留率_分组',
        '图4-8': '第四章_弯曲_强度保留率_分组',
        '图4-9': '第七章_紫外_红外光谱',      # FTIR（紫外）

        # ===== 第五章 =====
        '图5-1': '第五章_拉伸_曲线_全',
        '图5-2': '第五章_拉伸_强度_分组柱状图',
        '图5-3': '第五章_拉伸_延伸率_分组柱状图',
        '图5-4': '第五章_弯曲_曲线_全',
        '图5-5': '第五章_弯曲_强度_分组柱状图',
        '图5-6': '第五章_弯曲_延伸率_分组柱状图',
        '图5-7': '第五章_拉伸_强度保留率_分组',
        '图5-8': '第五章_弯曲_强度保留率_分组',
        '图5-9': '第七章_湿热_红外光谱',      # FTIR（湿热）

        # ===== 第六章 =====
        '图6-1': '第六章_拉伸_曲线_全',
        '图6-2': '第六章_拉伸_强度柱状图',
        '图6-3': '第六章_拉伸_位移柱状图',    # 文件名是位移
        '图6-4': '第六章_弯曲_曲线_全',
        '图6-5': '第六章_弯曲_强度柱状图',
        '图6-6': '第六章_弯曲_位移柱状图',    # 文件名是位移
        '图6-7': '第六章_拉伸_强度保留率',
        '图6-8': '第六章_弯曲_强度保留率',
        '图6-9': '第七章_紫外湿热_红外光谱',  # FTIR（紫外湿热）

        # ===== 第七章 =====
        '图7-1': '第七章_拉伸强度_分组柱状图',
        '图7-2': '第七章_弯曲强度_分组柱状图',
        '图7-3': '第七章_冲击强度_分组柱状图',
        '图7-4': '第七章_强度保留率_分组',
        '图7-5': '第七章_弯曲模量_分组柱状图',
        '图7-6': '第七章_冲击韧性_分组柱状图',
    }

    replaced = 0
    skipped = 0
    not_found_in_doc = 0

    for i, child in enumerate(children):
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)

        for fig_label, chart_pattern in chart_map.items():
            # 图题格式: "图X-Y  " (图号后跟空格)
            if not (text.startswith(fig_label + '  ') or text == fig_label):
                continue

            # 确定搜索目录
            if '红外光谱' in chart_pattern:
                dirs = search_dirs_ftir
            else:
                dirs = search_dirs_thesis

            # 查找图表文件
            chart_path = find_chart_file(chart_pattern, dirs)
            if not chart_path:
                print(f"  跳过 {fig_label}: 找不到 {chart_pattern}")
                skipped += 1
                continue

            # 从图题向前找图片段落
            img_para_idx = None
            for j in range(i - 1, max(i - 10, 0), -1):
                prev = children[j]
                if prev.tag != f'{{{W}}}p':
                    # Check for sdt with drawing
                    sdt_drawings = prev.findall(f'.//{{{W}}}drawing')
                    if sdt_drawings:
                        img_para_idx = j
                        break
                    continue
                prev_text = get_text(prev)
                has_drawing = prev.findall(f'.//{{{W}}}drawing')
                if has_drawing:
                    img_para_idx = j
                    break
                elif prev_text:
                    break

            if img_para_idx is None:
                print(f"  跳过 {fig_label}: 找不到前面的图片")
                skipped += 1
                continue

            # 获取旧图片 rId
            old_rId = get_image_rel_id(children[img_para_idx], main_part)
            if not old_rId:
                print(f"  跳过 {fig_label}: 图片无 rId")
                skipped += 1
                continue

            # 插入新图片（获取新 rId）
            new_rId, new_image_part = main_part.get_or_add_image(chart_path)

            # 替换 rId
            drawing_elem = children[img_para_idx]
            blips = drawing_elem.findall(f'.//{{{A}}}blip')
            for blip in blips:
                embed_attr = blip.get(f'{{{REL}}}embed')
                if embed_attr == old_rId:
                    blip.set(f'{{{REL}}}embed', new_rId)

            print(f"  替换 {fig_label}: {os.path.basename(chart_path)}")
            replaced += 1
            break

    # 保存
    if replaced > 0:
        doc.save(DOCX_PATH)
        print(f"\n已保存: {os.path.basename(DOCX_PATH)}")
        print(f"替换: {replaced} 张, 跳过: {skipped} 张")
    else:
        print(f"\n未替换任何图表")

    # 验证
    doc2 = Document(DOCX_PATH)
    body2 = doc2.element.body
    WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    inline_count = len(body2.findall(f'.//{{{WP_NS}}}inline'))
    anchor_count = len(body2.findall(f'.//{{{WP_NS}}}anchor'))
    print(f"文档图片总数: {inline_count + anchor_count}")

    with open('/tmp/refresh_all_log.json', 'w', encoding='utf-8') as f:
        json.dump({"replaced": replaced, "skipped": skipped, "saved": replaced > 0}, f, ensure_ascii=False)


if __name__ == '__main__':
    main()
