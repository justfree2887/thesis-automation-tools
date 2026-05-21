#!/usr/bin/env python3
"""
update_thesis_figures.py
综合更新论文图表：
1. 替换 Fig.3-1, 4-1, 5-1, 6-1 为新的应力-应变曲线（解决Issue 3&4）
2. 刷新所有柱状图为MPa单位版本（解决Issue 3）
"""

import os, re, json
from docx import Document
from lxml import etree

# ========== 路径配置 ==========
DOCX_PATH = '/Users/shiberlin/Desktop/毕业论文修改/师江柏毕业论文119.docx'
CHARTS_DIR = '/Users/shiberlin/Desktop/毕业论文/charts'
THESIS_CHARTS_DIR = os.path.join(CHARTS_DIR, 'thesis')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def get_text(elem):
    return ''.join(t.text or '' for t in elem.findall(f'.//{{{W}}}t')).strip()

def find_chart_file(pattern, directory=None):
    """按模式查找图表文件"""
    search_dir = directory or THESIS_CHARTS_DIR
    if not os.path.exists(search_dir):
        print(f"    [WARN] Directory not found: {search_dir}")
        return None
    for f in os.listdir(search_dir):
        if pattern in f and f.endswith('.png'):
            return os.path.join(search_dir, f)
    return None

def get_image_info(para_elem, doc_part):
    """获取段落中图片的 rId 和尺寸"""
    drawings = para_elem.findall(f'.//{{{W}}}drawing')
    for drawing in drawings:
        blips = drawing.findall(f'.//{{{A}}}blip')
        extents = drawing.findall(f'.//{{{A}}}ext')
        
        rId = None
        for blip in blips:
            rId = blip.get(f'{{{REL}}}embed')
            if rId and rId in doc_part.rels:
                break
        
        ext_info = None
        for ext in extents:
            cx = ext.get('cx')
            cy = ext.get('cy')
            if cx and cy:
                ext_info = (int(cx), int(cy))
                break
        
        return rId, ext_info
    return None, None

def replace_image_in_paragraph(para_elem, doc_part, new_image_path):
    """替换段落中的图片"""
    old_rId, old_ext = get_image_info(para_elem, doc_part)
    if not old_rId:
        return False
    
    # 获取新图片
    new_rId, new_image_part = doc_part.get_or_add_image(new_image_path)
    
    # 替换 blip embed
    blips = para_elem.findall(f'.//{{{A}}}blip')
    for blip in blips:
        if blip.get(f'{{{REL}}}embed') == old_rId:
            blip.set(f'{{{REL}}}embed', new_rId)
    
    # 更新图片名称
    cNvPrs = para_elem.findall(f'.//{{{A}}}cNvPr')
    for cNvPr in cNvPrs:
        base_name = os.path.splitext(os.path.basename(new_image_path))[0]
        cNvPr.set('name', base_name)
    
    return True

def replace_stress_strain_curves(doc):
    """替换第四章的应力-应变曲线图"""
    body = doc.element.body
    children = list(body)
    main_part = doc.part
    
    # 应力-应变曲线映射：图题关键词 → 新图表文件
    ss_map = {
        '图3-1': ('Fig3-1_initial_stress_strain.png', CHARTS_DIR),
        '图4-1': ('Fig4-1_uv_stress_strain.png', CHARTS_DIR),
        '图5-1': ('Fig5-1_hydro_stress_strain.png', CHARTS_DIR),
        '图6-1': ('Fig6-1_coupled_stress_strain.png', CHARTS_DIR),
    }
    
    replaced = 0
    for fig_label, (fname, directory) in ss_map.items():
        chart_path = find_chart_file(fname, directory)
        if not chart_path:
            print(f"  ✗ {fig_label}: 找不到 {fname}")
            continue
        
        # 查找图题段落
        found = False
        for i, child in enumerate(children):
            if child.tag != f'{{{W}}}p':
                continue
            text = get_text(child)
            # 严格匹配图题
            if text.startswith(fig_label + '  ') or text == fig_label:
                # 向前找图片段落
                for j in range(i - 1, max(i - 10, 0), -1):
                    prev = children[j]
                    if prev.tag != f'{{{W}}}p':
                        continue
                    prev_text = get_text(prev)
                    if prev.findall(f'.//{{{W}}}drawing'):
                        if replace_image_in_paragraph(prev, main_part, chart_path):
                            print(f"  ✓ {fig_label}: 替换为 {os.path.basename(chart_path)}")
                            replaced += 1
                            found = True
                        break
                    elif prev_text:
                        break
                break
        if not found:
            print(f"  ? {fig_label}: 未找到图题段落")
    
    return replaced

def refresh_bar_charts(doc):
    """刷新所有柱状图为MPa版本"""
    body = doc.element.body
    children = list(body)
    main_part = doc.part
    
    # 柱状图映射
    chart_map = {
        '图4-2': '第四章_拉伸_强度_分组柱状图',
        '图4-3': '第四章_拉伸_延伸率_分组柱状图',
        '图4-5': '第四章_弯曲_强度_分组柱状图',
        '图4-6': '第四章_弯曲_延伸率_分组柱状图',
        '图4-7': '第四章_拉伸_强度保留率_分组',
        '图4-8': '第四章_弯曲_强度保留率_分组',
        '图5-2': '第五章_拉伸_强度_分组柱状图',
        '图5-3': '第五章_拉伸_延伸率_分组柱状图',
        '图5-5': '第五章_弯曲_强度_分组柱状图',
        '图5-6': '第五章_弯曲_延伸率_分组柱状图',
        '图5-7': '第五章_拉伸_强度保留率_分组',
        '图5-8': '第五章_弯曲_强度保留率_分组',
    }
    
    replaced = 0
    skipped = 0
    
    for i, child in enumerate(children):
        if child.tag != f'{{{W}}}p':
            continue
        text = get_text(child)
        
        for fig_label, chart_pattern in chart_map.items():
            if not (text.startswith(fig_label + '  ') or text == fig_label):
                continue
            
            chart_path = find_chart_file(chart_pattern, THESIS_CHARTS_DIR)
            if not chart_path:
                print(f"  ✗ {fig_label}: 找不到 {chart_pattern}")
                skipped += 1
                continue
            
            # 向前找图片段落
            for j in range(i - 1, max(i - 10, 0), -1):
                prev = children[j]
                if prev.tag != f'{{{W}}}p':
                    continue
                prev_text = get_text(prev)
                if prev.findall(f'.//{{{W}}}drawing'):
                    if replace_image_in_paragraph(prev, main_part, chart_path):
                        print(f"  ✓ {fig_label}: 替换为 {os.path.basename(chart_path)}")
                        replaced += 1
                    break
                elif prev_text:
                    break
            break
    
    return replaced, skipped

def main():
    print("=" * 60)
    print("更新论文图表（解决Issue 3 & 4）")
    print("=" * 60)
    
    doc = Document(DOCX_PATH)
    
    # 1. 替换应力-应变曲线
    print("\n[1/2] 替换拉伸应力-应变曲线（Force→Stress, Displacement→Strain）...")
    ss_count = replace_stress_strain_curves(doc)
    
    # 2. 刷新柱状图（N→MPa）
    print("\n[2/2] 刷新柱状图（单位N→MPa）...")
    bar_count, bar_skipped = refresh_bar_charts(doc)
    
    # 保存
    total = ss_count + bar_count
    if total > 0:
        doc.save(DOCX_PATH)
        print(f"\n{'='*60}")
        print(f"完成！共替换 {total} 张图表")
        print(f"  应力-应变曲线: {ss_count} 张")
        print(f"  柱状图: {bar_count} 张 (跳过 {bar_skipped})")
        print(f"已保存到: {DOCX_PATH}")
        print(f"{'='*60}")
    else:
        print(f"\n未替换任何图表")

if __name__ == '__main__':
    main()
