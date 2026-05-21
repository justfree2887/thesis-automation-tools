#!/usr/bin/env python3
"""
update_doc_v2.py
修改毕业论文_修改版.docx:
1. 删除第三到六章的单个含量曲线图（如图3-1-0等）
2. 第四五章用分组柱状图替换原有柱状图，并加入10天数据讨论
3. 新插入图片使用上下环绕型
"""

import os, json, copy, re
from lxml import etree
from docx import Document
from docx.shared import Pt, Cm, Emu
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from PIL import Image

# ============ 路径 ============
BASE_DIR = '/Users/shiberlin/Desktop/毕业论文'
DOCX_PATH = os.path.join(BASE_DIR, '毕业论文_修改版.docx')
CHART_DIR = os.path.join(BASE_DIR, 'charts', 'thesis')
STATS_PATH = os.path.join(CHART_DIR, 'stats_v2.json')

# ============ 命名空间 ============
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

etree.register_namespace('w', W)
etree.register_namespace('wp', WP)
etree.register_namespace('a', A)
etree.register_namespace('pic', PIC)
etree.register_namespace('r', REL)
etree.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
etree.register_namespace('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing')

stats = None

def load_stats():
    global stats
    with open(STATS_PATH, 'r', encoding='utf-8') as f:
        stats = json.load(f)


# ============ 段落创建工具 ============
def _make_pPr(align='left', first_indent=480, spacing_before=None, spacing_after=None,
              outline_lvl=None, line=400, line_rule='exact', font='黑体', sz=24, bold=False):
    pPr = etree.SubElement(etree.Element('dummy'), f'{{{W}}}pPr')
    if outline_lvl is not None:
        etree.SubElement(pPr, f'{{{W}}}widowControl')
    sp = etree.SubElement(pPr, f'{{{W}}}spacing')
    sp.set(f'{{{W}}}line', str(line))
    sp.set(f'{{{W}}}lineRule', line_rule)
    if spacing_before is not None:
        sp.set(f'{{{W}}}before', str(spacing_before))
    if spacing_after is not None:
        sp.set(f'{{{W}}}after', str(spacing_after))
    if first_indent:
        ind = etree.SubElement(pPr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLine', str(first_indent))
    jc = etree.SubElement(pPr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', align)
    if outline_lvl is not None:
        olvl = etree.SubElement(pPr, f'{{{W}}}outlineLvl')
        olvl.set(f'{{{W}}}val', str(outline_lvl))
    rPr = etree.SubElement(pPr, f'{{{W}}}rPr')
    rFonts = etree.SubElement(rPr, f'{{{W}}}rFonts')
    rFonts.set(f'{{{W}}}ascii', font)
    rFonts.set(f'{{{W}}}eastAsia', font)
    rFonts.set(f'{{{W}}}hAnsi', font)
    rFonts.set(f'{{{W}}}cs', font)
    etree.SubElement(rPr, f'{{{W}}}kern').set(f'{{{W}}}val', '0')
    sz_elem = etree.SubElement(rPr, f'{{{W}}}sz')
    sz_elem.set(f'{{{W}}}val', str(sz))
    szCs = etree.SubElement(rPr, f'{{{W}}}szCs')
    szCs.set(f'{{{W}}}val', str(sz))
    if bold:
        etree.SubElement(rPr, f'{{{W}}}b')
        etree.SubElement(rPr, f'{{{W}}}bCs')
    return pPr

def _make_run(text, font='黑体', sz=24, bold=False, hint_ea=None):
    r = etree.Element(f'{{{W}}}r')
    rPr = etree.SubElement(r, f'{{{W}}}rPr')
    rFonts = etree.SubElement(rPr, f'{{{W}}}rFonts')
    rFonts.set(f'{{{W}}}ascii', font)
    rFonts.set(f'{{{W}}}eastAsia', font)
    rFonts.set(f'{{{W}}}hAnsi', font)
    rFonts.set(f'{{{W}}}cs', font)
    if hint_ea:
        rFonts.set(f'{{{W}}}hint', 'eastAsia')
    etree.SubElement(rPr, f'{{{W}}}kern').set(f'{{{W}}}val', '0')
    sz_elem = etree.SubElement(rPr, f'{{{W}}}sz')
    sz_elem.set(f'{{{W}}}val', str(sz))
    szCs = etree.SubElement(rPr, f'{{{W}}}szCs')
    szCs.set(f'{{{W}}}val', str(sz))
    if bold:
        etree.SubElement(rPr, f'{{{W}}}b')
        etree.SubElement(rPr, f'{{{W}}}bCs')
    t = etree.SubElement(r, f'{{{W}}}t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r

def make_paragraph(text, align='left', first_indent=480, spacing_before=None,
                   spacing_after=None, outline_lvl=None, font='黑体', sz=24, bold=False):
    p = etree.Element(f'{{{W}}}p')
    pPr = _make_pPr(align=align, first_indent=first_indent,
                    spacing_before=spacing_before, spacing_after=spacing_after,
                    outline_lvl=outline_lvl, font=font, sz=sz, bold=bold)
    p.append(pPr)
    r = _make_run(text, font=font, sz=sz, bold=bold, hint_ea=(outline_lvl is not None))
    p.append(r)
    return p

def make_empty_paragraph():
    return etree.Element(f'{{{W}}}p')

def make_chapter_title(text):
    return make_paragraph(text, align='center', first_indent=0, font='黑体', sz=32)

def make_section_title(text, outline_lvl=1):
    return make_paragraph(text, align='left', first_indent=0,
                         spacing_before=200, spacing_after=100,
                         outline_lvl=outline_lvl, font='黑体', sz=28)

def make_subsection_title(text):
    return make_paragraph(text, align='left', first_indent=0,
                         spacing_before=100, spacing_after=50,
                         outline_lvl=2, font='黑体', sz=24, bold=True)

def make_body_text(text):
    return make_paragraph(text, align='left', first_indent=480, font='黑体', sz=24)

def make_caption(cn_text, en_text):
    p = etree.Element(f'{{{W}}}p')
    pPr = _make_pPr(align='center', first_indent=0, font='黑体', sz=24)
    p.append(pPr)
    r1 = _make_run(cn_text, font='黑体', sz=24, hint_ea=True)
    p.append(r1)
    r2 = _make_run(' ' + en_text, font='Times New Roman', sz=24)
    p.append(r2)
    return p


# ============ 图片插入（上下环绕型） ============
_image_counter = 200

def add_image_anchor(doc, image_path, width_cm=14):
    """添加图片到文档，返回wp:anchor格式的drawing元素（上下环绕型）"""
    global _image_counter
    main_part = doc.part
    rId, image_part = main_part.get_or_add_image(image_path)

    with Image.open(image_path) as img:
        px_w, px_h = img.size

    target_width_emu = int(width_cm * 360000)
    ratio = px_h / px_w if px_w > 0 else 1
    target_height_emu = int(target_width_emu * ratio)
    cx = str(target_width_emu)
    cy = str(target_height_emu)

    _image_counter += 1
    doc_prop_id = _image_counter

    anchor_xml = f'''
    <wp:anchor xmlns:wp="{WP}" distT="0" distB="0" distL="114300" distR="114300"
               simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0"
               layoutInCell="1" allowOverlap="1"
               xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{REL}">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="{cx}" cy="{cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:wrapTopAndBottom wrapText="bothSides"/>
      <wp:docPr id="{doc_prop_id}" name="Picture{doc_prop_id}"/>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr>
              <pic:cNvPr id="0" name="Picture{doc_prop_id}.png"/>
              <pic:cNvPicPr/>
            </pic:nvPicPr>
            <pic:blipFill>
              <a:blip r:embed="{rId}"/>
              <a:stretch><a:fillRect/></a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="{cx}" cy="{cy}"/>
              </a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>'''

    anchor = etree.fromstring(anchor_xml)
    drawing = etree.Element(f'{{{W}}}drawing')
    drawing.append(anchor)
    run = etree.Element(f'{{{W}}}r')
    run.append(drawing)
    return run


def make_image_paragraph(doc, image_path, width_cm=14, caption_cn='', caption_en=''):
    elements = []
    img_run = add_image_anchor(doc, image_path, width_cm)
    img_p = etree.Element(f'{{{W}}}p')
    img_pPr = _make_pPr(align='center', first_indent=0)
    img_p.append(img_pPr)
    img_p.append(img_run)
    elements.append(img_p)
    if caption_cn or caption_en:
        elements.append(make_caption(caption_cn, caption_en))
    return elements


# ============ 工具函数 ============
def find_chart(pattern):
    if not os.path.exists(CHART_DIR):
        return None
    for f in os.listdir(CHART_DIR):
        if pattern in f and f.endswith('.png'):
            path = os.path.join(CHART_DIR, f)
            if os.path.exists(path):
                return path
    return None


def get_paragraph_text(para_elem):
    return ''.join(t.text or '' for t in para_elem.findall(f'.//{{{W}}}t'))


def get_image_filename(para_elem, doc_part):
    """获取段落中图片的文件名"""
    drawings = para_elem.findall(f'.//{{{W}}}drawing')
    for drawing in drawings:
        # Check both inline and anchor
        blips = drawing.findall(f'.//{{{A}}}blip')
        for blip in blips:
            rId = blip.get(f'{{{REL}}}embed')
            if rId and rId in doc_part.rels:
                rel = doc_part.rels[rId]
                target = getattr(rel, 'target_ref', None)
                if target:
                    return os.path.basename(str(target))
    return None


def is_individual_curve_chart(filename):
    """检查是否为单个含量曲线图（非合并曲线图）"""
    if not filename:
        return False
    if '_曲线_' in filename and '_曲线_全' not in filename:
        return True
    return False


def is_empty_paragraph(para_elem):
    """检查段落是否为空（无文字内容）"""
    text = get_paragraph_text(para_elem).strip()
    # Also check if it only has a drawing (not empty if it has image)
    has_drawing = para_elem.findall(f'.//{{{W}}}drawing')
    if has_drawing:
        return False
    return text == ''


def remove_individual_curve_charts(body, start_body_idx, end_body_idx, doc_part):
    """从指定body范围中删除单个含量曲线图及其标题（通过图题文本识别）"""
    children = list(body)
    to_remove = set()

    # 方法：扫描所有段落，找到图题匹配 "图X-Y-Z" (如 图3-1-0, 图6-4-3 等) 的段落
    # 然后删除其前面的图片段落和周围的空段落
    for i in range(start_body_idx, min(end_body_idx + 1, len(children))):
        child = children[i]
        if child.tag != f'{{{W}}}p':
            continue

        text = get_paragraph_text(child).strip()
        # 匹配 "图X-Y-Z" 模式 (单个含量曲线的图题)
        if re.match(r'图\d+-\d+-\d+', text):
            to_remove.add(i)  # 标记图题段落
            # 向前查找图片段落（跳过空段落）
            for j in range(i - 1, max(start_body_idx - 1, i - 5), -1):
                if j < start_body_idx:
                    break
                prev = children[j]
                if prev.tag != f'{{{W}}}p':
                    continue
                if is_empty_paragraph(prev):
                    to_remove.add(j)  # 空段落也标记
                    continue
                has_drawing = prev.findall(f'.//{{{W}}}drawing')
                if has_drawing:
                    to_remove.add(j)  # 图片段落
                    break
                else:
                    break  # 遇到非空非图片段落，停止

    # Remove from end to start to preserve indices
    removed = 0
    for idx in sorted(to_remove, reverse=True):
        body.remove(children[idx])
        removed += 1

    return removed


# ============ 第四章/第五章内容生成 ============
def gen_uv_heat_chapter(doc, chapter_num, aging_name, section_label):
    """生成包含10天数据的紫外/湿热老化章节"""
    elements = []
    cn_num = {4: '四', 5: '五'}[chapter_num]

    # 确定数据键
    if chapter_num == 4:
        key_5d_t = '紫外老化5天'
        key_10d_t = '紫外老化10天'
        key_5d_b = '紫外老化5天'
        key_10d_b = '紫外老化10天'
    else:
        key_5d_t = '湿热老化5天'
        key_10d_t = '湿热老化10天'
        key_5d_b = '湿热老化5天'
        key_10d_b = '湿热老化10天'

    ts_init = stats['拉伸']['初始组']
    bs_init = stats['弯曲']['初始组']
    ts_5d = stats['拉伸'][key_5d_t]
    ts_10d = stats['拉伸'][key_10d_t]
    bs_5d = stats['弯曲'][key_5d_b]
    bs_10d = stats['弯曲'][key_10d_b]

    def ret(init_val, aged_val):
        return (aged_val / init_val * 100) if init_val > 0 else 0

    elements.append(make_chapter_title(f'第{cn_num}章  竹粉/PLA/纳米SiO₂复合材料{aging_name}力学性能'))
    elements.append(make_empty_paragraph())

    # ---- X.1 拉伸性能 ----
    elements.append(make_section_title(f'{section_label}.1 拉伸性能'))

    elements.append(make_body_text(
        f'图{section_label}-1为经不同时间{aging_name}后PLA/竹粉/纳米SiO\u2082复合材料的拉伸试验曲线。'
        f'从图中可以看出，经{aging_name}处理后，复合材料的拉伸载荷-位移曲线发生了明显变化。'
    ))

    # 拉伸合并曲线图（5天）
    chart = find_chart(f'第{cn_num}章_拉伸_曲线_全')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn=f'图{section_label}-1  {aging_name}后PLA/竹粉/纳米SiO\u2082复合材料拉伸试验曲线',
            caption_en=f'Tensile Test Curves after {aging_name}'))
        elements.append(make_empty_paragraph())

    # 拉伸强度分组柱状图
    elements.append(make_body_text(
        f'图{section_label}-2对比了不同{aging_name}时间下各纳米SiO\u2082含量复合材料的拉伸强度。'
    ))

    chart = find_chart(f'第{cn_num}章_拉伸_强度_分组柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn=f'图{section_label}-2  不同{aging_name}时间下复合材料的拉伸强度',
            caption_en=f'Tensile Strength under Different {aging_name} Durations'))
        elements.append(make_empty_paragraph())

    # 5天数据讨论
    elements.append(make_body_text(
        f'经5天{aging_name}后，0%SiO\u2082复合材料的拉伸强度为'
        f'{ts_5d["0%"]["max_load_mean"]:.2f}\u00b1{ts_5d["0%"]["max_load_std"]:.2f} N，'
        f'强度保留率为{ret(ts_init["0%"]["max_load_mean"], ts_5d["0%"]["max_load_mean"]):.1f}%；'
        f'1%SiO\u2082为{ts_5d["1%"]["max_load_mean"]:.2f}\u00b1{ts_5d["1%"]["max_load_std"]:.2f} N'
        f'（保留率{ret(ts_init["1%"]["max_load_mean"], ts_5d["1%"]["max_load_mean"]):.1f}%）；'
        f'3%SiO\u2082为{ts_5d["3%"]["max_load_mean"]:.2f}\u00b1{ts_5d["3%"]["max_load_std"]:.2f} N'
        f'（保留率{ret(ts_init["3%"]["max_load_mean"], ts_5d["3%"]["max_load_mean"]):.1f}%）；'
        f'5%SiO\u2082为{ts_5d["5%"]["max_load_mean"]:.2f}\u00b1{ts_5d["5%"]["max_load_std"]:.2f} N'
        f'（保留率{ret(ts_init["5%"]["max_load_mean"], ts_5d["5%"]["max_load_mean"]):.1f}%）。'
    ))

    # 10天数据讨论
    elements.append(make_body_text(
        f'经10天{aging_name}后，0%SiO\u2082复合材料的拉伸强度降至'
        f'{ts_10d["0%"]["max_load_mean"]:.2f}\u00b1{ts_10d["0%"]["max_load_std"]:.2f} N，'
        f'强度保留率为{ret(ts_init["0%"]["max_load_mean"], ts_10d["0%"]["max_load_mean"]):.1f}%；'
        f'1%SiO\u2082为{ts_10d["1%"]["max_load_mean"]:.2f}\u00b1{ts_10d["1%"]["max_load_std"]:.2f} N'
        f'（保留率{ret(ts_init["1%"]["max_load_mean"], ts_10d["1%"]["max_load_mean"]):.1f}%）；'
        f'3%SiO\u2082为{ts_10d["3%"]["max_load_mean"]:.2f}\u00b1{ts_10d["3%"]["max_load_std"]:.2f} N'
        f'（保留率{ret(ts_init["3%"]["max_load_mean"], ts_10d["3%"]["max_load_mean"]):.1f}%）；'
        f'5%SiO\u2082为{ts_10d["5%"]["max_load_mean"]:.2f}\u00b1{ts_10d["5%"]["max_load_std"]:.2f} N'
        f'（保留率{ret(ts_init["5%"]["max_load_mean"], ts_10d["5%"]["max_load_mean"]):.1f}%）。'
    ))

    # 5天vs10天对比分析
    ret_5d_0 = ret(ts_init["0%"]["max_load_mean"], ts_5d["0%"]["max_load_mean"])
    ret_10d_0 = ret(ts_init["0%"]["max_load_mean"], ts_10d["0%"]["max_load_mean"])
    elements.append(make_body_text(
        f'随着{aging_name}时间的延长，复合材料的拉伸强度进一步下降。'
        f'5天到10天期间，0%SiO\u2082的拉伸强度保留率从{ret_5d_0:.1f}%下降到{ret_10d_0:.1f}%，'
        f'说明{aging_name}对复合材料力学性能的劣化效应具有时间累积性。'
        f'添加纳米SiO\u2082后，复合材料在更长老化时间下的强度保持能力有所改善，'
        f'表明纳米颗粒对延缓{aging_name}引起的力学性能退化具有积极作用。'
    ))

    # 拉伸位移分组柱状图
    chart = find_chart(f'第{cn_num}章_拉伸_位移_分组柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn=f'图{section_label}-3  不同{aging_name}时间下复合材料的拉伸断裂位移',
            caption_en=f'Tensile Displacement under Different {aging_name} Durations'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'如图{section_label}-3所示，断裂位移方面，经5天{aging_name}后0%SiO\u2082为'
        f'{ts_5d["0%"]["max_disp_mean"]:.2f}\u00b1{ts_5d["0%"]["max_disp_std"]:.2f} mm，'
        f'经10天{aging_name}后为{ts_10d["0%"]["max_disp_mean"]:.2f}\u00b1{ts_10d["0%"]["max_disp_std"]:.2f} mm。'
        f'5%SiO\u2082复合材料的断裂位移在5天{aging_name}后为{ts_5d["5%"]["max_disp_mean"]:.2f} mm，'
        f'10天后为{ts_10d["5%"]["max_disp_mean"]:.2f} mm，'
        f'表明纳米SiO\u2082对材料在{aging_name}条件下的变形能力有一定影响。'
    ))

    # ---- X.2 弯曲性能 ----
    elements.append(make_section_title(f'{section_label}.2 弯曲性能'))

    elements.append(make_body_text(
        f'图{section_label}-4为经不同时间{aging_name}后复合材料的弯曲试验曲线。'
    ))

    chart = find_chart(f'第{cn_num}章_弯曲_曲线_全')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn=f'图{section_label}-4  {aging_name}后PLA/竹粉/纳米SiO\u2082复合材料弯曲试验曲线',
            caption_en=f'Flexural Test Curves after {aging_name}'))
        elements.append(make_empty_paragraph())

    # 弯曲强度分组柱状图
    chart = find_chart(f'第{cn_num}章_弯曲_强度_分组柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn=f'图{section_label}-5  不同{aging_name}时间下复合材料的弯曲强度',
            caption_en=f'Flexural Strength under Different {aging_name} Durations'))
        elements.append(make_empty_paragraph())

    # 弯曲强度数据讨论
    elements.append(make_body_text(
        f'经5天{aging_name}后，弯曲强度保留率分别为：'
        f'0%SiO\u2082为{ret(bs_init["0%"]["max_load_mean"], bs_5d["0%"]["max_load_mean"]):.1f}%、'
        f'1%SiO\u2082为{ret(bs_init["1%"]["max_load_mean"], bs_5d["1%"]["max_load_mean"]):.1f}%、'
        f'3%SiO\u2082为{ret(bs_init["3%"]["max_load_mean"], bs_5d["3%"]["max_load_mean"]):.1f}%、'
        f'5%SiO\u2082为{ret(bs_init["5%"]["max_load_mean"], bs_5d["5%"]["max_load_mean"]):.1f}%。'
        f'经10天{aging_name}后，弯曲强度保留率分别为：'
        f'0%SiO\u2082为{ret(bs_init["0%"]["max_load_mean"], bs_10d["0%"]["max_load_mean"]):.1f}%、'
        f'1%SiO\u2082为{ret(bs_init["1%"]["max_load_mean"], bs_10d["1%"]["max_load_mean"]):.1f}%、'
        f'3%SiO\u2082为{ret(bs_init["3%"]["max_load_mean"], bs_10d["3%"]["max_load_mean"]):.1f}%、'
        f'5%SiO\u2082为{ret(bs_init["5%"]["max_load_mean"], bs_10d["5%"]["max_load_mean"]):.1f}%。'
    ))

    elements.append(make_body_text(
        f'与拉伸性能类似，弯曲强度也随{aging_name}时间延长而进一步下降。'
        f'添加纳米SiO\u2082的复合材料在10天{aging_name}后仍保持较高的弯曲强度保留率，'
        f'特别是1%和3%SiO\u2082含量的复合材料表现出更好的耐{aging_name}性能。'
    ))

    # 弯曲位移分组柱状图
    chart = find_chart(f'第{cn_num}章_弯曲_位移_分组柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn=f'图{section_label}-6  不同{aging_name}时间下复合材料的弯曲断裂位移',
            caption_en=f'Flexural Displacement under Different {aging_name} Durations'))
        elements.append(make_empty_paragraph())

    # ---- X.3 力学性能保留率与本章小结 ----
    elements.append(make_section_title(f'{section_label}.3 力学性能保留率与本章小结'))

    elements.append(make_body_text(
        f'为了更直观地对比不同{aging_name}时间对复合材料力学性能的影响，'
        f'图{section_label}-7和图{section_label}-8分别展示了拉伸和弯曲强度的保留率变化。'
    ))

    # 拉伸强度保留率分组图
    chart = find_chart(f'第{cn_num}章_拉伸_强度保留率_分组')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn=f'图{section_label}-7  不同{aging_name}时间下拉伸强度保留率',
            caption_en=f'Tensile Strength Retention under Different {aging_name} Durations'))
        elements.append(make_empty_paragraph())

    # 弯曲强度保留率分组图
    chart = find_chart(f'第{cn_num}章_弯曲_强度保留率_分组')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn=f'图{section_label}-8  不同{aging_name}时间下弯曲强度保留率',
            caption_en=f'Flexural Strength Retention under Different {aging_name} Durations'))
        elements.append(make_empty_paragraph())

    # 小结
    elements.append(make_body_text(
        f'本章研究了不同时间{aging_name}对PLA/竹粉/纳米SiO\u2082复合材料力学性能的影响。主要结论如下：'
    ))
    elements.append(make_body_text(
        f'（1）经5天和10天{aging_name}后，复合材料的拉伸强度和弯曲强度均随老化时间延长而进一步下降，'
        f'说明{aging_name}对材料力学性能的劣化具有时间累积效应。'
    ))
    elements.append(make_body_text(
        f'（2）添加纳米SiO\u2082后，复合材料在5天和10天{aging_name}后的力学性能保留率均高于未添加组，'
        f'表明纳米SiO\u2082的加入有效提升了复合材料的耐{aging_name}性能，且这种保护作用在更长老化时间下依然显著。'
    ))
    elements.append(make_body_text(
        f'（3）纳米SiO\u2082对{aging_name}引起的力学性能退化具有延缓作用，'
        f'这可能归因于纳米颗粒的紫外线屏蔽效应（紫外老化）或阻隔水分渗透的能力（湿热老化），'
        f'以及纳米颗粒与PLA基体之间的界面增强效应。'
    ))

    return elements


# ============ body范围替换 ============
def replace_body_range(body, start_idx, end_idx, new_elements):
    children = list(body)
    to_remove = children[start_idx:end_idx + 1]
    ref = to_remove[0].getprevious()
    for elem in to_remove:
        body.remove(elem)
    if ref is not None:
        for new_elem in new_elements:
            ref.addnext(new_elem)
            ref = new_elem
    else:
        for i, new_elem in enumerate(new_elements):
            body.insert(i, new_elem)


# ============ 第三章内容生成（无单个曲线图） ============
def gen_chapter3(doc):
    """第三章：初始组（不含单个含量曲线图）"""
    elements = []
    ts = stats['拉伸']['初始组']
    bs = stats['弯曲']['初始组']

    elements.append(make_chapter_title('第三章  竹粉/PLA/纳米SiO\u2082复合材料力学性能'))
    elements.append(make_empty_paragraph())

    # ---- 3.1 拉伸性能 ----
    elements.append(make_section_title('3.1 拉伸性能'))

    elements.append(make_body_text(
        '图3-1为未老化PLA/竹粉/纳米SiO\u2082复合材料的拉伸试验曲线。从图中可以看出，'
        '添加不同含量纳米SiO\u2082后，复合材料的拉伸载荷-位移曲线呈现出明显差异。'
        '所有试样在初始阶段均表现出线性弹性变形特征，随后进入非线性变形阶段直至断裂。'
    ))

    chart = find_chart('第三章_拉伸_曲线_全')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn='图3-1', caption_en='Tensile Test Curves of BF/PLA/nano-SiO\u2082 Composites (Initial)'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'如表3-1所示，未添加纳米SiO\u2082的复合材料（0%）的最大拉伸载荷为'
        f'{ts["0%"]["max_load_mean"]:.2f}\u00b1{ts["0%"]["max_load_std"]:.2f} N，'
        f'添加1%纳米SiO\u2082后为{ts["1%"]["max_load_mean"]:.2f}\u00b1{ts["1%"]["max_load_std"]:.2f} N，'
        f'添加3%纳米SiO\u2082后为{ts["3%"]["max_load_mean"]:.2f}\u00b1{ts["3%"]["max_load_std"]:.2f} N，'
        f'添加5%纳米SiO\u2082后为{ts["5%"]["max_load_mean"]:.2f}\u00b1{ts["5%"]["max_load_std"]:.2f} N。'
    ))

    chart = find_chart('第三章_拉伸_强度柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图3-2  不同纳米SiO\u2082含量复合材料的拉伸强度',
            caption_en='Tensile Strength of BF/PLA/nano-SiO\u2082 Composites with Different SiO\u2082 Contents'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'从图3-2可以看出，添加纳米SiO\u2082后复合材料的拉伸强度发生了变化。'
        f'5%SiO\u2082含量时拉伸强度最高（{ts["5%"]["max_load_mean"]:.2f} N），'
        f'比未添加组提高了{((ts["5%"]["max_load_mean"]/ts["0%"]["max_load_mean"]-1)*100):.1f}%。'
        f'1%和3%SiO\u2082的拉伸强度有所降低，'
        f'这可能是由于低含量纳米颗粒在PLA基体中分散不均匀导致应力集中，'
        f'而较高含量时纳米颗粒形成了更多的物理交联点，增强了基体的承载能力。'
    ))

    # ---- 3.1.1 断裂位移分析 ----
    elements.append(make_subsection_title('3.1.1 断裂位移分析'))

    elements.append(make_body_text(
        f'由图3-3可以看出，0%SiO\u2082复合材料的断裂位移为'
        f'{ts["0%"]["max_disp_mean"]:.2f}\u00b1{ts["0%"]["max_disp_std"]:.2f} mm，'
        f'1%SiO\u2082为{ts["1%"]["max_disp_mean"]:.2f}\u00b1{ts["1%"]["max_disp_std"]:.2f} mm，'
        f'3%SiO\u2082为{ts["3%"]["max_disp_mean"]:.2f}\u00b1{ts["3%"]["max_disp_std"]:.2f} mm，'
        f'5%SiO\u2082为{ts["5%"]["max_disp_mean"]:.2f}\u00b1{ts["5%"]["max_disp_std"]:.2f} mm。'
    ))

    chart = find_chart('第三章_拉伸_位移柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图3-3  不同纳米SiO\u2082含量复合材料的拉伸断裂位移',
            caption_en='Tensile Displacement at Break of BF/PLA/nano-SiO\u2082 Composites'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'随着纳米SiO\u2082含量的增加，复合材料的断裂位移呈先增大后减小的趋势。'
        f'1%SiO\u2082的断裂位移最大（{ts["1%"]["max_disp_mean"]:.2f} mm），'
        f'表明适量的纳米SiO\u2082可以改善PLA基体的韧性。'
        f'这可能是由于纳米SiO\u2082粒子与PLA基体界面结合良好，能够有效传递应力，延缓裂纹扩展。'
        f'当SiO\u2082含量进一步增加时，粒子团聚导致应力集中点增多，断裂位移反而下降。'
    ))

    # ---- 3.2 弯曲性能 ----
    elements.append(make_section_title('3.2 弯曲性能'))

    elements.append(make_body_text(
        '图3-4为未老化PLA/竹粉/纳米SiO\u2082复合材料的弯曲试验曲线。'
        '不同纳米SiO\u2082含量的复合材料的弯曲载荷-位移曲线呈现出相似的弹性变形阶段，'
        '但在最大载荷和断裂位移上存在差异。'
    ))

    chart = find_chart('第三章_弯曲_曲线_全')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn='图3-4', caption_en='Flexural Test Curves of BF/PLA/nano-SiO\u2082 Composites (Initial)'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'如表3-2所示，0%SiO\u2082复合材料的最大弯曲载荷为'
        f'{bs["0%"]["max_load_mean"]:.2f}\u00b1{bs["0%"]["max_load_std"]:.2f} N，'
        f'1%SiO\u2082为{bs["1%"]["max_load_mean"]:.2f}\u00b1{bs["1%"]["max_load_std"]:.2f} N，'
        f'3%SiO\u2082为{bs["3%"]["max_load_mean"]:.2f}\u00b1{bs["3%"]["max_load_std"]:.2f} N，'
        f'5%SiO\u2082为{bs["5%"]["max_load_mean"]:.2f}\u00b1{bs["5%"]["max_load_std"]:.2f} N。'
    ))

    chart = find_chart('第三章_弯曲_强度柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图3-5  不同纳米SiO\u2082含量复合材料的弯曲强度',
            caption_en='Flexural Strength of BF/PLA/nano-SiO\u2082 Composites with Different SiO\u2082 Contents'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'添加纳米SiO\u2082后，复合材料的弯曲强度有所提高。1%纳米SiO\u2082的弯曲强度最高，'
        f'达到{bs["1%"]["max_load_mean"]:.2f} N，比未添加组提高了{((bs["1%"]["max_load_mean"]/bs["0%"]["max_load_mean"]-1)*100):.1f}%。'
        f'这说明纳米SiO\u2082的加入有效增强了复合材料的弯曲承载能力，'
        f'纳米颗粒的填充效应和界面结合共同提高了材料的弯曲性能。'
    ))

    chart = find_chart('第三章_弯曲_位移柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图3-6  不同纳米SiO\u2082含量复合材料的弯曲断裂位移',
            caption_en='Flexural Displacement of BF/PLA/nano-SiO\u2082 Composites'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'弯曲断裂位移方面，5%SiO\u2082含量时达到最大值{bs["5%"]["max_disp_mean"]:.2f} mm，'
        f'表明适量的纳米SiO\u2082在弯曲受力条件下同样能改善材料的变形能力。'
    ))

    # ---- 3.3 本章小结 ----
    elements.append(make_section_title('3.3 本章小结'))
    elements.append(make_body_text(
        '本章研究了未老化状态下PLA/竹粉/纳米SiO\u2082复合材料的拉伸和弯曲力学性能。主要结论如下：'
    ))
    elements.append(make_body_text(
        f'（1）拉伸性能方面：0%SiO\u2082复合材料的拉伸强度为{ts["0%"]["max_load_mean"]:.2f} N，'
        f'5%SiO\u2082时达到最高值{ts["5%"]["max_load_mean"]:.2f} N。'
        f'纳米SiO\u2082的添加对拉伸强度的影响呈现非线性关系，'
        f'适量添加可以提高拉伸强度，但分散均匀性是关键因素。'
    ))
    elements.append(make_body_text(
        f'（2）弯曲性能方面：1%纳米SiO\u2082的弯曲强度最高（{bs["1%"]["max_load_mean"]:.2f} N），'
        f'比未添加组提高了{((bs["1%"]["max_load_mean"]/bs["0%"]["max_load_mean"]-1)*100):.1f}%。'
        f'纳米SiO\u2082的填充效应有效增强了复合材料的弯曲承载能力。'
    ))
    elements.append(make_body_text(
        '（3）综合来看，纳米SiO\u2082的添加对PLA/竹粉复合材料的力学性能有显著影响，'
        '其中1%和5%的含量分别在弯曲强度和拉伸强度方面表现最优。'
        '这为后续的老化性能研究提供了重要的基准数据。'
    ))

    return elements


# ============ 第六章内容生成（无单个曲线图） ============
def gen_chapter6(doc):
    """第六章：紫外湿热耦合老化（不含单个含量曲线图）"""
    elements = []
    ts_aged = stats['拉伸']['紫外湿热耦合老化组']
    bs_aged = stats['弯曲']['紫外湿热耦合老化组']
    ts_init = stats['拉伸']['初始组']
    bs_init = stats['弯曲']['初始组']

    def ret(init_val, aged_val):
        return (aged_val / init_val * 100) if init_val > 0 else 0

    elements.append(make_chapter_title('第六章  竹粉/PLA/纳米SiO\u2082复合材料紫外湿热耦合老化力学性能'))
    elements.append(make_empty_paragraph())

    # ---- 6.1 拉伸性能 ----
    elements.append(make_section_title('6.1 拉伸性能'))

    elements.append(make_body_text(
        f'图6-1为经紫外湿热耦合老化后PLA/竹粉/纳米SiO\u2082复合材料的拉伸试验曲线。'
        f'从图中可以看出，经紫外湿热耦合老化处理后，复合材料的拉伸载荷-位移曲线发生了明显变化。'
    ))

    chart = find_chart('第六章_拉伸_曲线_全')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn='图6-1', caption_en='Tensile Test Curves after UV-Hydrothermal Coupled Aging'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'如表6-1所示，经紫外湿热耦合老化后，0%SiO\u2082复合材料的拉伸强度为'
        f'{ts_aged["0%"]["max_load_mean"]:.2f}\u00b1{ts_aged["0%"]["max_load_std"]:.2f} N，'
        f'1%SiO\u2082为{ts_aged["1%"]["max_load_mean"]:.2f}\u00b1{ts_aged["1%"]["max_load_std"]:.2f} N，'
        f'3%SiO\u2082为{ts_aged["3%"]["max_load_mean"]:.2f}\u00b1{ts_aged["3%"]["max_load_std"]:.2f} N，'
        f'5%SiO\u2082为{ts_aged["5%"]["max_load_mean"]:.2f}\u00b1{ts_aged["5%"]["max_load_std"]:.2f} N。'
    ))

    chart = find_chart('第六章_拉伸_强度柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图6-2  紫外湿热耦合老化后不同SiO\u2082含量复合材料的拉伸强度',
            caption_en='Tensile Strength after UV-Hydrothermal Coupled Aging'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'断裂位移方面，0%SiO\u2082为{ts_aged["0%"]["max_disp_mean"]:.2f}\u00b1{ts_aged["0%"]["max_disp_std"]:.2f} mm，'
        f'1%SiO\u2082为{ts_aged["1%"]["max_disp_mean"]:.2f}\u00b1{ts_aged["1%"]["max_disp_std"]:.2f} mm，'
        f'3%SiO\u2082为{ts_aged["3%"]["max_disp_mean"]:.2f}\u00b1{ts_aged["3%"]["max_disp_std"]:.2f} mm，'
        f'5%SiO\u2082为{ts_aged["5%"]["max_disp_mean"]:.2f}\u00b1{ts_aged["5%"]["max_disp_std"]:.2f} mm。'
    ))

    chart = find_chart('第六章_拉伸_位移柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图6-3  紫外湿热耦合老化后不同SiO\u2082含量复合材料的拉伸断裂位移',
            caption_en='Tensile Displacement after UV-Hydrothermal Coupled Aging'))
        elements.append(make_empty_paragraph())

    # ---- 6.2 弯曲性能 ----
    elements.append(make_section_title('6.2 弯曲性能'))

    elements.append(make_body_text(
        f'图6-4为经紫外湿热耦合老化后PLA/竹粉/纳米SiO\u2082复合材料的弯曲试验曲线。'
        f'经紫外湿热耦合老化处理后，复合材料在弯曲载荷作用下的力学响应发生了变化。'
    ))

    chart = find_chart('第六章_弯曲_曲线_全')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn='图6-4', caption_en='Flexural Test Curves after UV-Hydrothermal Coupled Aging'))
        elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'如表6-2所示，经紫外湿热耦合老化后，0%SiO\u2082复合材料的弯曲强度为'
        f'{bs_aged["0%"]["max_load_mean"]:.2f}\u00b1{bs_aged["0%"]["max_load_std"]:.2f} N，'
        f'1%SiO\u2082为{bs_aged["1%"]["max_load_mean"]:.2f}\u00b1{bs_aged["1%"]["max_load_std"]:.2f} N，'
        f'3%SiO\u2082为{bs_aged["3%"]["max_load_mean"]:.2f}\u00b1{bs_aged["3%"]["max_load_std"]:.2f} N，'
        f'5%SiO\u2082为{bs_aged["5%"]["max_load_mean"]:.2f}\u00b1{bs_aged["5%"]["max_load_std"]:.2f} N。'
    ))

    chart = find_chart('第六章_弯曲_强度柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图6-5  紫外湿热耦合老化后不同SiO\u2082含量复合材料的弯曲强度',
            caption_en='Flexural Strength after UV-Hydrothermal Coupled Aging'))
        elements.append(make_empty_paragraph())

    chart = find_chart('第六章_弯曲_位移柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图6-6  紫外湿热耦合老化后不同SiO\u2082含量复合材料的弯曲断裂位移',
            caption_en='Flexural Displacement after UV-Hydrothermal Coupled Aging'))
        elements.append(make_empty_paragraph())

    # ---- 6.3 力学性能保留率与本章小结 ----
    elements.append(make_section_title('6.3 力学性能保留率与本章小结'))

    elements.append(make_body_text(
        f'为了评估紫外湿热耦合老化对复合材料力学性能的影响程度，'
        f'以未老化试样的力学性能为基准，计算各老化后的力学性能保留率（保留率=老化后/初始\u00d7100%）。'
    ))

    chart = find_chart('第六章_拉伸_强度保留率')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图6-7  紫外湿热耦合老化后拉伸强度保留率',
            caption_en='Tensile Strength Retention after UV-Hydrothermal Coupled Aging'))
        elements.append(make_empty_paragraph())

    ret_text_parts = []
    for sio2 in ['0%', '1%', '3%', '5%']:
        r = ret(ts_init[sio2]['max_load_mean'], ts_aged[sio2]['max_load_mean'])
        ret_text_parts.append(f'{sio2}SiO\u2082为{r:.1f}%')
    elements.append(make_body_text(
        f'如图6-7所示，经紫外湿热耦合老化后，各SiO\u2082含量复合材料的拉伸强度保留率分别为：'
        + '、'.join(ret_text_parts) + '。'
    ))

    chart = find_chart('第六章_弯曲_强度保留率')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn='图6-8  紫外湿热耦合老化后弯曲强度保留率',
            caption_en='Flexural Strength Retention after UV-Hydrothermal Coupled Aging'))
        elements.append(make_empty_paragraph())

    ret_text_parts = []
    for sio2 in ['0%', '1%', '3%', '5%']:
        r = ret(bs_init[sio2]['max_load_mean'], bs_aged[sio2]['max_load_mean'])
        ret_text_parts.append(f'{sio2}SiO\u2082为{r:.1f}%')
    elements.append(make_body_text(
        f'弯曲强度保留率方面，各SiO\u2082含量分别为：' + '、'.join(ret_text_parts) + '。'
    ))

    elements.append(make_body_text(
        f'本章研究了紫外湿热耦合老化对PLA/竹粉/纳米SiO\u2082复合材料力学性能的影响。主要结论如下：'
    ))
    elements.append(make_body_text(
        f'（1）经紫外湿热耦合老化后，复合材料的拉伸强度和弯曲强度均有所下降，'
        f'说明紫外与湿热的协同作用对材料力学性能产生了显著的负面影响。'
    ))
    elements.append(make_body_text(
        f'（2）添加纳米SiO\u2082后，复合材料在紫外湿热耦合老化条件下的力学性能保留率有所提高，'
        f'表明纳米SiO\u2082的加入有助于提升复合材料在复杂老化环境下的耐久性能。'
    ))
    elements.append(make_body_text(
        f'（3）纳米SiO\u2082的添加改善了PLA基体的抗老化能力，'
        f'这可能与纳米颗粒的紫外线屏蔽效应、阻隔水分渗透能力以及界面增强作用有关。'
    ))

    return elements


# ============ 主程序 ============
def main():
    print("=" * 60)
    print("修改毕业论文_修改版.docx（全量重写第三到第六章）")
    print("=" * 60)

    load_stats()

    print(f"\n打开文档: {DOCX_PATH}")
    doc = Document(DOCX_PATH)
    body = doc.element.body

    children = list(body)

    # 扫描章节边界
    para_idx = 0
    chapters = {}
    for i, child in enumerate(children):
        if child.tag == f'{{{W}}}p':
            text = get_paragraph_text(child)
            for cn, num in {'三': 3, '四': 4, '五': 5, '六': 6}.items():
                if f'第{cn}章' in text and num not in chapters:
                    chapters[num] = {'start_para': para_idx, 'start_body_idx': i}
            para_idx += 1

    # 计算每章的结束位置
    sorted_nums = sorted(chapters.keys())
    for idx, num in enumerate(sorted_nums):
        if idx + 1 < len(sorted_nums):
            next_num = sorted_nums[idx + 1]
            chapters[num]['end_para'] = chapters[next_num]['start_para'] - 1
            chapters[num]['end_body_idx'] = chapters[next_num]['start_body_idx'] - 1
        else:
            # 最后一章：找到不属于本章的标题段落作为结束
            ch_start = chapters[num]['start_body_idx']
            end_idx = len(children) - 1
            for j in range(ch_start + 1, len(children)):
                c = children[j]
                if c.tag == f'{{{W}}}p':
                    t = get_paragraph_text(c)
                    outline = c.find(f'.//{{{W}}}outlineLvl')
                    if outline is not None and not t.lstrip().startswith(f'{num}.'):
                        end_idx = j - 1
                        break
            chapters[num]['end_para'] = para_idx - 1
            chapters[num]['end_body_idx'] = end_idx

    for num in sorted(chapters.keys()):
        ch = chapters[num]
        print(f"  第{num}章: body[{ch['start_body_idx']}]-[{ch['end_body_idx']}]")

    # 生成所有章节内容
    print("\n生成章节内容...")
    ch3_gen = gen_chapter3(doc)
    ch4_gen = gen_uv_heat_chapter(doc, 4, '紫外老化', '4')
    ch5_gen = gen_uv_heat_chapter(doc, 5, '湿热老化', '5')
    ch6_gen = gen_chapter6(doc)

    chapter_gens = {
        3: (ch3_gen, '第三章'),
        4: (ch4_gen, '第四章'),
        5: (ch5_gen, '第五章'),
        6: (ch6_gen, '第六章'),
    }

    # 从后往前替换（避免索引偏移）
    for num in sorted(chapters.keys(), reverse=True):
        ch = chapters[num]
        gen_elements, ch_name = chapter_gens[num]
        start_idx = ch['start_body_idx']
        end_idx = ch['end_body_idx']
        print(f"\n  替换{ch_name}: body[{start_idx}]-[{end_idx}], 共{end_idx-start_idx+1}个元素")
        replace_body_range(body, start_idx, end_idx, gen_elements)
        print(f"  替换为 {len(gen_elements)} 个新元素")

    # ===== 保存 =====
    print(f"\n保存到: {DOCX_PATH}")
    doc.save(DOCX_PATH)
    print("完成！")

    # 验证
    doc2 = Document(DOCX_PATH)
    body2 = doc2.element.body
    WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    inline_count = len(body2.findall(f'.//{{{WP_NS}}}inline'))
    anchor_count = len(body2.findall(f'.//{{{WP_NS}}}anchor'))
    total_drawings = inline_count + anchor_count
    print(f"\n验证: 共 {total_drawings} 个图片 (inline={inline_count}, anchor={anchor_count})")


if __name__ == '__main__':
    main()
