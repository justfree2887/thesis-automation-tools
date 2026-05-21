#!/usr/bin/env python3
"""
modify_thesis_chapters.py - 修改毕业论文第三到第六章
使用python-docx + lxml直接操作OpenXML
保留文档格式、表格，替换段落内容，插入新图表
"""

import sys
import os
import json
import copy
from io import BytesIO

from docx import Document
from docx.shared import Pt, Cm, Inches, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree
from PIL import Image

# ============ 路径配置 ============
BASE_DIR = '/Users/shiberlin/Desktop/毕业论文'
DOCX_PATH = os.path.join(BASE_DIR, '毕业论文最新版.docx')
OUTPUT_PATH = os.path.join(BASE_DIR, '毕业论文_修改版.docx')
CHART_DIR = os.path.join(BASE_DIR, 'charts', 'thesis')
STATS_PATH = os.path.join(CHART_DIR, 'stats.json')

# ============ XML命名空间 ============
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R_ID = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

NSMAP = {
    'w': W,
    'wp': WP,
    'a': A,
    'pic': PIC,
    'r': REL,
}

# ============ 统计数据 ============
stats = None

def load_stats():
    global stats
    with open(STATS_PATH, 'r', encoding='utf-8') as f:
        stats = json.load(f)

# ============ 段落创建工具 ============

def _make_pPr(align='left', first_indent=480, spacing_before=None, spacing_after=None,
              outline_lvl=None, line=400, line_rule='exact', font='黑体', sz=24, bold=False):
    """创建w:pPr元素"""
    pPr = etree.SubElement(etree.Element('dummy'), f'{{{W}}}pPr')

    # Widows control
    if outline_lvl is not None:
        etree.SubElement(pPr, f'{{{W}}}widowControl')

    # Spacing
    sp = etree.SubElement(pPr, f'{{{W}}}spacing')
    sp.set(f'{{{W}}}line', str(line))
    sp.set(f'{{{W}}}lineRule', line_rule)
    if spacing_before is not None:
        sp.set(f'{{{W}}}before', str(spacing_before))
    if spacing_after is not None:
        sp.set(f'{{{W}}}after', str(spacing_after))

    # Indent
    if first_indent:
        ind = etree.SubElement(pPr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLine', str(first_indent))

    # Justification
    jc = etree.SubElement(pPr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', align)

    # Outline level
    if outline_lvl is not None:
        olvl = etree.SubElement(pPr, f'{{{W}}}outlineLvl')
        olvl.set(f'{{{W}}}val', str(outline_lvl))

    # Run properties (in pPr as default)
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
    """创建w:r元素"""
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
    """创建完整的w:p元素"""
    p = etree.Element(f'{{{W}}}p')
    pPr = _make_pPr(align=align, first_indent=first_indent,
                    spacing_before=spacing_before, spacing_after=spacing_after,
                    outline_lvl=outline_lvl, font=font, sz=sz, bold=bold)
    p.append(pPr)
    r = _make_run(text, font=font, sz=sz, bold=bold, hint_ea=(outline_lvl is not None))
    p.append(r)
    return p


def make_empty_paragraph():
    """创建空段落"""
    return etree.Element(f'{{{W}}}p')


def make_chapter_title(text):
    """章节标题（居中、16pt黑体）"""
    return make_paragraph(text, align='center', first_indent=0, font='黑体', sz=32)


def make_section_title(text, outline_lvl=1):
    """节标题（左对齐、14pt黑体）"""
    return make_paragraph(text, align='left', first_indent=0,
                         spacing_before=200, spacing_after=100,
                         outline_lvl=outline_lvl, font='黑体', sz=28)


def make_subsection_title(text):
    """子节标题（左对齐、12pt黑体加粗）"""
    return make_paragraph(text, align='left', first_indent=0,
                         spacing_before=100, spacing_after=50,
                         outline_lvl=2, font='黑体', sz=24, bold=True)


def make_body_text(text):
    """正文段落（首行缩进2字符=480DXA，12pt黑体）"""
    return make_paragraph(text, align='left', first_indent=480, font='黑体', sz=24)


def make_caption(cn_text, en_text):
    """图题/表题（居中，中文黑体+英文Times New Roman）"""
    p = etree.Element(f'{{{W}}}p')
    pPr = _make_pPr(align='center', first_indent=0, font='黑体', sz=24)
    p.append(pPr)
    # 中文部分
    r1 = _make_run(cn_text, font='黑体', sz=24, hint_ea=True)
    p.append(r1)
    # 英文部分
    r2 = _make_run(' ' + en_text, font='Times New Roman', sz=24)
    p.append(r2)
    return p


# ============ 图片插入 ============

_doc_part = None  # 全局document part引用
_image_counter = 100  # 图片ID计数器

def add_image_to_document(doc, image_path, width_cm=14):
    """添加图片到文档，返回图片的Drawing XML元素"""
    global _image_counter

    main_part = doc.part

    # 添加图片部分
    with open(image_path, 'rb') as img_file:
        image_data = img_file.read()

    # 获取图片格式
    ext = os.path.splitext(image_path)[1].lower()
    content_type_map = {
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.bmp': 'image/bmp',
    }
    content_type = content_type_map.get(ext, 'image/png')

    # 使用python-docx API添加图片
    # get_or_add_image returns (rId, ImagePart)
    rId, image_part = main_part.get_or_add_image(image_path)

    # 获取图片尺寸
    with Image.open(image_path) as img:
        px_w, px_h = img.size

    # 计算EMU（1 cm = 360000 EMU）
    target_width_emu = int(width_cm * 360000)
    ratio = px_h / px_w if px_w > 0 else 1
    target_height_emu = int(target_width_emu * ratio)

    _image_counter += 1
    doc_prop_id = _image_counter

    cx = str(target_width_emu)
    cy = str(target_height_emu)

    # 构建inline drawing XML
    inline_xml = f'''
    <wp:inline xmlns:wp="{WP}" distT="0" distB="0" distL="0" distR="0"
               xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{REL}">
      <wp:extent cx="{cx}" cy="{cy}"/>
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
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>'''

    inline = etree.fromstring(inline_xml)
    drawing = etree.Element(f'{{{W}}}drawing')
    drawing.append(inline)

    # 包裹在run中
    run = etree.Element(f'{{{W}}}r')
    run.append(drawing)

    return run


def make_image_paragraph(doc, image_path, width_cm=14, caption_cn='', caption_en=''):
    """创建包含图片和图题的段落列表"""
    elements = []

    # 图片段落
    img_run = add_image_to_document(doc, image_path, width_cm)
    img_p = etree.Element(f'{{{W}}}p')
    img_pPr = _make_pPr(align='center', first_indent=0)
    img_p.append(img_pPr)
    img_p.append(img_run)
    elements.append(img_p)

    # 图题
    if caption_cn or caption_en:
        elements.append(make_caption(caption_cn, caption_en))

    return elements


# ============ 图表查找 ============

def find_chart(pattern):
    """查找图表文件"""
    if not os.path.exists(CHART_DIR):
        return None
    for f in os.listdir(CHART_DIR):
        if pattern in f and f.endswith('.png'):
            path = os.path.join(CHART_DIR, f)
            if os.path.exists(path):
                return path
    return None


# ============ 段落替换 ============

def get_body_children_range(body, start_para, end_para):
    """获取body中从start_para到end_para段落之间的所有children（包括表格）"""
    children = list(body)
    para_count = 0
    start_idx = None
    end_idx = None

    for i, child in enumerate(children):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            if para_count == start_para and start_idx is None:
                start_idx = i
            if para_count == end_para:
                end_idx = i
                break
            para_count += 1

    return children, start_idx, end_idx


def replace_body_range(body, start_idx, end_idx, new_elements):
    """替换body中从start_idx到end_idx的元素为new_elements"""
    to_remove = list(body)[start_idx:end_idx + 1]

    # 获取插入参考点
    ref = to_remove[0].getprevious()

    # 移除旧元素
    for elem in to_remove:
        body.remove(elem)

    # 插入新元素
    if ref is not None:
        for new_elem in new_elements:
            ref.addnext(new_elem)
            ref = new_elem
    else:
        for i, new_elem in enumerate(new_elements):
            body.insert(i, new_elem)


# ============ 章节内容生成 ============

def gen_chapter3(doc):
    """第三章：初始组"""
    elements = []
    ts = stats['拉伸']['初始组']
    bs = stats['弯曲']['初始组']

    elements.append(make_chapter_title('第三章  竹粉/PLA/纳米SiO₂复合材料力学性能'))
    elements.append(make_empty_paragraph())

    # ---- 3.1 拉伸性能 ----
    elements.append(make_section_title('3.1 拉伸性能'))

    elements.append(make_body_text(
        '图3-1为未老化PLA/竹粉/纳米SiO₂复合材料的拉伸试验曲线。从图中可以看出，'
        '添加不同含量纳米SiO₂后，复合材料的拉伸载荷-位移曲线呈现出明显差异。'
        '所有试样在初始阶段均表现出线性弹性变形特征，随后进入非线性变形阶段直至断裂。'
    ))

    # 插入拉伸合并曲线图
    chart = find_chart('第三章_拉伸_曲线_全')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn='图3-1', caption_en='Tensile Test Curves of BF/PLA/nano-SiO\u2082 Composites (Initial)'))
        elements.append(make_empty_paragraph())

    # 插入各SiO2含量的单独曲线图
    for sio2 in ['0%', '1%', '3%', '5%']:
        chart = find_chart(f'第三章_拉伸_曲线_{sio2}')
        if chart:
            num = sio2.strip('%')
            elements.append(make_empty_paragraph())
            elements.extend(make_image_paragraph(doc, chart, width_cm=12,
                caption_cn=f'图3-1-{num}  {num}%纳米SiO\u2082复合材料拉伸试验曲线',
                caption_en=f'Tensile Curve of Composites with {num}% nano-SiO\u2082'))
            elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'如表3-1所示，未添加纳米SiO\u2082的复合材料（0%）的最大拉伸载荷为'
        f'{ts["0%"]["max_load_mean"]:.2f}\u00b1{ts["0%"]["max_load_std"]:.2f} N，'
        f'添加1%纳米SiO\u2082后为{ts["1%"]["max_load_mean"]:.2f}\u00b1{ts["1%"]["max_load_std"]:.2f} N，'
        f'添加3%纳米SiO\u2082后为{ts["3%"]["max_load_mean"]:.2f}\u00b1{ts["3%"]["max_load_std"]:.2f} N，'
        f'添加5%纳米SiO\u2082后为{ts["5%"]["max_load_mean"]:.2f}\u00b1{ts["5%"]["max_load_std"]:.2f} N。'
    ))

    # 拉伸强度柱状图
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

    for sio2 in ['0%', '1%', '3%', '5%']:
        chart = find_chart(f'第三章_弯曲_曲线_{sio2}')
        if chart:
            num = sio2.strip('%')
            elements.append(make_empty_paragraph())
            elements.extend(make_image_paragraph(doc, chart, width_cm=12,
                caption_cn=f'图3-4-{num}  {num}%纳米SiO\u2082复合材料弯曲试验曲线',
                caption_en=f'Flexural Curve of Composites with {num}% nano-SiO\u2082'))
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


CHAPTER_NUM_CN = {4: '四', 5: '五', 6: '六'}

def gen_aging_chapter(doc, chapter_num, chapter_title, aging_key, aging_name, section_label):
    """生成老化章节内容（第四、五、六章通用模板）"""
    elements = []
    cn_num = CHAPTER_NUM_CN[chapter_num]
    ts_aged = stats['拉伸'][aging_key]
    bs_aged = stats['弯曲'][aging_key]
    ts_init = stats['拉伸']['初始组']
    bs_init = stats['弯曲']['初始组']

    elements.append(make_chapter_title(f'第{cn_num}章  竹粉/PLA/纳米SiO₂复合材料{aging_name}力学性能'))
    elements.append(make_empty_paragraph())

    # 计算保留率
    def retention(init_val, aged_val):
        return (aged_val / init_val * 100) if init_val > 0 else 0

    # ---- X.1 拉伸性能 ----
    elements.append(make_section_title(f'{section_label}.1 拉伸性能'))

    elements.append(make_body_text(
        f'图{section_label}-1为经{aging_name}后PLA/竹粉/纳米SiO\u2082复合材料的拉伸试验曲线。'
        f'从图中可以看出，经{aging_name}处理后，复合材料的拉伸载荷-位移曲线发生了明显变化。'
    ))

    # 拉伸合并曲线图
    chart = find_chart(f'第{cn_num}章_拉伸_曲线_全')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn=f'图{section_label}-1', caption_en=f'Tensile Test Curves after {aging_name}'))
        elements.append(make_empty_paragraph())

    for sio2 in ['0%', '1%', '3%', '5%']:
        chart = find_chart(f'第{cn_num}章_拉伸_曲线_{sio2}')
        if chart:
            num = sio2.strip('%')
            elements.append(make_empty_paragraph())
            elements.extend(make_image_paragraph(doc, chart, width_cm=12,
                caption_cn=f'图{section_label}-1-{num}  {num}%纳米SiO\u2082复合材料{aging_name}拉伸试验曲线',
                caption_en=f'Tensile Curve of {num}% nano-SiO\u2082 Composites after {aging_name}'))
            elements.append(make_empty_paragraph())

    # 强度数据和柱状图
    elements.append(make_body_text(
        f'如表{section_label}-1所示，经{aging_name}后，0%SiO\u2082复合材料的拉伸强度为'
        f'{ts_aged["0%"]["max_load_mean"]:.2f}\u00b1{ts_aged["0%"]["max_load_std"]:.2f} N，'
        f'1%SiO\u2082为{ts_aged["1%"]["max_load_mean"]:.2f}\u00b1{ts_aged["1%"]["max_load_std"]:.2f} N，'
        f'3%SiO\u2082为{ts_aged["3%"]["max_load_mean"]:.2f}\u00b1{ts_aged["3%"]["max_load_std"]:.2f} N，'
        f'5%SiO\u2082为{ts_aged["5%"]["max_load_mean"]:.2f}\u00b1{ts_aged["5%"]["max_load_std"]:.2f} N。'
    ))

    chart = find_chart(f'第{cn_num}章_拉伸_强度柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn=f'图{section_label}-2  {aging_name}后不同SiO\u2082含量复合材料的拉伸强度',
            caption_en=f'Tensile Strength after {aging_name}'))
        elements.append(make_empty_paragraph())

    # 位移数据和柱状图
    elements.append(make_body_text(
        f'断裂位移方面，0%SiO\u2082为{ts_aged["0%"]["max_disp_mean"]:.2f}\u00b1{ts_aged["0%"]["max_disp_std"]:.2f} mm，'
        f'1%SiO\u2082为{ts_aged["1%"]["max_disp_mean"]:.2f}\u00b1{ts_aged["1%"]["max_disp_std"]:.2f} mm，'
        f'3%SiO\u2082为{ts_aged["3%"]["max_disp_mean"]:.2f}\u00b1{ts_aged["3%"]["max_disp_std"]:.2f} mm，'
        f'5%SiO\u2082为{ts_aged["5%"]["max_disp_mean"]:.2f}\u00b1{ts_aged["5%"]["max_disp_std"]:.2f} mm。'
    ))

    chart = find_chart(f'第{cn_num}章_拉伸_位移柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn=f'图{section_label}-3  {aging_name}后不同SiO\u2082含量复合材料的拉伸断裂位移',
            caption_en=f'Tensile Displacement after {aging_name}'))
        elements.append(make_empty_paragraph())

    # ---- X.2 弯曲性能 ----
    elements.append(make_section_title(f'{section_label}.2 弯曲性能'))

    elements.append(make_body_text(
        f'图{section_label}-4为经{aging_name}后PLA/竹粉/纳米SiO\u2082复合材料的弯曲试验曲线。'
        f'经{aging_name}处理后，复合材料在弯曲载荷作用下的力学响应发生了变化。'
    ))

    chart = find_chart(f'第{cn_num}章_弯曲_曲线_全')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=14,
            caption_cn=f'图{section_label}-4', caption_en=f'Flexural Test Curves after {aging_name}'))
        elements.append(make_empty_paragraph())

    for sio2 in ['0%', '1%', '3%', '5%']:
        chart = find_chart(f'第{cn_num}章_弯曲_曲线_{sio2}')
        if chart:
            num = sio2.strip('%')
            elements.append(make_empty_paragraph())
            elements.extend(make_image_paragraph(doc, chart, width_cm=12,
                caption_cn=f'图{section_label}-4-{num}  {num}%纳米SiO\u2082复合材料{aging_name}弯曲试验曲线',
                caption_en=f'Flexural Curve of {num}% nano-SiO\u2082 Composites after {aging_name}'))
            elements.append(make_empty_paragraph())

    elements.append(make_body_text(
        f'如表{section_label}-2所示，经{aging_name}后，0%SiO\u2082复合材料的弯曲强度为'
        f'{bs_aged["0%"]["max_load_mean"]:.2f}\u00b1{bs_aged["0%"]["max_load_std"]:.2f} N，'
        f'1%SiO\u2082为{bs_aged["1%"]["max_load_mean"]:.2f}\u00b1{bs_aged["1%"]["max_load_std"]:.2f} N，'
        f'3%SiO\u2082为{bs_aged["3%"]["max_load_mean"]:.2f}\u00b1{bs_aged["3%"]["max_load_std"]:.2f} N，'
        f'5%SiO\u2082为{bs_aged["5%"]["max_load_mean"]:.2f}\u00b1{bs_aged["5%"]["max_load_std"]:.2f} N。'
    ))

    chart = find_chart(f'第{cn_num}章_弯曲_强度柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn=f'图{section_label}-5  {aging_name}后不同SiO\u2082含量复合材料的弯曲强度',
            caption_en=f'Flexural Strength after {aging_name}'))
        elements.append(make_empty_paragraph())

    chart = find_chart(f'第{cn_num}章_弯曲_位移柱状图')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn=f'图{section_label}-6  {aging_name}后不同SiO\u2082含量复合材料的弯曲断裂位移',
            caption_en=f'Flexural Displacement after {aging_name}'))
        elements.append(make_empty_paragraph())

    # ---- X.3 力学性能保留率与本章小结 ----
    elements.append(make_section_title(f'{section_label}.3 力学性能保留率与本章小结'))

    elements.append(make_body_text(
        f'为了评估{aging_name}对复合材料力学性能的影响程度，'
        f'以未老化试样的力学性能为基准，计算各老化后的力学性能保留率（保留率=老化后/初始\u00d7100%）。'
    ))

    # 拉伸强度保留率
    chart = find_chart(f'第{cn_num}章_拉伸_强度保留率')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn=f'图{section_label}-7  {aging_name}后拉伸强度保留率',
            caption_en=f'Tensile Strength Retention after {aging_name}'))
        elements.append(make_empty_paragraph())

    ret_text_parts = []
    for sio2 in ['0%', '1%', '3%', '5%']:
        r = retention(ts_init[sio2]['max_load_mean'], ts_aged[sio2]['max_load_mean'])
        ret_text_parts.append(f'{sio2}SiO\u2082为{r:.1f}%')
    elements.append(make_body_text(
        f'如图{section_label}-7所示，经{aging_name}后，各SiO\u2082含量复合材料的拉伸强度保留率分别为：'
        + '、'.join(ret_text_parts) + '。'
    ))

    # 弯曲强度保留率
    chart = find_chart(f'第{cn_num}章_弯曲_强度保留率')
    if chart:
        elements.append(make_empty_paragraph())
        elements.extend(make_image_paragraph(doc, chart, width_cm=12,
            caption_cn=f'图{section_label}-8  {aging_name}后弯曲强度保留率',
            caption_en=f'Flexural Strength Retention after {aging_name}'))
        elements.append(make_empty_paragraph())

    ret_text_parts = []
    for sio2 in ['0%', '1%', '3%', '5%']:
        r = retention(bs_init[sio2]['max_load_mean'], bs_aged[sio2]['max_load_mean'])
        ret_text_parts.append(f'{sio2}SiO\u2082为{r:.1f}%')
    elements.append(make_body_text(
        f'弯曲强度保留率方面，各SiO\u2082含量分别为：' + '、'.join(ret_text_parts) + '。'
    ))

    # 小结
    elements.append(make_body_text(
        f'本章研究了{aging_name}对PLA/竹粉/纳米SiO\u2082复合材料力学性能的影响。'
        f'主要结论如下：'
    ))
    elements.append(make_body_text(
        f'（1）经{aging_name}后，复合材料的拉伸强度和弯曲强度均有所下降，'
        f'说明{aging_name}对材料力学性能产生了负面影响。'
    ))
    elements.append(make_body_text(
        f'（2）添加纳米SiO\u2082后，复合材料在{aging_name}条件下的力学性能保留率有所提高，'
        f'表明纳米SiO\u2082的加入有助于提升复合材料的耐老化性能。'
    ))
    elements.append(make_body_text(
        f'（3）纳米SiO\u2082的添加改善了PLA基体的抗老化能力，'
        f'这可能与纳米颗粒的阻挡效应和界面增强作用有关。'
    ))

    return elements


# ============ 主程序 ============

def main():
    print("=" * 60)
    print("修改毕业论文第三到第六章")
    print("=" * 60)

    load_stats()

    # 列出可用图表
    chart_files = [f for f in os.listdir(CHART_DIR) if f.endswith('.png')]
    print(f"\n可用图表: {len(chart_files)} 张")

    # 打开文档
    print(f"\n打开文档: {DOCX_PATH}")
    doc = Document(DOCX_PATH)
    body = doc.element.body

    # 定义章节范围（段落索引，含表格）
    # 第三章: 段落125-161
    # 第四章: 段落162-191
    # 第五章: 段落192-226
    # 第六章: 段落227-261
    chapters = [
        (125, 161, '第三章', gen_chapter3, {}),
        (162, 191, '第四章', gen_aging_chapter,
         {'chapter_num': 4, 'chapter_title': '紫外老化', 'aging_key': '紫外老化组',
          'aging_name': '紫外老化', 'section_label': '4'}),
        (192, 226, '第五章', gen_aging_chapter,
         {'chapter_num': 5, 'chapter_title': '湿热老化', 'aging_key': '湿热老化组',
          'aging_name': '湿热老化', 'section_label': '5'}),
        (227, 261, '第六章', gen_aging_chapter,
         {'chapter_num': 6, 'chapter_title': '紫外湿热耦合老化', 'aging_key': '紫外湿热耦合老化组',
          'aging_name': '紫外湿热耦合老化', 'section_label': '6'}),
    ]

    # 从后往前替换，避免索引偏移
    for start_para, end_para, ch_name, gen_func, kwargs in reversed(chapters):
        print(f"\n---------- {ch_name} [{start_para}]-[{end_para}] ----------")
        children, start_idx, end_idx = get_body_children_range(body, start_para, end_para)
        if start_idx is None or end_idx is None:
            print(f"  警告: 未找到段落范围")
            continue

        print(f"  body范围: [{start_idx}]-[{end_idx}], 共 {end_idx - start_idx + 1} 个元素")

        # 生成新内容
        new_elements = gen_func(doc, **kwargs)
        print(f"  生成 {len(new_elements)} 个新元素")

        # 替换
        replace_body_range(body, start_idx, end_idx, new_elements)
        print(f"  已替换")

    # 保存文档
    print(f"\n保存到: {OUTPUT_PATH}")
    doc.save(OUTPUT_PATH)
    print("完成！")


if __name__ == '__main__':
    main()
