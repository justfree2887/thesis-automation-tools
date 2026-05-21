#!/usr/bin/env python3
"""
modify_thesis.py - 修改毕业论文第三到第六章
- 替换段落内容
- 插入图表
- 保留原有格式
"""

from docx import Document
from docx.shared import Pt, Cm, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from lxml import etree
import json
import os
import copy

BASE_DIR = '/Users/shiberlin/Desktop/毕业论文'
DOCX_PATH = os.path.join(BASE_DIR, '毕业论文最新版.docx')
OUTPUT_PATH = os.path.join(BASE_DIR, '毕业论文最新版_已修改.docx')
CHART_DIR = os.path.join(BASE_DIR, 'charts', 'thesis')
STATS_PATH = os.path.join(CHART_DIR, 'stats.json')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def load_stats():
    with open(STATS_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def make_run(text, sz=24, bold=False, font='黑体'):
    """创建一个格式化的run"""
    r = parse_xml(f'<w:r xmlns:w="{W}"><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:cs="{font}"/><w:kern w:val="0"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>{"<w:b/><w:bCs/>" if bold else ""}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>')
    return r

def make_para_xml(text, sz=24, bold=False, font='黑体', align='left', first_indent=480, spacing_before=None, spacing_after=None, outline_lvl=None):
    """创建一个完整的段落XML"""
    jc = f'<w:jc w:val="{align}"/>'
    spacing_parts = '<w:spacing w:line="400" w:lineRule="exact"/>'
    if spacing_before is not None or spacing_after is not None:
        sb = f' w:before="{spacing_before}"' if spacing_before else ''
        sa = f' w:after="{spacing_after}"' if spacing_after else ''
        spacing_parts = f'<w:spacing{sb}{sa} w:line="400" w:lineRule="exact"/>'
    indent = f'<w:ind w:firstLine="{first_indent}"/>' if first_indent else ''
    outline = f'<w:outlineLvl w:val="{outline_lvl}"/>' if outline_lvl is not None else ''
    widows = '<w:widowControl/>' if outline_lvl is not None else ''

    rpr = f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:cs="{font}"/><w:kern w:val="0"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>{"<w:b/><w:bCs/>" if bold else ""}'

    para_xml = f'''<w:p xmlns:w="{W}">
      <w:pPr>
        {widows}{spacing_parts}{indent}{jc}{outline}
        <w:rPr>{rpr}</w:rPr>
      </w:pPr>
      <w:r>
        <w:rPr>{rpr}</w:rPr>
        <w:t xml:space="preserve">{text}</w:t>
      </w:r>
    </w:p>'''
    return etree.fromstring(para_xml)

def make_chapter_title_xml(text):
    """创建章节标题（居中、16pt、黑体）"""
    return make_para_xml(text, sz=32, bold=False, font='黑体', align='center', first_indent=0, outline_lvl=None)

def make_section_title_xml(text, outline_lvl=1):
    """创建节标题（14pt、黑体、左对齐）"""
    return make_para_xml(text, sz=28, bold=False, font='黑体', align='left', first_indent=0,
                        spacing_before=200, spacing_after=100, outline_lvl=outline_lvl)

def make_subsection_title_xml(text):
    """创建子节标题（12pt、黑体加粗、左对齐）"""
    return make_para_xml(text, sz=24, bold=True, font='黑体', align='left', first_indent=0,
                        spacing_before=100, spacing_after=50, outline_lvl=2)

def make_body_xml(text, sz=24, first_indent=480):
    """创建正文段落"""
    return make_para_xml(text, sz=sz, bold=False, font='黑体', align='left', first_indent=first_indent)

def make_caption_xml(cn_text, en_text):
    """创建图题/表题"""
    rpr_cn = f'<w:rFonts w:ascii="黑体" w:eastAsia="黑体" w:hAnsi="黑体" w:cs="黑体"/><w:kern w:val="0"/><w:sz w:val="24"/><w:szCs w:val="24"/>'
    rpr_en = f'<w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:kern w:val="0"/><w:sz w:val="24"/><w:szCs w:val="24"/>'
    para_xml = f'''<w:p xmlns:w="{W}">
      <w:pPr>
        <w:spacing w:line="400" w:lineRule="exact"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r><w:rPr>{rpr_cn}</w:rPr><w:t xml:space="preserve">{cn_text}</w:t></w:r>
      <w:r><w:rPr>{rpr_en}</w:rPr><w:t xml:space="preserve"> {en_text}</w:t></w:r>
    </w:p>'''
    return etree.fromstring(para_xml)

def make_empty_para():
    """创建空段落"""
    return etree.fromstring(f'<w:p xmlns:w="{W}"/>')

def insert_image_para(doc, image_path, width_cm=14, caption_cn='', caption_en=''):
    """创建包含图片的段落，返回段落元素列表"""
    paragraphs = []

    # 图片段落
    img_para = etree.fromstring(f'<w:p xmlns:w="{W}"/>')
    pPr = etree.SubElement(img_para, qn('w:pPr'))
    jc = etree.SubElement(pPr, qn('w:jc'))
    jc.set(qn('w:val'), 'center')
    run = etree.SubElement(img_para, qn('w:r'))

    # 添加图片到文档并获取关系ID
    img_part = doc.part.get_or_add_image_part(image_path)
    rId = doc.part.relate_to(img_part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')

    # 获取图片尺寸
    from PIL import Image
    with Image.open(image_path) as img:
        px_w, px_h = img.size

    # 计算EMU（1 cm = 360000 EMU）
    target_width_emu = int(width_cm * 360000)
    px_per_emu = px_w / target_width_emu if target_width_emu > 0 else 1
    target_height_emu = int(px_h / px_per_emu)

    # 构建Drawing XML
    doc_prop_id = 100  # 简单的ID计数器
    cx = str(target_width_emu)
    cy = str(target_height_emu)

    drawing_xml = f'''
    <w:drawing xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}">
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="{cx}" cy="{cy}"/>
        <wp:docPr id="{doc_prop_id}" name="Picture"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="Picture.png"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="{rId}" xmlns:r="{R}"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>'''

    drawing = etree.fromstring(drawing_xml)
    run.append(drawing)
    paragraphs.append(img_para)

    # 图题
    if caption_cn or caption_en:
        cap_para = make_caption_xml(caption_cn, caption_en)
        paragraphs.append(cap_para)

    return paragraphs


def get_chart_file(pattern):
    """在charts/thesis目录中查找匹配的图表文件"""
    if not os.path.exists(CHART_DIR):
        return None
    for f in os.listdir(CHART_DIR):
        if pattern in f and f.endswith('.png'):
            return os.path.join(CHART_DIR, f)
    return None


def replace_paragraphs(doc, start_idx, end_idx, new_elements):
    """
    替换文档中 start_idx 到 end_idx (inclusive) 的段落
    new_elements: list of etree.Element (w:p)
    """
    body = doc.element.body
    all_paras = body.findall(qn('w:p'))

    # 获取要替换的段落
    to_remove = all_paras[start_idx:end_idx + 1]

    if not to_remove:
        return

    # 获取第一个被替换段落的前一个兄弟
    first = to_remove[0]
    ref_element = first.getprevious()

    # 获取最后一个被替换段落的后续非段落元素（如表格）
    last = to_remove[-1]
    next_sibling = last.getnext()

    # 收集要保留的后续兄弟元素（表格等）
    after_elements = []
    sibling = next_sibling
    while sibling is not None:
        tag = etree.QName(sibling.tag).localname if isinstance(sibling.tag, str) else ''
        # 也需要检查是否属于被替换范围
        if sibling in to_remove:
            sibling = sibling.getnext()
            continue
        after_elements.append(sibling)
        sibling = sibling.getnext()

    # 移除所有要替换的段落
    for elem in to_remove:
        body.remove(elem)

    # 插入新元素
    if ref_element is not None:
        for new_elem in new_elements:
            ref_element.addnext(new_elem)
            ref_element = new_elem
    else:
        # 如果没有前一个元素，插入到body开头
        for i, new_elem in enumerate(new_elements):
            body.insert(0, new_elem)

    print(f"  替换了段落 [{start_idx}]-[{end_idx}]，共 {len(to_remove)} 个 → {len(new_elements)} 个新元素")


def generate_chapter3_content(stats):
    """生成第三章（初始组）的内容元素列表"""
    elements = []

    # 章节标题
    elements.append(make_chapter_title_xml('第三章  竹粉/PLA/纳米SiO₂复合材料力学性能'))

    # 3.1 拉伸性能
    elements.append(make_section_title_xml('3.1 拉伸性能'))

    ts = stats['拉伸']['初始组']
    elements.append(make_body_xml(
        f'图3-1为未老化PLA/竹粉/纳米SiO₂复合材料的拉伸试验曲线。从图中可以看出，'
        f'添加不同含量纳米SiO₂后，复合材料的拉伸载荷-位移曲线呈现出明显差异。'
    ))
    elements.append(make_body_xml(
        f'如表3-1所示，未添加纳米SiO₂的复合材料（0%）的拉伸强度为{ts["0%"]["max_load_mean"]:.2f}±{ts["0%"]["max_load_std"]:.2f} N，'
        f'添加1%纳米SiO₂后拉伸强度为{ts["1%"]["max_load_mean"]:.2f}±{ts["1%"]["max_load_std"]:.2f} N，'
        f'添加3%纳米SiO₂后为{ts["3%"]["max_load_mean"]:.2f}±{ts["3%"]["max_load_std"]:.2f} N，'
        f'添加5%纳米SiO₂后为{ts["5%"]["max_load_mean"]:.2f}±{ts["5%"]["max_load_std"]:.2f} N。'
    ))
    elements.append(make_body_xml(
        f'添加纳米SiO₂后，复合材料的拉伸强度总体呈先下降后上升的趋势。1%和3%纳米SiO₂的添加使得拉伸强度有所降低，'
        f'这可能是因为纳米SiO₂颗粒在PLA基体中的分散不均匀导致应力集中。而当纳米SiO₂含量增加至5%时，'
        f'拉伸强度反而有所提高，达到了{ts["5%"]["max_load_mean"]:.2f} N，这可能是由于较高含量的纳米颗粒形成了更多的物理交联点，'
        f'增强了基体的承载能力。'
    ))

    # 插入拉伸曲线图（全图）
    chart_path = get_chart_file('第三章_拉伸_曲线_全')
    if chart_path and os.path.exists(chart_path):
        elements.append(make_empty_para())
        elements.extend(insert_image_para(None, chart_path, width_cm=14,
            caption_cn='图3-1', caption_en='Tensile Test Curves of BF/PLA/nano-SiO₂ Composites (Initial)'))
        elements.append(make_empty_para())

    # 插入单图
    for sio2 in ['0%', '1%', '3%', '5%']:
        chart_path = get_chart_file(f'第三章_拉伸_曲线_{sio2}')
        if chart_path and os.path.exists(chart_path):
            elements.append(make_empty_para())
            elements.extend(insert_image_para(None, chart_path, width_cm=12,
                caption_cn=f'图3-1-{sio2.strip("%")}纳米SiO₂复合材料拉伸试验曲线',
                caption_en=f'Tensile Curve of Composites with {sio2} nano-SiO₂'))
            elements.append(make_empty_para())

    # 3.1.1 断裂位移
    elements.append(make_subsection_title_xml('3.1.1 断裂位移分析'))
    elements.append(make_body_xml(
        f'由表3-2可以看出，0%SiO₂复合材料的断裂位移为{ts["0%"]["max_disp_mean"]:.2f}±{ts["0%"]["max_disp_std"]:.2f} mm，'
        f'1%SiO₂为{ts["1%"]["max_disp_mean"]:.2f}±{ts["1%"]["max_disp_std"]:.2f} mm，'
        f'3%SiO₂为{ts["3%"]["max_disp_mean"]:.2f}±{ts["3%"]["max_disp_std"]:.2f} mm，'
        f'5%SiO₂为{ts["5%"]["max_disp_mean"]:.2f}±{ts["5%"]["max_disp_std"]:.2f} mm。'
    ))
    elements.append(make_body_xml(
        f'随着纳米SiO₂含量的增加，复合材料的断裂位移呈先增大后减小的趋势。'
        f'1%SiO₂的断裂位移最大（{ts["1%"]["max_disp_mean"]:.2f} mm），表明适量的纳米SiO₂可以改善PLA基体的韧性。'
        f'这可能是由于纳米SiO₂粒子与PLA基体界面结合良好，能够有效传递应力，延缓裂纹扩展。'
    ))

    # 插入拉伸柱状图
    chart_path = get_chart_file('第三章_拉伸_强度柱状图')
    if chart_path and os.path.exists(chart_path):
        elements.append(make_empty_para())
        elements.extend(insert_image_para(None, chart_path, width_cm=12,
            caption_cn='图3-2', caption_en='Tensile Strength of BF/PLA/nano-SiO₂ Composites (Initial)'))
        elements.append(make_empty_para())

    chart_path = get_chart_file('第三章_拉伸_位移柱状图')
    if chart_path and os.path.exists(chart_path):
        elements.append(make_empty_para())
        elements.extend(insert_image_para(None, chart_path, width_cm=12,
            caption_cn='图3-3', caption_en='Displacement at Break of BF/PLA/nano-SiO₂ Composites (Initial)'))
        elements.append(make_empty_para())

    # 3.2 弯曲性能
    elements.append(make_section_title_xml('3.2 弯曲性能'))
    bs = stats['弯曲']['初始组']
    elements.append(make_body_xml(
        f'图3-4为未老化PLA/竹粉/纳米SiO₂复合材料的弯曲试验曲线。从图中可以看出，'
        f'不同纳米SiO₂含量的复合材料的弯曲载荷-位移曲线呈现出相似的弹性变形阶段，'
        f'但在最大载荷和断裂位移上存在差异。'
    ))
    elements.append(make_body_xml(
        f'如表3-3所示，0%SiO₂复合材料的弯曲强度为{bs["0%"]["max_load_mean"]:.2f}±{bs["0%"]["max_load_std"]:.2f} N，'
        f'1%SiO₂为{bs["1%"]["max_load_mean"]:.2f}±{bs["1%"]["max_load_std"]:.2f} N，'
        f'3%SiO₂为{bs["3%"]["max_load_mean"]:.2f}±{bs["3%"]["max_load_std"]:.2f} N，'
        f'5%SiO₂为{bs["5%"]["max_load_mean"]:.2f}±{bs["5%"]["max_load_std"]:.2f} N。'
    ))
    elements.append(make_body_xml(
        f'添加纳米SiO₂后，复合材料的弯曲强度有所提高。1%纳米SiO₂的弯曲强度最高，'
        f'达到{bs["1%"]["max_load_mean"]:.2f} N，比未添加组提高了{((bs["1%"]["max_load_mean"]/bs["0%"]["max_load_mean"]-1)*100):.1f}%。'
        f'这说明纳米SiO₂的加入有效增强了复合材料的弯曲承载能力，'
        f'纳米颗粒的填充效应和界面结合共同提高了材料的弯曲性能。'
    ))

    # 插入弯曲曲线图
    chart_path = get_chart_file('第三章_弯曲_曲线_全')
    if chart_path and os.path.exists(chart_path):
        elements.append(make_empty_para())
        elements.extend(insert_image_para(None, chart_path, width_cm=14,
            caption_cn='图3-4', caption_en='Flexural Test Curves of BF/PLA/nano-SiO₂ Composites (Initial)'))
        elements.append(make_empty_para())

    for sio2 in ['0%', '1%', '3%', '5%']:
        chart_path = get_chart_file(f'第三章_弯曲_曲线_{sio2}')
        if chart_path and os.path.exists(chart_path):
            elements.append(make_empty_para())
            elements.extend(insert_image_para(None, chart_path, width_cm=12,
                caption_cn=f'图3-4-{sio2.strip("%")}纳米SiO₂复合材料弯曲试验曲线',
                caption_en=f'Flexural Curve of Composites with {sio2} nano-SiO₂'))
            elements.append(make_empty_para())

    chart_path = get_chart_file('第三章_弯曲_强度柱状图')
    if chart_path and os.path.exists(chart_path):
        elements.append(make_empty_para())
        elements.extend(insert_image_para(None, chart_path, width_cm=12,
            caption_cn='图3-5', caption_en='Flexural Strength of BF/PLA/nano-SiO₂ Composites (Initial)'))
        elements.append(make_empty_para())

    chart_path = get_chart_file('第三章_弯曲_位移柱状图')
    if chart_path and os.path.exists(chart_path):
        elements.append(make_empty_para())
        elements.extend(insert_image_para(None, chart_path, width_cm=12,
            caption_cn='图3-6', caption_en='Flexural Displacement of BF/PLA/nano-SiO₂ Composites (Initial)'))
        elements.append(make_empty_para())

    # 3.3 本章小结
    elements.append(make_section_title_xml('3.3 本章小结'))
    elements.append(make_body_xml(
        f'本章研究了未老化状态下PLA/竹粉/纳米SiO₂复合材料的拉伸和弯曲力学性能。'
        f'主要结论如下：'
    ))
    elements.append(make_body_xml(
        f'（1）拉伸性能方面：0%SiO₂复合材料的拉伸强度为{ts["0%"]["max_load_mean"]:.2f} N，'
        f'5%SiO₂时达到最高值{ts["5%"]["max_load_mean"]:.2f} N。纳米SiO₂的添加对拉伸强度的影响呈现非线性关系，'
        f'适量添加可以提高拉伸强度，但分散均匀性是关键因素。'
    ))
    elements.append(make_body_xml(
        f'（2）弯曲性能方面：1%纳米SiO₂的弯曲强度最高（{bs["1%"]["max_load_mean"]:.2f} N），'
        f'比未添加组提高了{((bs["1%"]["max_load_mean"]/bs["0%"]["max_load_mean"]-1)*100):.1f}%。'
        f'纳米SiO₂的填充效应有效增强了复合材料的弯曲承载能力。'
    ))
    elements.append(make_body_xml(
        f'（3）综合来看，纳米SiO₂的添加对PLA/竹粉复合材料的力学性能有显著影响，'
        f'其中1%和5%的含量分别在弯曲强度和拉伸强度方面表现最优。'
        f'这为后续的老化性能研究提供了重要的基准数据。'
    ))

    return elements


def main():
    print("加载统计数据...")
    stats = load_stats()

    print("打开文档...")
    doc = Document(DOCX_PATH)

    # 先列出可用的图表
    print("\n可用图表：")
    chart_files = [f for f in os.listdir(CHART_DIR) if f.endswith('.png')]
    for f in sorted(chart_files):
        print(f"  {f}")

    # 先复制文档
    print(f"\n将修改保存到: {OUTPUT_PATH}")

    # 修改第三章 [125]-[161]
    print("\n========== 修改第三章 ==========")
    ch3_elements = generate_chapter3_content(stats)

    # 由于insert_image_para需要document对象来添加图片关系，
    # 我们改用不同的方法：先构建文本内容，图片部分用占位符标记，
    # 然后统一处理图片插入

    print(f"  生成了 {len(ch3_elements)} 个内容元素")

    # TODO: 实际替换段落（需要处理图片关系）
    # 由于python-docx的图片插入限制，使用C#脚本更可靠

    print("\n注意：由于python-docx对图片关系处理限制，"
          "建议使用minimax-docx C#脚本来完成实际的文档修改。")
    print("第三章内容已准备完毕，包含所有文字和图表引用。")

if __name__ == '__main__':
    main()
