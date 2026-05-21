#!/usr/bin/env python3
"""
change_to_wrap_top_bottom.py
将毕业论文_修改版.docx中所有 wp:inline 图片转换为 wp:anchor + wrapTopAndBottom（上下环绕型）
"""

import os
import copy
from lxml import etree
from docx import Document

# ============ 路径 ============
BASE_DIR = '/Users/shiberlin/Desktop/毕业论文'
DOCX_PATH = os.path.join(BASE_DIR, '毕业论文_修改版.docx')

# ============ 命名空间 ============
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# 确保序列化时保留命名空间前缀
etree.register_namespace('w', W)
etree.register_namespace('wp', WP)
etree.register_namespace('a', A)
etree.register_namespace('pic', PIC)
etree.register_namespace('r', REL)
etree.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
etree.register_namespace('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing')

def inline_to_anchor(inline_elem):
    """将一个 wp:inline 元素转换为 wp:anchor（上下环绕型），返回新的 anchor 元素"""

    # 提取关键属性和子元素
    extent = inline_elem.find(f'{{{WP}}}extent')
    cx = extent.get('cx', '0') if extent is not None else '0'
    cy = extent.get('cy', '0') if extent is not None else '0'

    docPr = inline_elem.find(f'{{{WP}}}docPr')
    cNvGraphicFramePr = inline_elem.find(f'{{{WP}}}cNvGraphicFramePr')
    graphic = inline_elem.find(f'{{{A}}}graphic')

    # 如果有必要，复制 graphicData 的 xmlns 声明到 pic:pic
    # python-docx 插入时 pic:pic 上可能没有 xmlns，但 graphicData 上有

    # 创建 anchor 元素
    anchor = etree.Element(f'{{{WP}}}anchor')
    anchor.set('distT', '0')
    anchor.set('distB', '0')
    anchor.set('distL', '114300')   # 0.1cm 左边距
    anchor.set('distR', '114300')   # 0.1cm 右边距
    anchor.set('simplePos', '0')
    anchor.set('relativeHeight', '251659264')
    anchor.set('behindDoc', '0')
    anchor.set('locked', '0')
    anchor.set('layoutInCell', '1')
    anchor.set('allowOverlap', '1')

    # <wp:simplePos x="0" y="0"/>
    simplePos = etree.SubElement(anchor, f'{{{WP}}}simplePos')
    simplePos.set('x', '0')
    simplePos.set('y', '0')

    # <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
    positionH = etree.SubElement(anchor, f'{{{WP}}}positionH')
    positionH.set('relativeFrom', 'column')
    alignH = etree.SubElement(positionH, f'{{{WP}}}align')
    alignH.text = 'center'

    # <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
    positionV = etree.SubElement(anchor, f'{{{WP}}}positionV')
    positionV.set('relativeFrom', 'paragraph')
    posOffset = etree.SubElement(positionV, f'{{{WP}}}posOffset')
    posOffset.text = '0'

    # <wp:extent cx="..." cy="..."/>
    new_extent = etree.SubElement(anchor, f'{{{WP}}}extent')
    new_extent.set('cx', cx)
    new_extent.set('cy', cy)

    # <wp:effectExtent l="0" t="0" r="0" b="0"/>
    effectExtent = etree.SubElement(anchor, f'{{{WP}}}effectExtent')
    effectExtent.set('l', '0')
    effectExtent.set('t', '0')
    effectExtent.set('r', '0')
    effectExtent.set('b', '0')

    # <wp:wrapTopAndBottom wrapText="bothSides"/>
    wrap = etree.SubElement(anchor, f'{{{WP}}}wrapTopAndBottom')
    wrap.set('wrapText', 'bothSides')

    # <wp:docPr .../> (deepcopy)
    if docPr is not None:
        anchor.append(copy.deepcopy(docPr))

    # <wp:cNvGraphicFramePr .../> (deepcopy)
    if cNvGraphicFramePr is not None:
        anchor.append(copy.deepcopy(cNvGraphicFramePr))

    # <a:graphic>...</a:graphic> (deepcopy)
    if graphic is not None:
        anchor.append(copy.deepcopy(graphic))

    return anchor


def main():
    print(f"正在打开文档: {DOCX_PATH}")
    doc = Document(DOCX_PATH)
    body = doc.element.body

    # 查找所有 wp:inline 元素
    inline_elements = body.findall(f'.//{{{WP}}}inline')
    print(f"找到 {len(inline_elements)} 个嵌入型图片")

    if len(inline_elements) == 0:
        print("没有找到需要转换的图片，退出。")
        return

    # 逐个转换
    converted = 0
    for inline in inline_elements:
        parent = inline.getparent()
        anchor = inline_to_anchor(inline)

        # 在 inline 的位置插入 anchor，然后删除 inline
        parent.replace(inline, anchor)
        converted += 1

    # 保存
    doc.save(DOCX_PATH)
    print(f"\n成功转换 {converted} 个图片为上下环绕型")
    print(f"文档已保存: {DOCX_PATH}")


if __name__ == '__main__':
    main()
