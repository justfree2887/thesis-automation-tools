"""
第七章完整重构脚本 - 按步骤执行
从后往前修改避免索引偏移
"""

from lxml import etree
from docx import Document
from copy import deepcopy
import re

doc_path = '/Users/shiberlin/Desktop/毕业论文1.2.docx'
doc = Document(doc_path)
body = doc.element.body
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def get_text(child):
    texts = child.findall(f'.//{W}t')
    return ''.join(t.text or '' for t in texts).strip()

def replace_in_text(elem, old, new):
    """替换元素内所有文字中的子串"""
    for t in elem.findall(f'.//{W}t'):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)

def make_para(text):
    """创建正文段落（小四号，宋体）"""
    p = etree.Element(f'{W}p')
    pPr = etree.SubElement(p, f'{W}pPr')
    spacing = etree.SubElement(pPr, f'{W}spacing')
    spacing.set(f'{W}line', '360')
    spacing.set(f'{W}lineRule', 'auto')
    # 首行缩进
    ind = etree.SubElement(pPr, f'{W}ind')
    ind.set(f'{W}firstLineChars', '200')
    ind.set(f'{W}firstLine', '480')
    
    r = etree.SubElement(p, f'{W}r')
    rPr = etree.SubElement(r, f'{W}rPr')
    rFonts = etree.SubElement(rPr, f'{W}rFonts')
    rFonts.set(f'{W}eastAsia', '宋体')
    rFonts.set(f'{W}ascii', 'Times New Roman')
    rFonts.set(f'{W}hAnsi', 'Times New Roman')
    sz = etree.SubElement(rPr, f'{W}sz')
    sz.set(f'{W}val', '24')
    szCs = etree.SubElement(rPr, f'{W}szCs')
    szCs.set(f'{W}val', '24')
    t = etree.SubElement(r, f'{W}t')
    t.text = text
    # 确保中文字体生效
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return p

def make_heading(text, bold=True):
    """创建二级标题段落（四号黑体加粗）"""
    p = etree.Element(f'{W}p')
    pPr = etree.SubElement(p, f'{W}pPr')
    spacing = etree.SubElement(pPr, f'{W}spacing')
    spacing.set(f'{W}before', '240')
    spacing.set(f'{W}after', '120')
    
    r = etree.SubElement(p, f'{W}r')
    rPr = etree.SubElement(r, f'{W}rPr')
    rFonts = etree.SubElement(rPr, f'{W}rFonts')
    rFonts.set(f'{W}eastAsia', '黑体')
    rFonts.set(f'{W}ascii', 'Times New Roman')
    rFonts.set(f'{W}hAnsi', 'Times New Roman')
    sz = etree.SubElement(rPr, f'{W}sz')
    sz.set(f'{W}val', '28')  # 四号 = 14pt
    szCs = etree.SubElement(rPr, f'{W}szCs')
    szCs.set(f'{W}val', '28')
    if bold:
        b = etree.SubElement(rPr, f'{W}b')
    t = etree.SubElement(r, f'{W}t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return p

def make_caption(text, center=True):
    """创建图/表标题段落"""
    p = etree.Element(f'{W}p')
    pPr = etree.SubElement(p, f'{W}pPr')
    if center:
        jc = etree.SubElement(pPr, f'{W}jc')
        jc.set(f'{W}val', 'center')
    spacing = etree.SubElement(pPr, f'{W}spacing')
    spacing.set(f'{W}before', '120')
    spacing.set(f'{W}after', '120')
    
    r = etree.SubElement(p, f'{W}r')
    rPr = etree.SubElement(r, f'{W}rPr')
    rFonts = etree.SubElement(rPr, f'{W}rFonts')
    rFonts.set(f'{W}eastAsia', '宋体')
    rFonts.set(f'{W}ascii', 'Times New Roman')
    rFonts.set(f'{W}hAnsi', 'Times New Roman')
    sz = etree.SubElement(rPr, f'{W}sz')
    sz.set(f'{W}val', '21')  # 五号 = 10.5pt
    szCs = etree.SubElement(rPr, f'{W}szCs')
    szCs.set(f'{W}val', '21')
    t = etree.SubElement(r, f'{W}t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return p

def insert_elements_before(body, ref_index, elements):
    """在ref_index位置前插入多个元素"""
    ref = body[ref_index]
    for elem in elements:
        ref.addprevious(elem)

# ================================================================
# PHASE 1: 备份所有需要移动的元素
# ================================================================
print("=== PHASE 1: 备份元素 ===")

# FTIR图元素 (每个 = 图片段落 + 中文标题 + 英文标题)
# 注意：Ch7中FTIR图的顺序是 7-5(初始)→7-6(湿热)→7-7(紫外)→7-8(耦合)
# 但我们要移到：初始→Ch3, 紫外→Ch4, 湿热→Ch5, 耦合→Ch6

ftir_to_ch3 = [
    deepcopy(body[389]),  # 图7-5 图片
    deepcopy(body[390]),  # 图7-5 中文标题
    deepcopy(body[391]),  # Fig.7-5 英文标题
]

ftir_to_ch4 = [
    deepcopy(body[397]),  # 图7-7 图片
    deepcopy(body[398]),  # 图7-7 中文标题
    deepcopy(body[399]),  # Fig.7-7 英文标题
]

ftir_to_ch5 = [
    deepcopy(body[393]),  # 图7-6 图片
    deepcopy(body[394]),  # 图7-6 中文标题
    deepcopy(body[395]),  # Fig.7-6 英文标题
]

ftir_to_ch6 = [
    deepcopy(body[401]),  # 图7-8 图片
    deepcopy(body[402]),  # 图7-8 中文标题
    deepcopy(body[403]),  # Fig.7-8 英文标题
]

# 修改FTIR图的编号
for elem in ftir_to_ch3:
    replace_in_text(elem, '7-5', '3-7')
for elem in ftir_to_ch4:
    replace_in_text(elem, '7-7', '4-9')
for elem in ftir_to_ch5:
    replace_in_text(elem, '7-6', '5-9')
for elem in ftir_to_ch6:
    replace_in_text(elem, '7-8', '6-9')

print("FTIR图备份并编号完成")

# ================================================================
# PHASE 2: 备份Ch7中保留的元素（用于重建新Ch7）
# ================================================================
print("\n=== PHASE 2: 备份Ch7保留元素 ===")

# 对比图
fig_7_1 = [deepcopy(body[351]), deepcopy(body[352]), deepcopy(body[353])]  # 拉伸
fig_7_2 = [deepcopy(body[354]), deepcopy(body[355]), deepcopy(body[356])]  # 弯曲
fig_7_3 = [deepcopy(body[357]), deepcopy(body[358]), deepcopy(body[359])]  # 冲击

# 保留率图 - body[376]是图片，body[381]是中文标题，body[382]是英文标题
fig_7_4 = [deepcopy(body[376]), deepcopy(body[381]), deepcopy(body[382])]

# 弯曲模量对比图 body[417]是图片(在7.6节中), body[422]cn, body[423]en
fig_7_9 = [deepcopy(body[417]), deepcopy(body[422]), deepcopy(body[423])]

# 冲击韧性对比图 body[421]是图片, body[426]cn, body[427]en
fig_7_10 = [deepcopy(body[421]), deepcopy(body[426]), deepcopy(body[427])]

# 表格（保留所有6个表格及其标题）
# 表7-1: body[360]引言 + body[361]表格 + body[362]标题
tbl_7_1 = [deepcopy(body[361]), deepcopy(body[362])]
tbl_7_2 = [deepcopy(body[367]), deepcopy(body[368])]
tbl_7_3 = [deepcopy(body[370]), deepcopy(body[371])]
tbl_7_4 = [deepcopy(body[372]), deepcopy(body[373])]
tbl_7_5 = [deepcopy(body[374]), deepcopy(body[375])]
tbl_7_6 = [deepcopy(body[433]), deepcopy(body[432]), deepcopy(body[436])]  # 表格+标题+图片

# 重新编号保留的图
for elem in fig_7_9:
    replace_in_text(elem, '7-9', '7-5')
for elem in fig_7_10:
    replace_in_text(elem, '7-10', '7-6')
for elem in tbl_7_6:
    replace_in_text(elem, '7-11', '7-7')
    replace_in_text(elem, '7-6', '7-6')  # table caption says 表7-6, keep

print("Ch7保留元素备份完成")

# ================================================================
# PHASE 3: 修改第六章 - 插入6.4节 + 重写小结 (从后往前先改Ch6)
# ================================================================
print("\n=== PHASE 3: 修改第六章 ===")

# Ch6小结在 body[344]-[346]（body[347]是第七章标题）
# 先在body[344]前插入6.4节内容

ch6_64_elements = [
    make_heading('6.4 紫外湿热耦合老化机理分析'),
    make_para(
        '图6-9为紫外湿热耦合老化后竹粉/PLA/纳米SiO₂复合材料的红外光谱图。'
        '耦合老化后的红外光谱呈现出介于纯紫外老化和纯湿热老化之间的特征，'
        '1750 cm⁻¹处C=O峰有所增强，1710 cm⁻¹处的羧基峰也清晰可见，'
        '但强度低于纯湿热老化组。这表明耦合老化中PLA基体同时经历了光氧化和水解两种降解过程，'
        '但耦合老化试验中试样并未完全浸没于水中，实际参与水解的水量低于纯水浴条件，'
        '因此水解程度相对较轻。[3,6]'
    ),
]
ch6_64_elements.extend(ftir_to_ch6)  # 插入图6-9

ch6_64_elements.extend([
    make_para(
        '理论上，紫外光和湿热条件应产生协同加速效应。水分子可以促进紫外光引发的自由基反应'
        '（作为反应介质），紫外光导致的分子链断裂又增加了水分渗透的通道。'
        '然而，在本实验中，由于耦合老化是在紫外老化箱中以冷凝方式提供水分，'
        '试样表面的实际含水量低于纯水浴条件，水解反应程度较轻，'
        '导致整体性能下降介于两种单一老化之间。如果延长老化时间或提高湿度条件，'
        '协同加速效应可能会更加明显。'
    ),
    make_para(
        '纳米SiO₂的防护机理主要包括三个方面：（1）紫外屏蔽效应——纳米SiO₂粒子散射和吸收紫外光，'
        '减少对基体的直接辐照；（2）阻隔效应——纳米粒子在基体中形成曲折的渗透路径，'
        '延缓水分向材料内部扩散；（3）界面增强效应——SiO₂表面的硅羟基与PLA形成氢键，'
        '提高基体致密性和界面结合强度。[22]'
    ),
    make_para(
        '在弯曲模量方面，耦合老化后5%SiO₂组的弯曲模量上升幅度最大，达24.0%'
        '（4.764→5.908 GPa）。3%SiO₂组模量也略有升高（+5.7%），'
        '而0%和1%组模量均下降。耦合老化同时具备热和紫外两个促进因素，'
        'PLA的二次结晶和光诱导轻度交联效应叠加，使得高SiO₂含量组的模量上升最为显著。'
        '然而需要指出的是，模量的上升并不意味着材料整体性能的改善——5%组的拉伸强度'
        '和冲击强度在耦合老化后仍然下降，说明这种"增强"主要局限于刚度的提升。'
    ),
    make_heading('6.5 本章小结'),
    make_para(
        '本章研究了紫外湿热耦合老化对PLA/竹粉/纳米SiO₂复合材料力学性能的影响。主要结论如下：'
    ),
    make_para(
        '（1）经紫外湿热耦合老化后，复合材料的拉伸强度和弯曲强度均有所下降，'
        '说明紫外与湿热的协同作用对材料力学性能产生了显著的负面影响。'
        '三种老化条件对复合材料性能的影响程度为：湿热老化 > 耦合老化 > 紫外老化。'
    ),
    make_para(
        '（2）添加纳米SiO₂后，复合材料在紫外湿热耦合老化条件下的力学性能保留率有所提高，'
        '表明纳米SiO₂的加入有助于提升复合材料在复杂老化环境下的耐久性能。'
        '纳米SiO₂通过紫外屏蔽、水分阻隔和界面增强三种机制协同保护复合材料。'
    ),
    make_para(
        '（3）红外光谱分析表明，耦合老化后PLA基体同时发生了光氧化和水解反应，'
        '但水解程度低于纯湿热老化条件。5%SiO₂组弯曲模量在耦合老化后上升24.0%，'
        '归因于SiO₂促进PLA二次结晶和紫外光诱导交联的"补偿效应"。'
        '综合来看，耦合老化的降解程度介于紫外老化和湿热老化之间，'
        'SiO₂的添加有效缓解了复合材料在复杂老化环境下的性能退化。[48]'
    ),
])

# 删除旧小结 body[344]-[346]
old_ch6_summary = [body[344], body[345], body[346]]
for elem in old_ch6_summary:
    body.remove(elem)

# 在原位置（body[344]现在指向原来的body[347]即第七章标题之前）插入新内容
insert_elements_before(body, 344, ch6_64_elements)

print("第六章修改完成：新增6.4机理分析 + 6.5小结")

# ================================================================
# PHASE 4: 修改第五章
# ================================================================
print("\n=== PHASE 4: 修改第五章 ===")

# Ch5小结在 body[286]-[289]（body[291]是第六章标题）
# 先在body[286]前插入5.4节

ch5_54_elements = [
    make_heading('5.4 湿热老化机理分析'),
    make_para(
        '图5-9为湿热老化后竹粉/PLA/纳米SiO₂复合材料的红外光谱图。'
        '经湿热老化后，1750 cm⁻¹处的C=O吸收峰强度明显增强，'
        '同时出现了约1710 cm⁻¹的新吸收峰，该峰归属于PLA水解产生的羧基（-COOH）中C=O的伸缩振动。'
        '此外，1600 cm⁻¹和1510 cm⁻¹附近的木质素芳环吸收峰强度也有所增加，'
        '表明竹粉中的木质素在湿热条件下也发生了结构变化。'
        '水分渗入复合材料后不仅导致PLA基体的酯键水解，还可能引起竹粉中半纤维素的水解溶出'
        '和木质素的缩合反应，进一步加剧了复合材料的性能退化。[48][6]'
    ),
]
ch5_54_elements.extend(ftir_to_ch5)  # 插入图5-9

ch5_54_elements.extend([
    make_para(
        '湿热老化的核心机理包括两个并行过程：一是PLA的水解反应，'
        '水分子攻击PLA分子链上的酯键，发生无规断链反应，生成低分子量的醇和酸端基；'
        '二是竹粉的吸水膨胀，竹粉中的亲水性羟基吸收水分后发生膨胀，体积增大，'
        '在纤维-基体界面产生内应力，导致微裂纹的萌生和界面脱粘。'
        '这两种反应相互促进——PLA水解使基体变弱，界面脱粘又为水分渗透提供了新的通道，'
        '形成恶性循环。[7,51]'
    ),
    make_para(
        '在弯曲模量方面，湿热老化后5%SiO₂组的弯曲模量上升10.2%'
        '（4.764→5.250 GPa）。PLA是一种半结晶型聚合物，在60 ℃湿热条件下，'
        '非晶区的分子链获得了足够的运动能力向更稳定的结晶态转变。'
        '5%高添加量下大量纳米SiO₂粒子提供了充足的异相成核位点，'
        '加速了老化过程中PLA的结晶度提高。此外，水解断链产生的短链段运动能力更强，'
        '更容易利用SiO₂成核位点迅速结晶，形成大量微晶，补偿了水解对基体的弱化效应。'
        '3%SiO₂组模量也略有升高（+1.8%），而0%和1%组缺乏足够成核位点，模量变化不大。'
    ),
    make_heading('5.5 本章小结'),
    make_para(
        '本章研究了不同时间湿热老化对PLA/竹粉/纳米SiO₂复合材料力学性能的影响。主要结论如下：'
    ),
    make_para(
        '（1）经5天和10天湿热老化后，复合材料的拉伸强度和弯曲强度均随老化时间延长而进一步下降，'
        '说明湿热老化对材料力学性能的劣化具有时间累积效应。'
    ),
    make_para(
        '（2）添加纳米SiO₂后，复合材料在5天和10天湿热老化后的力学性能保留率均高于未添加组，'
        '表明纳米SiO₂的加入有效提升了复合材料的耐湿热老化性能，'
        '且这种保护作用在更长老化时间下依然显著。'
    ),
    make_para(
        '（3）红外光谱分析表明，湿热老化导致PLA发生明显的水解反应，'
        '产生1710 cm⁻¹的羧基特征峰，同时竹粉中木质素也发生结构变化。'
        '纳米SiO₂的"迷宫效应"有效延缓了水分向材料内部的渗透，'
        '5%SiO₂组弯曲模量上升10.2%，归因于SiO₂促进PLA二次结晶的补偿效应。[6,8]'
    ),
])

# 删除旧小结 body[286]-[289]
old_ch5_summary = [body[286], body[287], body[288], body[289]]
for elem in old_ch5_summary:
    body.remove(elem)

insert_elements_before(body, 286, ch5_54_elements)

print("第五章修改完成：新增5.4机理分析 + 5.5小结")

# ================================================================
# PHASE 5: 修改第四章
# ================================================================
print("\n=== PHASE 5: 修改第四章 ===")

# Ch4小结在 body[226]-[229]（body[231]是第五章标题）
ch4_44_elements = [
    make_heading('4.4 紫外老化机理分析'),
    make_para(
        '图4-9为紫外老化后竹粉/PLA/纳米SiO₂复合材料的红外光谱图。'
        '紫外老化后的红外光谱变化相对较小，1750 cm⁻¹处C=O峰的强度变化不明显，'
        '仅在约1710 cm⁻¹处出现了微弱的肩峰，说明紫外老化导致的PLA光氧化降解程度较轻。'
        '这与力学性能测试结果一致——紫外老化对复合材料力学性能的影响最小。[3,5]'
    ),
]
ch4_44_elements.extend(ftir_to_ch4)  # 插入图4-9

ch4_44_elements.extend([
    make_para(
        '紫外老化的核心机理为：紫外光（UVA-340，295-365 nm）被PLA基体吸收后，'
        '激发分子产生自由基，引发链式光氧化反应。主要反应路径包括：'
        'PLA分子链上的叔碳原子脱氢形成过氧自由基，过氧自由基进一步分解产生羰基、'
        '羧基等含氧官能团，导致分子链断裂和分子量降低。'
        '然而，在本实验条件下（168 h，50 ℃冷凝温度），'
        '光氧化降解程度较轻，主要局限于材料表层，因此力学性能下降幅度较小。[50,51]'
    ),
    make_para(
        '在弯曲模量方面，紫外老化后5%SiO₂组的弯曲模量上升幅度达18.4%'
        '（4.764→5.638 GPa），3%SiO₂组也略有升高（+1.8%），'
        '而0%和1%组模量均下降或基本不变。'
        '这归因于两个因素：一是老化试验箱内灯管温度较高（约60-70 ℃），'
        '纳米SiO₂粒子表面作为异相成核位点促进PLA二次结晶，微晶区域的增加直接提升了材料刚度；'
        '二是紫外光诱导PLA发生轻度交联，SiO₂表面的硅羟基和缺陷位点可能对自由基反应有催化作用，'
        '加速交联进程。高SiO₂含量下两种效应叠加，使得模量上升最为显著。'
    ),
    make_heading('4.5 本章小结'),
    make_para(
        '本章研究了不同时间紫外老化对PLA/竹粉/纳米SiO₂复合材料力学性能的影响。主要结论如下：'
    ),
    make_para(
        '（1）经5天和10天紫外老化后，复合材料的拉伸强度和弯曲强度均随老化时间延长而进一步下降，'
        '说明紫外老化对材料力学性能的劣化具有时间累积效应，'
        '但整体退化程度较湿热老化轻得多。'
    ),
    make_para(
        '（2）添加纳米SiO₂后，复合材料在5天和10天紫外老化后的力学性能保留率均高于未添加组，'
        '表明纳米SiO₂的加入有效提升了复合材料的耐紫外老化性能，'
        '这主要归因于纳米颗粒的紫外线屏蔽效应和界面增强效应。'
    ),
    make_para(
        '（3）红外光谱分析表明，紫外老化对PLA化学结构的破坏较轻，仅在1710 cm⁻¹处出现微弱肩峰。'
        '5%SiO₂组弯曲模量在紫外老化后上升18.4%，归因于SiO₂促进PLA二次结晶'
        '和紫外光诱导轻度交联的"补偿效应"。[3,7]'
    ),
])

# 删除旧小结 body[226]-[229]
old_ch4_summary = [body[226], body[227], body[228], body[229]]
for elem in old_ch4_summary:
    body.remove(elem)

insert_elements_before(body, 226, ch4_44_elements)

print("第四章修改完成：新增4.4机理分析 + 4.5小结")

# ================================================================
# PHASE 6: 修改第三章
# ================================================================
print("\n=== PHASE 6: 修改第三章 ===")

# Ch3小结在 body[165]-[169]（body[165]是"3.3 本章小结"标题）
ch3_34_elements = [
    make_heading('3.4 红外光谱分析'),
    make_para(
        '图3-7为未老化竹粉/PLA/纳米SiO₂复合材料的红外光谱图。'
        '在初始试样的红外光谱中，可以观察到PLA的特征吸收峰：'
        '1750 cm⁻¹附近的C=O伸缩振动峰、1180 cm⁻¹和1085 cm⁻¹附近的C-O-C伸缩振动峰、'
        '以及1450 cm⁻¹附近的-CH₃变形振动峰。'
        '此外，3340 cm⁻¹附近的宽吸收峰对应于竹粉中纤维素的-OH伸缩振动。[49]'
        '随着SiO₂含量的增加，1100 cm⁻¹附近Si-O-Si伸缩振动峰的强度逐渐增强，'
        '证实了纳米SiO₂在复合材料中的存在。'
        '含SiO₂的复合材料在1750 cm⁻¹处的C=O峰强度略低于纯PLA组，'
        '这可能是因为SiO₂粒子对PLA分子链运动的限制作用在一定程度上抑制了酯键的氧化。'
    ),
]
ch3_34_elements.extend(ftir_to_ch3)  # 插入图3-7

ch3_34_elements.extend([
    make_heading('3.5 本章小结'),
    make_para(
        '本章研究了未老化状态下PLA/竹粉/纳米SiO₂复合材料的力学性能及红外光谱特征。主要结论如下：'
    ),
    make_para(
        '（1）拉伸性能方面：0%SiO₂复合材料的拉伸强度为2424.84 N，'
        '5%SiO₂时达到最高值2604.03 N。纳米SiO₂的添加对拉伸强度的影响呈现非线性关系，'
        '适量添加可以提高拉伸强度，但分散均匀性是关键因素。'
    ),
    make_para(
        '（2）弯曲性能方面：1%纳米SiO₂的弯曲强度最高（193.50 N），'
        '比未添加组提高了10.6%。纳米SiO₂的填充效应有效增强了复合材料的弯曲承载能力。'
    ),
    make_para(
        '（3）红外光谱分析表明，复合材料中PLA基体的特征吸收峰清晰可见，'
        '1100 cm⁻¹处的Si-O-Si振动峰证实了纳米SiO₂的成功引入。'
        'SiO₂对PLA分子链运动的限制作用有助于延缓基体的氧化降解，'
        '为后续老化性能研究提供了结构层面的解释依据。[12,13]'
    ),
])

# 删除旧小结 body[165]-[169]
old_ch3_summary = [body[165], body[166], body[167], body[168], body[169]]
for elem in old_ch3_summary:
    body.remove(elem)

insert_elements_before(body, 165, ch3_34_elements)

print("第三章修改完成：新增3.4红外光谱分析 + 3.5小结")

# ================================================================
# PHASE 7: 删除旧第七章，重建新第七章
# ================================================================
print("\n=== PHASE 7: 重建第七章 ===")

# 重新查找第七章标题的当前位置（前面的插入可能导致索引变化）
ch7_title_idx = None
for i, child in enumerate(list(body)):
    if etree.QName(child.tag).localname == 'p':
        text = get_text(child)
        if '第七章' in text and '综合' in text:
            ch7_title_idx = i
            break
        elif '第七章' in text:
            ch7_title_idx = i
            break

if ch7_title_idx is None:
    # 尝试找其他特征
    for i, child in enumerate(list(body)):
        if etree.QName(child.tag).localname == 'p':
            text = get_text(child)
            if text.startswith('第七章'):
                ch7_title_idx = i
                break

print(f"第七章标题位置: body[{ch7_title_idx}] = {get_text(body[ch7_title_idx])[:50]}")

# 找第八章标题位置
ch8_title_idx = None
for i, child in enumerate(list(body)):
    if etree.QName(child.tag).localname == 'p':
        text = get_text(child)
        if '第八章' in text:
            ch8_title_idx = i
            break

print(f"第八章标题位置: body[{ch8_title_idx}] = {get_text(body[ch8_title_idx])[:50]}")

# 删除旧第七章所有内容（从Ch7标题到Ch8标题之前）
elements_to_remove = []
for i in range(ch7_title_idx, ch8_title_idx):
    elements_to_remove.append(body[i])

for elem in elements_to_remove:
    body.remove(elem)

print(f"已删除旧第七章 {len(elements_to_remove)} 个元素")

# 重新获取Ch8标题位置（删除后索引变了）
ch8_new_idx = None
for i, child in enumerate(list(body)):
    if etree.QName(child.tag).localname == 'p':
        text = get_text(child)
        if '第八章' in text:
            ch8_new_idx = i
            break

print(f"新第八章位置: body[{ch8_new_idx}]")

# ================================================================
# 构建新第七章内容
# ================================================================
new_ch7 = []

# 章标题
new_ch7.append(make_heading('第七章 不同老化条件下复合材料性能对比分析', bold=True))

# --- 7.1 拉伸性能对比 ---
new_ch7.append(make_heading('7.1 不同老化条件下拉伸性能对比分析'))
new_ch7.append(make_para(
    '图7-1为不同老化条件下竹粉/PLA/纳米SiO₂复合材料的拉伸强度对比。'
    '可以看出，三种老化条件对拉伸强度的影响程度存在显著差异：'
    '湿热老化导致的拉伸强度下降最为严重，紫外老化的影响最小，耦合老化介于两者之间。'
    '表7-1列出了各条件下拉伸强度的完整数据。'
))
new_ch7.extend(fig_7_1)  # 图7-1
new_ch7.extend(tbl_7_1)  # 表7-1
new_ch7.append(make_para(
    '从纳米SiO₂的防护效果来看，在湿热老化和耦合老化条件下，'
    '纳米SiO₂的添加均表现出一定的性能保护作用。'
    '在湿热老化条件下，0%SiO₂组的拉伸强度保留率仅为76.7%，'
    '而5%SiO₂组保持在83.4%；耦合老化条件下，'
    '1%SiO₂组的保留率最高（93.4%），但5%SiO₂组反而降至83.2%，'
    '这是因为过量的纳米粒子团聚形成了应力集中区域，抵消了部分防护效果。'
    '紫外老化条件下各组的保留率差异较小，均在92%以上，'
    '表明紫外老化对复合材料拉伸性能的影响相对温和。'
))

# --- 7.2 弯曲性能对比 ---
new_ch7.append(make_heading('7.2 不同老化条件下弯曲性能对比分析'))
new_ch7.append(make_para(
    '图7-2和图7-5分别为不同老化条件下复合材料的弯曲强度和弯曲模量对比。'
    '表7-2和表7-4列出了完整数据。弯曲强度的变化趋势与拉伸强度类似，'
    '湿热老化影响最大，紫外老化影响最小。'
))
new_ch7.extend(fig_7_2)  # 图7-2
new_ch7.extend(tbl_7_2)  # 表7-2
new_ch7.append(make_para(
    '值得注意的是，对于弯曲强度，5%SiO₂组在耦合老化后的保留率超过了100%，'
    '显示出优异的抗老化性能。在弯曲模量方面，5%SiO₂组在三种老化条件下均表现出上升趋势：'
    '紫外老化后上升18.4%，湿热老化后上升10.2%，耦合老化后上升幅度最大达24.0%。'
    '这归因于高含量纳米SiO₂促进PLA二次结晶和光诱导交联的"补偿效应"。'
    '然而，弯曲模量的上升并不意味着材料整体性能的改善——5%组的拉伸强度在老化后仍然下降，'
    '说明这种"增强"主要局限于刚度的提升。'
))
new_ch7.extend(tbl_7_4)  # 表7-4 弯曲模量
new_ch7.extend(fig_7_9)  # 图7-5 (was 7-9) 弯曲模量对比

# --- 7.3 冲击性能对比 ---
new_ch7.append(make_heading('7.3 不同老化条件下冲击性能对比'))
new_ch7.append(make_para(
    '图7-3、图7-4和图7-6分别为不同老化条件下复合材料的冲击强度、力学性能保留率和冲击韧性对比。'
    '表7-3和表7-5列出了完整数据。与拉伸和弯曲性能不同，'
    '冲击强度在老化后并未单调下降——在湿热和耦合老化条件下，'
    '部分组的冲击强度反而有所升高，表现出"增韧"效应。'
))
new_ch7.extend(fig_7_3)  # 图7-3
new_ch7.extend(tbl_7_3)  # 表7-3
new_ch7.extend(fig_7_4)  # 图7-4 保留率
new_ch7.append(make_para(
    '从保留率来看，冲击强度在所有老化条件下均达到96%以上，'
    '部分组甚至超过100%，这与前文分析的"增韧"效应一致。'
    '冲击韧性随老化条件的变化趋势显示，随着老化程度从"初始→紫外→耦合→湿热"递增，'
    '拉伸强度和弯曲强度呈整体下降趋势，而冲击强度在湿热和耦合老化后反而升高。'
))
new_ch7.extend(tbl_7_5)  # 表7-5
new_ch7.extend(fig_7_10)  # 图7-6 (was 7-10) 冲击韧性

# --- 7.4 吸湿性能对比 ---
new_ch7.append(make_heading('7.4 吸湿性能对比'))
new_ch7.append(make_para(
    '为评估纳米SiO₂改性PLA基生物质复合材料的吸湿性能，'
    '参照GB/T 1034-2008标准进行了吸湿试验。将质量分数为20%的木粉与不同含量'
    '（0%、1%、3%、5%）的纳米SiO₂及PLA基体混合后，经注塑成型制备标准试样，'
    '每组设置4个平行样。试验前将试样置于60°C干燥箱中干燥至恒重，记录初始质量M₀，'
    '然后将试样浸泡于蒸馏水中，在室温条件下分别于0、6、18、42、66、114、138 h取出，'
    '擦干表面水分后称量湿态质量Mt，按公式ΔM=(Mt−M₀)/M₀×100%计算吸湿率。'
))
new_ch7.append(make_para(
    '各组复合材料的吸湿率随时间变化数据如表7-6所示。可以看出，在138 h浸泡周期内，'
    '所有试样的吸湿率均随时间延长而增大，但增长速率逐渐趋缓，表现出典型的Fickian扩散特征。'
    '其中，3%纳米SiO₂含量组的吸湿率最高，138 h吸湿率达到4.10%，'
    '远高于纯木粉/PLA复合材料的1.94%。'
))
new_ch7.extend(tbl_7_6)  # 表7-6 + 图7-7
new_ch7.append(make_para(
    '这一现象可归因于纳米SiO₂在复合材料中的分散效应。当SiO₂含量为3%时，'
    '纳米颗粒在基体中形成较为均匀的分散网络，颗粒间的界面间隙为水分渗透提供了额外的通道，'
    '导致吸湿率升高。而当SiO₂含量增至5%时，纳米颗粒发生部分团聚，'
    '反而形成了更为致密的阻隔结构，抑制了水分的扩散路径，使吸湿率降至1.50%。'
    '这表明纳米SiO₂含量存在一个临界值，超过该值后颗粒团聚效应将主导吸湿行为。'
))

# --- 7.5 小结 ---
new_ch7.append(make_heading('7.5 本章小结'))
new_ch7.append(make_para(
    '本章对不同老化条件下竹粉/PLA/纳米SiO₂复合材料的力学性能和吸湿性能进行了系统对比分析，主要结论如下：'
))
new_ch7.append(make_para(
    '（1）三种老化条件对复合材料力学性能的影响程度为：湿热老化 > 耦合老化 > 紫外老化。'
    '湿热老化是导致复合材料性能下降的主要因素，'
    '而紫外老化对性能的影响相对温和，主要局限于材料表层。'
))
new_ch7.append(make_para(
    '（2）纳米SiO₂的添加在不同老化条件下表现出差异化的防护效果。'
    '对于拉伸强度，1%SiO₂组在耦合老化中的保留率最高（93.4%），'
    '但5%SiO₂组因粒子团聚反而降至83.2%；对于弯曲强度，'
    '5%SiO₂组在耦合老化后保留率超过100%，保护效果最为显著。'
    '综合考虑各项性能，1%SiO₂在大多数条件下表现最优。'
))
new_ch7.append(make_para(
    '（3）5%SiO₂组的弯曲模量在三种老化条件下均呈上升趋势'
    '（紫外18.4%、湿热10.2%、耦合24.0%），'
    '归因于纳米SiO₂促进PLA二次结晶和紫外光诱导交联的"补偿效应"。'
    '吸湿性能方面，3%SiO₂组的吸湿率最高（4.10%），'
    '而5%SiO₂组因颗粒团聚形成致密阻隔结构，吸湿率降至1.50%。'
))

# 插入新第七章
insert_elements_before(body, ch8_new_idx, new_ch7)
print(f"新第七章已插入: {len(new_ch7)} 个元素")

# ================================================================
# PHASE 8: 保存
# ================================================================
doc.save(doc_path)
print(f"\n文档已保存: {doc_path}")

# ================================================================
# PHASE 9: 验证
# ================================================================
print("\n=== 验证 ===")
doc2 = Document(doc_path)
body2 = doc2.element.body

# 检查章节标题
ch_titles = []
for i, child in enumerate(list(body2)):
    if etree.QName(child.tag).localname == 'p':
        text = get_text(child)
        if re.match(r'^[第]\S+[章]', text) or re.match(r'^[0-9]+\.\d', text):
            ch_titles.append(f"body[{i}]: {text[:60]}")
        # 也检查节标题
        if text.startswith('3.4') or text.startswith('3.5') or \
           text.startswith('4.4') or text.startswith('4.5') or \
           text.startswith('5.4') or text.startswith('5.5') or \
           text.startswith('6.4') or text.startswith('6.5') or \
           text.startswith('7.'):
            ch_titles.append(f"body[{i}]: {text[:60]}")

print("关键标题:")
for t in ch_titles:
    print(f"  {t}")

# 检查FTIR图编号
print("\nFTIR图编号检查:")
for i, child in enumerate(list(body2)):
    if etree.QName(child.tag).localname == 'p':
        text = get_text(child)
        if '图3-7' in text or '图4-9' in text or '图5-9' in text or '图6-9' in text:
            print(f"  body[{i}]: {text[:60]}")
        if '图7-5' in text and '弯曲' in text:
            print(f"  body[{i}]: {text[:60]}")
        if '图7-6' in text and '冲击' in text:
            print(f"  body[{i}]: {text[:60]}")
        if '图7-7' in text and '吸湿' in text:
            print(f"  body[{i}]: {text[:60]}")

print("\n重构完成!")
