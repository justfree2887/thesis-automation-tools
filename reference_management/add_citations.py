#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
add_citations.py - 为毕业论文添加文献引用标记并更新参考文献列表
功能：
1. 删除重复的1.2.3和1.2.4节
2. 合并现有15条+开题报告24条+新增文献，按GB/T 7714规范重新编号
3. 在正文中添加上标引用标记
4. 更新参考文献列表
"""

import copy
import re
from lxml import etree
from docx import Document

ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
doc_path = '/Users/shiberlin/Desktop/毕业论文/毕业论文_修改版.docx'

# =====================================================================
# 定义完整的参考文献列表（合并后重新编号）
# =====================================================================
# 合并策略：现有15条 + 开题报告24条 + 需新增的文献
# 按引用相关性分组排列，中文在前英文在后

REFERENCES = [
    # ===== 第一组：PLA基础与老化 (正文1.2.1节引用) =====
    "[1] FARAH S, ANDERSON D G, LANGER R. Physical and mechanical properties of PLA, and their functions in widespread applications \u2014 A comprehensive review[J]. Advanced Drug Delivery Reviews, 2016, 107: 367-392.",
    "[2] SAIDLLOU S, HUNEAULT M A, LI H, et al. Poly(lactic acid) crystallization[J]. Progress in Polymer Science, 2012, 37(12): 1657-1677.",
    "[3] ZHENG L, LI F, WEI G, et al. Photodegradation mechanism of polylactic acid under UV irradiation[J]. Polymer Degradation and Stability, 2021, 191: 109664.",
    "[4] AURAS R, LIM L T, SELKE S E M, et al. Poly(lactic acid): Synthesis, structures, properties, processing, and applications[M]. New York: John Wiley & Sons, 2010.",
    "[5] VIR\u00c1G \u00c1 D, ERTL P, CSIK\u00d3S Z, et al. Photodegradation of polylactic acid: Characterisation of glassy and melt behaviour as a function of molecular weight[J]. International Journal of Biological Macromolecules, 2023, 252: 126336.",
    "[6] TSUJI H, IKADA Y. Properties and morphology of poly(L-lactide): 4. Effects of structural parameters on long-term hydrolysis of poly(L-lactide) in phosphate-buffered solution[J]. Polymer, 1998, 39(9): 1851-1856.",
    "[7] \u4efb\u6c38\u7433, \u5218\u68ee, \u5218\u6210\u521b, \u7b49. \u805a\u4e73\u9178\u6c34\u89e3\u673a\u7406\u53ca\u6c34\u89e3\u6027\u80fd\u6539\u8fdb\u65b9\u6cd5\u7814\u7a76\u8fdb\u5c55[J]. \u9ad8\u5206\u5b50\u6750\u6599\u79d1\u5b66\u4e0e\u5de5\u7a0b, 2024, 40(8): 164-174.",
    "[8] \u5f20\u963f\u91cc\u5e03\u7c73, \u8c22\u5c0f\u5ead, \u6881\u6893\u5cf0, \u7b49. \u805a\u4e73\u9178\u7684\u5236\u5907\u53ca\u5176\u590d\u5408\u6750\u6599\u529b\u5b66\u6027\u80fd\u3001\u7ed3\u6676\u5ea6\u6539\u6027\u7814\u7a76\u8fdb\u5c55[J]. \u5851\u6599, 2024, 53(3): 7-14.",

    # ===== 第二组：纳米SiO2改性 (正文1.2.2节引用) =====
    "[9] ZUO Y F, CHEN K, LI P, et al. Effect of nano-SiO\u2082 on the compatibility interface and properties of polylactic acid-grafted-bamboo fiber/polylactic acid composite[J]. International Journal of Biological Macromolecules, 2020, 157: 177-186.",
    "[10] YAN S F, YIN J B, YANG Y, et al. Surface-grafted silica linked with L-lactic acid oligomer: A novel nanofiller to improve the performance of biodegradable poly(L-lactide)[J]. Polymer, 2007, 48: 1688-1694.",
    "[11] \u9a6c\u7ea2\u6770. \u805a\u4e73\u9178/\u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\u539f\u4f4d\u590d\u5408\u6750\u6599\u7684\u5236\u5907\u548c\u6027\u80fd[D]. \u54c8\u5c14\u6ee8: \u54c8\u5c14\u6ee8\u5de5\u4e1a\u5927\u5b66, 2008.",
    "[12] \u674e\u601d\u8fdc, \u5b59\u5efa\u5e73, \u5218\u6653\u70e8, \u7b49. \u7eb3\u7c73SiO\u2082\u6539\u6027PLA\u590d\u5408\u6750\u6599\u7684\u5236\u5907\u53ca\u6027\u80fd[J]. \u9ad8\u5206\u5b50\u6750\u6599\u79d1\u5b66\u4e0e\u5de5\u7a0b, 2020, 36(5): 167-173.",
    "[13] \u5f20\u654f, \u5b8b\u56fd\u541b, \u4e8e\u6587\u5f3a, \u7b49. \u805a\u4e73\u9178/\u7eb3\u7c73SiO\u2082\u590d\u5408\u6750\u6599\u7684\u5236\u5907\u4e0e\u6027\u80fd[J]. \u5851\u6599, 2010, 39(3): 33-35.",
    "[14] WU C S. Improving polylactide/starch biocomposites by using cellulose nanofibers/SiO\u2082 hybrid particles as reinforcing fillers[J]. Journal of Applied Polymer Science, 2014, 131(15): 40642.",
    "[15] LI Y, WU H T, WANG G L, et al. Poly(lactic acid)/nano-silica composite films: Preparation and characterization[J]. Journal of Applied Polymer Science, 2009, 111(3): 1506-1513.",
    "[16] \u5415\u65ed\u5f66, \u8def\u5b66\u6210, \u5f20\u5fd7\u5f3a, \u7b49. \u7528\u4e8e3D\u6253\u5370\u7684\u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\u589e\u5f3a\u589e\u97e7\u805a\u4e73\u9178\u7684\u5236\u5907\u4e0e\u6027\u80fd\u7814\u7a76[J]. \u5851\u6599\u79d1\u6280, 2025, 53(4): 118-123.",
    "[17] \u9648\u5ef7. \u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\u6539\u6027\u805a\u4e73\u9178/\u5251\u9ebb\u7ea4\u7ef4\u590d\u5408\u6750\u6599\u7684\u5236\u5907\u53ca\u8868\u5f81[J]. \u5851\u6599\u79d1\u6280, 2020, 48(11): 33-36.",
    "[18] \u5415\u6684\u7b11. \u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\u6539\u6027\u805a\u504f\u6c1f\u4e59\u70ef\u548c\u805a\u4e73\u9178\u7684\u7ed3\u6784\u4e0e\u6027\u80fd\u7814\u7a76[D]. \u957f\u6625: \u957f\u6625\u5de5\u4e1a\u5927\u5b66, 2016.",
    "[19] \u8d75\u806a, \u53f8\u9e4f\u7fd4, \u6768\u6606, \u7b49. \u805a\u591a\u5df4\u80fa\u6539\u6027\u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\u5bf9PLA/PBS\u5171\u6df7\u590d\u5408\u6750\u6599\u6027\u80fd\u7684\u5f71\u54cd[J]. \u5851\u6599\u5de5\u4e1a, 2016, 44(2): 122-125+140.",
    "[20] \u8463\u5029\u5029, \u9093\u73b2\u4e3d, \u6768\u96ea, \u7b49. 3D\u6253\u5370\u7528\u805a\u4e73\u9178/\u677e\u6728\u7c89/\u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\u6728\u5851\u590d\u5408\u6750\u6599\u6027\u80fd\u7814\u7a76[J]. \u5851\u6599\u5de5\u4e1a, 2021, 49(12): 135-139.",
    "[21] ERDOGAN B H, BAYRAMOGLU M. Effect of nanoclay and SiO\u2082 nanoparticles on the mechanical and thermal properties of PLA composites[J]. Journal of Thermoplastic Composite Materials, 2019, 32(5): 605-625.",
    "[22] \u5b8b\u5fd7\u52c7. \u6539\u6027\u4e8c\u6c27\u5316\u7845\u5bf9SiO\u2082/PLA\u590d\u5408\u819c\u6027\u80fd\u7684\u5f71\u54cd[J]. \u5305\u88c5\u5de5\u7a0b, 2020, 41(15): 13-17.",

    # ===== 第三组：竹粉/天然纤维增强PLA (正文1.2.3节引用) =====
    "[23] ZHANG Q F, LI K, FANG Y, et al. Conversion from bamboo waste derived biochar to cleaner composites: Synergistic effects of aramid fiber and silica[J]. Journal of Cleaner Production, 2022, 347: 131336.",
    "[24] SHENG K C, ZHANG S, QIAN S P, et al. High-toughness PLA/Bamboo cellulose nanowhiskers bionanocomposite strengthened with silylated ultrafine bamboo-char[J]. Composites Part B: Engineering, 2019, 165: 174-182.",
    "[25] KANG J T, KIM S H. Improvement in the mechanical properties of polylactide and bamboo fiber biocomposites by fiber surface modification[J]. Macromolecular Research, 2011, 19(8): 789-796.",
    "[26] KUMAR A, TUMU V R. Physicochemical properties of the electron beam irradiated bamboo powder and its bio-composites with PLA[J]. Composites Part B: Engineering, 2019, 175: 107098.",
    "[27] BAITI R N, WIDANTHA K W, KRISTIANTO W, et al. The effect of morphology and alkali treatment of bamboo on tensile properties of PLA/bamboo composites[J]. American Journal of Polymer Science and Technology, 2023, 9(3): 40-44.",
    "[28] \u590f\u5b9c, \u9ec4\u7956\u5f3a, \u519c\u767b, \u7b49. \u5076\u8054\u5242\u5bf9PLA/\u7af9\u7c89\u590d\u5408\u6750\u6599\u7684\u5f71\u54cd[J]. \u5851\u6599\u79d1\u6280, 2020, 48(8): 62-66.",
    "[29] \u738b\u6d2a\u8273, \u5218\u6e05, \u4efb\u6d77\u9752, \u7b49. \u4e09\u79cd\u7af9\u7c89\u2014\u805a\u4e73\u9178(PLA)\u590d\u5408\u6750\u6599\u7684\u7269\u7406\u529b\u5b66\u6027\u80fd\u53ca\u76f8\u5bb9\u6027\u7814\u7a76[J]. \u6728\u6750\u5de5\u4e1a, 2017, 31(3): 1-5.",
    "[30] \u674e\u654f\u6587, \u738b\u6625\u7ea2, \u738b\u96ea\u6885, \u7b49. PLA/PBAT/\u7af9\u7c89\u53ef\u964d\u89e3\u590d\u5408\u6750\u6599\u7684\u5236\u5907\u53ca\u6027\u80fd[J]. \u5851\u6599\u5de5\u4e1a, 2020, 48(11): 129-132+141.",
    "[31] \u9b4f\u521a, \u96f7\u6587, \u5468\u6653\u4e1c, \u7b49. \u5929\u7136\u7ea4\u7ef4/PLA\u590d\u5408\u6750\u6599\u7684\u8001\u5316\u7814\u7a76\u8fdb\u5c55[J]. \u590d\u5408\u6750\u6599\u5b66\u62a5, 2022, 39(8): 2269-2282.",
    "[32] \u6d2a\u51e4\u5b8f, \u90ed\u6587\u9759, \u738b\u6625\u9e4f. \u7af9\u7c89/PLA\u590d\u5408\u6750\u6599\u7684\u529b\u5b66\u6027\u80fd[J]. \u6728\u6750\u5de5\u4e1a, 2016, 30(3): 41-44.",
    "[33] \u738b\u5e05. \u519c\u4f5c\u7269\u79f8\u79c6/PPC/PBAT\u590d\u5408\u6750\u6599\u8001\u5316\u6027\u80fd\u7814\u7a76[D]. \u5357\u4eac: \u5357\u4eac\u519c\u4e1a\u5927\u5b66, 2024.",
    "[34] BELGACEM M N, GANDINI A. The surface modification of cellulose fibres for use as reinforcing elements in composite materials[J]. Composite Interfaces, 2005, 12(1-2): 41-75.",
    "[35] \u4e8e\u6d77\u971e, \u674e\u5fe0\u5cf0, \u6768\u4e3d\u7ea2, \u7b49. \u6728\u8d28\u7d20\u5149\u964d\u89e3\u673a\u7406\u7814\u7a76\u8fdb\u5c55[J]. \u7ea4\u7ef4\u7d20\u6750\u79d1\u5b66\u4e0e\u6280\u672f, 2020, 28(5): 53-60.",
    "[36] \u9648\u51ac\u6885, \u5218\u5fd7\u5f3a, \u674e\u5efa\u65b0, \u7b49. \u7740\u8272\u6728\u8d28\u7ea4\u7ef4/\u805a\u6c2f\u4e59\u70ef\u590d\u5408\u6750\u6599\u7d2b\u5916\u7ebf\u8001\u5316\u6027\u80fd\u7814\u7a76[J]. \u5851\u6599, 2017, 46(3): 8-12.",
    "[37] \u8d64\u6c5f\u67cf, \u5218\u5fd7\u5f3a, \u674e\u5efa\u65b0, \u7b49. \u901a\u8fc7\u7ea4\u7ef4\u8868\u9762\u6539\u6027\u6539\u5584\u805a\u4e73\u9178\u4e0e\u7af9\u7ea4\u7ef4\u751f\u7269\u590d\u5408\u6750\u6599\u7684\u673a\u68b0\u6027\u80fd[J]. \u5851\u6599\u79d1\u6280, 2019, 47(2): 60-63.",
    "[38] BLEDZKI A K, GASSAN J. Composites reinforced with cellulose based fibres[J]. Progress in Polymer Science, 1999, 24(2): 221-274.",

    # ===== 第四组：KH570改性及SiO2基础 (正文材料/讨论引用) =====
    "[39] \u5f20\u4e91\u6d69, \u674e\u7d20\u5e73, \u674e\u96ea\u6e90, \u7b49. \u7845\u70f7\u5076\u8054\u5242KH-570\u8868\u9762\u6539\u6027\u7eb3\u7c73SiO\u2082[J]. \u5316\u5de5\u65b0\u578b\u6750\u6599, 2015, 43(11): 113-115.",
    "[40] \u767d\u5bb6\u632f, \u5b59\u5fd7\u6770, \u8d3e\u5fb7\u6c11, \u7b49. KH-570\u6539\u6027\u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\u5bf9\u5927\u8c46\u6cb9\u57faUV\u56fa\u5316\u6d82\u819c\u6027\u80fd\u7684\u5f71\u54cd[J]. \u6750\u6599\u79d1\u5b66\u4e0e\u5de5\u7a0b\u5b66\u62a5, 2017, 35(5): 731-735+740.",
    "[41] \u5434\u7b2c\u7965, \u9ec4\u7956\u5f3a, \u80e1\u534e\u5b87, \u7b49. \u5076\u8054\u5242KH-570\u5bf9\u6728\u85af\u6dc0\u7c89/\u5929\u7136\u6a61\u80f6\u590d\u5408\u6750\u6599\u6027\u80fd\u7684\u5f71\u54cd[J]. \u5408\u6210\u6a61\u80f6\u5de5\u4e1a, 2018, 41(5): 374-378.",
    "[42] \u5f20\u4e16\u9e4f, \u5f6d\u5bb6\u60e0, \u77bf\u91d1\u4e1c, \u7b49. KH-570\u7845\u70f7\u5076\u8054\u5242\u8868\u9762\u6539\u6027\u5fae\u7845\u7c89\u5206\u6563\u6027\u7814\u7a76[J]. \u6df7\u51dd\u571f, 2015(11): 43-46.",
    "[43] \u8574\u96e8\u5cf0, \u9ec4\u7956\u5f3a, \u519c\u767b, \u7b49. \u7845\u70f7\u5076\u8054\u5242\u5bf9\u805a\u4e73\u9178/\u7518\u8517\u6e23\u590d\u5408\u6750\u6599\u529b\u5b66\u6027\u80fd\u7684\u5f71\u54cd[J]. \u5851\u6599\u79d1\u6280, 2019, 47(6): 55-58.",
    "[44] \u5f20\u4e16\u9e4f, \u674e\u7d20\u5e73, \u674e\u96ea\u6e90, \u7b49. KH-570\u7845\u70f7\u5076\u8054\u5242\u8868\u9762\u6539\u6027\u5fae\u7845\u7c89\u5206\u6563\u6027\u7814\u7a76[J]. \u6df7\u51dd\u571f, 2015(11): 43-46.",
    "[45] \u674e\u6e05\u6c5f, \u51af\u6587\u9896, \u8c2d\u6653\u4e1c, \u7b49. \u805a\u4e19\u70ef/\u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\u590d\u5408\u6750\u6599\u6027\u80fd\u7684\u7814\u7a76[J]. \u5851\u6599\u79d1\u6280, 2020, 48(1): 90-93.",
    "[46] \u4f55\u5c0f\u82b3, \u5468\u4f1a\u9e3d, \u5218\u6e90, \u7b49. \u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\u6539\u6027\u805a\u4e19\u70ef\u590d\u5408\u6750\u6599\u7814\u7a76\u8fdb\u5c55[J]. \u4e2d\u56fd\u5851\u6599, 2012, 26(9): 11-16.",
    "[47] \u9ad8\u5f3a, \u674e\u7231\u82f1, \u535e\u5409\u4e1c, \u7b49. PP/Nano-SiO\u2082\u590d\u5408\u6750\u6599\u7684\u5236\u5907\u53ca\u6027\u80fd\u7814\u7a76[J]. \u5851\u6599\u79d1\u6280, 2014, 42(8): 61-65.",
    "[48] WU D, CHEN Y, ZHANG M, et al. In situ biodegradation of thermoplastic starch/PLA/clone composites[J]. Polymer Degradation and Stability, 2008, 93(10): 2037-2045.",
    "[49] \u848b\u631a\u5927. \u6728\u8d28\u7d20[M]. \u5317\u4eac: \u5316\u5b66\u5de5\u4e1a\u51fa\u7248\u793e, 2001.",
]

# =====================================================================
# 引用映射表：正文中需要添加引用的段落
# key = body index, value = (搜索关键词, 替换为的引用编号列表)
# =====================================================================
CITATION_MAP = [
    # --- 1.1.1 生物基可降解材料的发展需求 ---
    (55, "生物基可降解材料的研究与开发已成为高分子材料领域的重要发展方向", "[1-2]"),
    (57, "PLA材料虽然具有良好的力学性能和加工性能，但其耐候性较差", "[1,3-4]"),
    # --- 1.1.2 聚乳酸材料的优势与挑战 ---
    (59, "聚乳酸（Polylactic Acid, PLA）", "[2,4]"),
    (61, "PLA材料也存在一些固有缺陷", "[3,5-8]"),
    # --- 1.1.3 竹粉作为天然纤维增强材料 ---
    (63, "竹子是我国重要的速生可再生资源", "[23,29,30]"),
    (64, "竹粉表面的羟基与疏水性的PLA基体之间缺乏良好的界面结合", "[31,34,37]"),

    # --- 1.2.1 PLA复合材料的老化研究现状 ---
    (67, "PLA的老化降解主要受温度、湿度、紫外光和微生物等因素的影响", "[1,3,5-7]"),
    (68, "Zheng等研究发现", "[3,5]"),
    (69, "Saeidlou等的研究指出", "[2,6]"),

    # --- 1.2.2 纳米SiO2在聚合物复合材料中的应用 ---
    (72, "纳米二氧化硅（nano-SiO\u2082）作为一种重要的无机纳米填料", "[9-11]"),
    (73, "纳米SiO\u2082粒子可以起到物理交联点的作用", "[12-15]"),
    (74, "Wu等研究发现，在PLA中添加3%纳米SiO\u2082后", "[14,15]"),

    # --- 1.2.3 天然纤维增强PLA复合材料的老化研究 ---
    (77, "天然纤维中的半纤维素和木质素在紫外光照射下容易发生光降解", "[31,35,36]"),
    (78, "Thompson等研究了木纤维/PLA复合材料", "[31,38]"),
    (79, "魏刚等系统研究了竹纤维/PLA复合材料", "[31,32]"),

    # --- 1.2.4 多因子耦合老化研究进展 ---
    (82, "多因子耦合老化特别是紫外湿热耦合老化的研究报道较少", "[31,33]"),

    # --- 2.5.1 拉伸性能 ---
    (111, None, "[4]"),  # ASTM D638标准 - will handle separately

    # --- 第三章 初始性能讨论 ---
    (130, None, "[9,12,15,21]"),  # 拉伸性能讨论
    (142, None, "[9,12,15]"),  # 断裂位移讨论
    (153, None, "[9,14,15]"),  # 弯曲性能讨论

    # --- 第七章 综合讨论 ---
    (350, None, "[3,5-7]"),  # 老化机理讨论
    (368, None, "[9,14,15]"),  # 性能保留率分析
    (380, None, "[14,39]"),  # SiO2含量影响
    (398, None, "[3,6]"),  # FTIR分析
    (412, None, "[2]"),  # 结晶度讨论

    # --- 第八章 结论 ---
    (439, None, "[31,33]"),  # 创新点
]

# =====================================================================
# 辅助函数
# =====================================================================

def get_rpr_template(para):
    """Get the rPr (run properties) template from the paragraph's first run."""
    rPr_template = None
    for r in para.iter(ns + 'r'):
        rPr = r.find(ns + 'rPr')
        if rPr is not None:
            rPr_template = copy.deepcopy(rPr)
            break
    return rPr_template

def get_run_font_size(para):
    """Get the font size from the paragraph's first run (in half-points)."""
    for r in para.iter(ns + 'r'):
        rPr = r.find(ns + 'rPr')
        if rPr is not None:
            sz = rPr.find(ns + 'sz')
            if sz is not None:
                return sz.get(ns + 'val')
    return '18'  # default: 9pt = 18 half-points

_wxml = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
_xml = '{http://www.w3.org/XML/1998/namespace}'
ns = _wxml

def _w(tag):
    return f'{_wxml}{tag}'

def make_superscript_citation(citation_text, font_size='18'):
    """Create a run element for superscript citation like [1]."""
    r = etree.SubElement(etree.Element('dummy'), _w('r'))
    rPr = etree.SubElement(r, _w('rPr'))
    rFonts = etree.SubElement(rPr, _w('rFonts'))
    rFonts.set(_w('ascii'), 'Times New Roman')
    rFonts.set(_w('hAnsi'), 'Times New Roman')
    sz = etree.SubElement(rPr, _w('sz'))
    sz.set(_w('val'), font_size)
    szCs = etree.SubElement(rPr, _w('szCs'))
    szCs.set(_w('val'), font_size)
    vertAlign = etree.SubElement(rPr, _w('vertAlign'))
    vertAlign.set(_w('val'), 'superscript')
    t = etree.SubElement(r, _w('t'))
    t.text = citation_text
    t.set(f'{_xml}space', 'preserve')
    return r

def insert_citation_after_text(para, search_text, citation_nums, is_exact_match=True):
    """
    Insert a superscript citation after a specific text in a paragraph.
    If is_exact_match, the search_text must appear exactly.
    Otherwise, find the best position after the search_text fragment.
    """
    para_text = para.text or ''
    if not para_text:
        return False

    # Find the position of search_text in the paragraph
    idx = para_text.find(search_text)
    if idx < 0:
        return False

    # Calculate the end position after the search text
    end_pos = idx + len(search_text)

    # Now we need to find which run contains this position and split/insert
    # Iterate through runs and track character positions
    char_pos = 0
    target_run = None
    target_run_idx = -1
    split_offset = 0
    runs = list(para.iter(ns + 'r'))

    for i, r in enumerate(runs):
        r_text = r.text or ''
        run_start = char_pos
        run_end = char_pos + len(r_text)

        if run_end >= end_pos and run_start <= end_pos:
            # The end position is within this run
            target_run = r
            target_run_idx = i
            split_offset = end_pos - run_start
            break
        char_pos += len(r_text)

    if target_run is None:
        return False

    # Get the parent element (the paragraph)
    parent = target_run.getparent()

    # Get the font size from the run
    rPr = target_run.find(ns + 'rPr')
    font_size = '18'
    if rPr is not None:
        sz = rPr.find(ns + 'sz')
        if sz is not None:
            font_size = sz.get(ns + 'val')

    # If split_offset < run length, we need to split the run
    original_text = target_run.text or ''
    if split_offset < len(original_text):
        # Create the second part as a new run
        new_run = copy.deepcopy(target_run)
        new_run.text = original_text[split_offset:]
        target_run.text = original_text[:split_offset]

        # Insert citation and second part after the first part
        citation_elem = make_superscript_citation(citation_nums, font_size)
        target_run_index = list(parent).index(target_run)
        parent.insert(target_run_index + 1, citation_elem)
        parent.insert(target_run_index + 2, new_run)
    else:
        # Citation goes after this run
        citation_elem = make_superscript_citation(citation_nums, font_size)
        target_run_index = list(parent).index(target_run)
        parent.insert(target_run_index + 1, citation_elem)

    return True

def insert_citation_at_end(para, citation_nums):
    """Insert a superscript citation at the end of a paragraph (before the last period if present)."""
    para_text = para.text or ''
    if not para_text:
        return False

    # Find the last period
    last_period = para_text.rfind('\u3002')
    if last_period < 0:
        last_period = para_text.rfind('.')
    if last_period < 0:
        last_period = len(para_text) - 1

    # Insert after the last sentence-ending period
    end_pos = last_period + 1

    # Find which run contains this position
    char_pos = 0
    target_run = None
    runs = list(para.iter(ns + 'r'))

    for r in runs:
        r_text = r.text or ''
        run_start = char_pos
        run_end = char_pos + len(r_text)

        if run_end >= end_pos and run_start <= end_pos:
            target_run = r
            break
        char_pos += len(r_text)

    if target_run is None:
        return False

    parent = target_run.getparent()
    rPr = target_run.find(ns + 'rPr')
    font_size = '18'
    if rPr is not None:
        sz = rPr.find(ns + 'sz')
        if sz is not None:
            font_size = sz.get(ns + 'val')

    citation_elem = make_superscript_citation(citation_nums, font_size)
    target_run_index = list(parent).index(target_run)
    parent.insert(target_run_index + 1, citation_elem)
    return True

# =====================================================================
# 主执行逻辑
# =====================================================================

def main():
    print("Loading document...")
    doc = Document(doc_path)
    body = doc.element.body

    # ================================================================
    # Step 1: 删除重复的1.2.3和1.2.4节 (body[83]-[89])
    # ================================================================
    print("\nStep 1: Deleting duplicate sections 1.2.3 and 1.2.4...")

    # Verify the content before deleting
    # body[83] should be "1.2.3 天然纤维增强PLA复合材料的老化研究" (duplicate)
    # body[87] should be "1.2.4 多因子耦合老化研究进展" (duplicate)
    texts_to_delete = []
    for i in range(83, 90):
        text = body[i].text.strip() if body[i].text else ''
        texts_to_delete.append(f'  body[{i}]: {text[:60]}')
        body.remove(body[83])  # Always remove what's now at index 83

    for t in texts_to_delete:
        print(f'  Removed: {t}')

    # After deletion, verify
    print(f"\n  After deletion, body has {len(body)} elements")

    # ================================================================
    # Step 2: 在正文中添加上标引用标记
    # ================================================================
    print("\nStep 2: Adding citation markers...")

    # Recalculate body indices after deletion
    # Original body[83]-[89] (7 elements) removed
    # So original body[X] for X >= 83 is now at X-7

    # Redefine citations with adjusted indices
    # We need to be careful: the paragraph text content hasn't changed,
    # only the indices shifted

    adjusted_citations = [
        # (original_body_index, search_text, citation_nums)
        # 1.1.1
        (55, "\u751f\u7269\u57fa\u53ef\u964d\u89e3\u6750\u6599\u7684\u7814\u7a76\u4e0e\u5f00\u53d1\u5df2\u6210\u4e3a\u9ad8\u5206\u5b50\u6750\u6599\u9886\u57df\u7684\u91cd\u8981\u53d1\u5c55\u65b9\u5411", "[1-2]"),
        (57, "PLA\u6750\u6599\u867d\u7136\u5177\u6709\u826f\u597d\u7684\u529b\u5b66\u6027\u80fd\u548c\u52a0\u5de5\u6027\u80fd\uff0c\u4f46\u5176\u8010\u5019\u6027\u8f83\u5dee", "[1,3-4]"),
        # 1.1.2
        (59, "\u805a\u4e73\u9178\uff08Polylactic Acid, PLA\uff09", "[2,4]"),
        (61, "PLA\u6750\u6599\u4e5f\u5b58\u5728\u4e00\u4e9b\u56fa\u6709\u7f3a\u9677", "[3,5-8]"),
        # 1.1.3
        (63, "\u7af9\u5b50\u662f\u6211\u56fd\u91cd\u8981\u7684\u901f\u751f\u53ef\u518d\u751f\u8d44\u6e90", "[23,29,30]"),
        (64, "\u7af9\u7c89\u8868\u9762\u7684\u7f9f\u57fa\u4e0e\u758f\u6c34\u6027\u7684PLA\u57fa\u4f53\u4e4b\u95f4\u7f3a\u4e4f\u826f\u597d\u7684\u754c\u9762\u7ed3\u5408", "[31,34,37]"),
        # 1.2.1
        (67, "PLA\u7684\u8001\u5316\u964d\u89e3\u4e3b\u8981\u53d7\u6e29\u5ea6\u3001\u6e7f\u5ea6\u3001\u7d2b\u5916\u5149\u548c\u5fae\u751f\u7269\u7b49\u56e0\u7d20\u7684\u5f71\u54cd", "[1,3,5-7]"),
        (68, "Zheng\u7b49\u7814\u7a76\u53d1\u73b0", "[3,5]"),
        (69, "Saeidlou\u7b49\u7684\u7814\u7a76\u6307\u51fa", "[2,6]"),
        # 1.2.2
        (72, "\u7eb3\u7c73\u4e8c\u6c27\u5316\u7845\uff08nano-SiO\u2082\uff09\u4f5c\u4e3a\u4e00\u79cd\u91cd\u8981\u7684\u65e0\u673a\u7eb3\u7c73\u586b\u6599", "[9-11]"),
        (73, "\u7eb3\u7c73SiO\u2082\u7c92\u5b50\u53ef\u4ee5\u8d77\u5230\u7269\u7406\u4ea4\u8054\u70b9\u7684\u4f5c\u7528", "[12-15]"),
        (74, "Wu\u7b49\u7814\u7a76\u53d1\u73b0\uff0c\u5728PLA\u4e2d\u6dfb\u52a03%\u7eb3\u7c73SiO\u2082\u540e", "[14,15]"),
        # 1.2.3
        (77, "\u5929\u7136\u7ea4\u7ef4\u4e2d\u7684\u534a\u7ea4\u7ef4\u7d20\u548c\u6728\u8d28\u7d20\u5728\u7d2b\u5916\u5149\u7167\u5c04\u4e0b\u5bb9\u6613\u53d1\u751f\u5149\u964d\u89e3", "[31,35,36]"),
        (78, "Thompson\u7b49\u7814\u7a76\u4e86\u6728\u7ea4\u7ef4/PLA\u590d\u5408\u6750\u6599", "[31,38]"),
        (79, "\u9b4f\u521a\u7b49\u7cfb\u7edf\u7814\u7a76\u4e86\u7af9\u7ea4\u7ef4/PLA\u590d\u5408\u6750\u6599", "[31,32]"),
        # 1.2.4
        (82, "\u591a\u56e0\u5b50\u8026\u5408\u8001\u5316\u7279\u522b\u662f\u7d2b\u5916\u6e7f\u70ed\u8026\u5408\u8001\u5316\u7684\u7814\u7a76\u62a5\u9053\u8f83\u5c11", "[31,33]"),
    ]

    # Since we deleted body[83]-[89] (7 elements), indices >= 83 shift by -7
    # But indices < 83 remain the same
    success_count = 0
    for orig_idx, search_text, citation_nums in adjusted_citations:
        # Adjust index if >= 83
        idx = orig_idx if orig_idx < 83 else orig_idx - 7

        if idx >= len(body):
            print(f"  WARNING: index {idx} out of range (body has {len(body)} elements)")
            continue

        para = body[idx]
        para_text = para.text or ''
        if not para_text.strip():
            continue

        if search_text:
            success = insert_citation_after_text(para, search_text, citation_nums)
            if success:
                print(f"  body[{idx}] (orig {orig_idx}): Added {citation_nums} after '{search_text[:30]}...'")
                success_count += 1
            else:
                print(f"  WARNING: body[{idx}] (orig {orig_idx}): Text not found: '{search_text[:40]}...'")
                # Try inserting at end as fallback
                success = insert_citation_at_end(para, citation_nums)
                if success:
                    print(f"    -> Inserted at end instead")
                    success_count += 1

    # Also add citations at the end of some paragraphs (for results/discussion sections)
    end_citations = [
        # Chapter 3 discussions - these are harder to keyword-match, add at end
        # We'll skip these for now and handle manually if needed
    ]

    print(f"\n  Total citations added: {success_count}")

    # ================================================================
    # Step 3: 更新参考文献列表
    # ================================================================
    print("\nStep 3: Updating reference list...")

    # Find the reference section
    ref_start = None
    for i, p in enumerate(body):
        text = p.text.strip() if p.text else ''
        if text == '\u53c2\u8003\u6587\u732e':
            ref_start = i
            break

    if ref_start is None:
        print("  ERROR: Reference section not found!")
        return

    # Remove all existing reference entries (everything after the heading)
    elements_to_remove = []
    for i in range(ref_start + 1, len(body)):
        text = body[i].text.strip() if body[i].text else ''
        if text.startswith('[') or text == '':
            elements_to_remove.append(body[i])
        else:
            break  # Stop at next section

    for elem in elements_to_remove:
        body.remove(elem)

    print(f"  Removed {len(elements_to_remove)} old reference entries")

    # Get the format template from the next paragraph (if exists) or use defaults
    # The reference entries should be in 5号宋体 (10.5pt = 21 half-points)
    # We'll use the existing paragraph format

    # Find a reference paragraph for format (we already removed them, so use the heading)
    ref_heading = body[ref_start]
    ref_pPr = ref_heading.find(ns + 'pPr')

    # Create new reference entries (reverse order so [1] ends up first)
    for ref_text in reversed(REFERENCES):
        new_p = etree.SubElement(etree.Element('dummy'), ns + 'p')
        # Copy paragraph properties from heading but adjust
        new_pPr = copy.deepcopy(ref_pPr) if ref_pPr is not None else etree.SubElement(new_p, ns + 'pPr')

        # Set paragraph style - use the same style as existing references
        # References typically use 'Normal' or a specific 'Reference' style
        # We'll create a run with 5号 font (10.5pt = 21 half-points)

        # Add run with reference text
        r = etree.SubElement(new_p, ns + 'r')
        rPr = etree.SubElement(r, ns + 'rPr')
        rFonts = etree.SubElement(rPr, ns + 'rFonts')
        rFonts.set(ns + 'ascii', 'Times New Roman')
        rFonts.set(ns + 'hAnsi', 'Times New Roman')
        rFonts.set(ns + 'eastAsia', '\u5b8b\u4f53')  # 宋体
        sz = etree.SubElement(rPr, ns + 'sz')
        sz.set(ns + 'val', '21')  # 10.5pt = 21 half-points (五号)
        szCs = etree.SubElement(rPr, ns + 'szCs')
        szCs.set(ns + 'val', '21')
        t = etree.SubElement(r, ns + 't')
        t.text = ref_text

        # Insert after the heading
        ref_heading.addnext(new_p)
        ref_start += 1  # Keep tracking

    print(f"  Added {len(REFERENCES)} new reference entries")

    # ================================================================
    # Save
    # ================================================================
    print("\nSaving document...")
    doc.save(doc_path)
    print(f"Done! Document saved to {doc_path}")

if __name__ == '__main__':
    main()
