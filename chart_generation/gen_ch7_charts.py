#!/usr/bin/env python3
"""
gen_ch7_charts.py
生成第七章（结论章）的柱状图，使用灰度+hatch黑白打印友好样式。
数据来源：stats_v2.json 或直接从表格数据中硬编码（确保与论文表格一致）。

图表列表：
- 图7-1: 不同老化条件下拉伸强度分组柱状图
- 图7-2: 不同老化条件下弯曲强度分组柱状图
- 图7-3: 不同老化条件下冲击强度分组柱状图
- 图7-4: 老化后力学性能保留率
- 图7-9: 力学性能对比图
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os

# ========== 中文字体 ==========
plt.rcParams['font.sans-serif'] = ['SimHei', 'Heiti SC', 'Arial Unicode MS', 'STHeiti', 'Hiragino Sans GB']
plt.rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 14
plt.rcParams['mathtext.fontset'] = 'custom'
plt.rcParams['mathtext.rm'] = 'Times New Roman'
plt.rcParams['mathtext.it'] = 'Times New Roman:italic'
plt.rcParams['mathtext.bf'] = 'Times New Roman:bold'

SIO2_ORDER = ['0%', '1%', '3%', '5%']

# 第七章有4组数据（vs 第四五章的3组），定义4种样式
# 灰度渐变：0.85 -> 0.65 -> 0.45 -> 0.15
# hatch: 无 -> 斜线 -> 反斜线 -> 点状
GROUP_STYLES = [
    {'color': '0.85',  'hatch': '',       'edgecolor': 'black', 'label': '初始'},
    {'color': '0.60',  'hatch': '///',    'edgecolor': 'black', 'label': '紫外老化'},
    {'color': '0.35',  'hatch': '\\\\\\', 'edgecolor': 'black', 'label': '湿热老化'},
    {'color': '0.15',  'hatch': '...',    'edgecolor': 'black', 'label': '紫外湿热耦合老化'},
]

# 表7-1到7-3的数据（从论文文档提取，与表格完全一致）
# 格式: {老化条件: {SiO2含量: (mean, std)}}
TENSILE_STRENGTH = {  # 表7-1 拉伸强度 (MPa)
    '初始':             {'0%': (62.83, 0.62), '1%': (57.67, 7.53), '3%': (59.50, 6.03), '5%': (65.10, 1.93)},
    '紫外老化':         {'0%': (57.80, 2.09), '1%': (59.84, 2.71), '3%': (58.95, 3.42), '5%': (64.45, 1.83)},
    '湿热老化':         {'0%': (49.23, 1.01), '1%': (51.62, 0.79), '3%': (52.06, 1.96), '5%': (54.89, 1.90)},
    '紫外湿热耦合老化': {'0%': (52.53, 1.87), '1%': (53.86, 3.56), '3%': (52.38, 9.09), '5%': (54.16, 7.50)},
}

FLEXURAL_STRENGTH = {  # 表7-2 弯曲强度 (MPa)
    '初始':             {'0%': (4.623, 0.025), '1%': (4.852, 0.272), '3%': (4.249, 0.395), '5%': (4.557, 0.018)},
    '紫外老化':         {'0%': (4.594, 0.279), '1%': (4.344, 0.276), '3%': (4.453, 0.172), '5%': (4.148, 0.143)},
    '湿热老化':         {'0%': (4.143, 0.254), '1%': (4.413, 0.139), '3%': (3.988, 0.222), '5%': (4.052, 0.142)},
    '紫外湿热耦合老化': {'0%': (4.486, 0.349), '1%': (4.577, 0.491), '3%': (3.839, 0.685), '5%': (3.962, 0.327)},
}

IMPACT_STRENGTH = {  # 表7-3 冲击强度 (kJ/m²)
    '初始':             {'0%': (105.00, 0.73), '1%': (116.10, 5.42), '3%': (112.38, 2.05), '5%': (106.77, 1.42)},
    '紫外老化':         {'0%': (100.36, 3.77), '1%': (104.00, 4.76), '3%': (108.40, 2.72), '5%': (108.73, 1.93)},
    '湿热老化':         {'0%': (86.98, 6.84),  '1%': (96.00, 0.91),  '3%': (97.73, 2.57),  '5%': (99.04, 0.68)},
    '紫外湿热耦合老化': {'0%': (92.80, 1.58),  '1%': (104.90, 1.70), '3%': (107.92, 0.69), '5%': (107.93, 2.34)},
}

# 表7-4 弯曲模量和表7-5 冲击韧性（用于图7-9对比图）
FLEXURAL_MODULUS = {  # 表7-4 弯曲模量 (GPa)
    '初始':             {'0%': (5.056, 0.114), '1%': (5.897, 0.308), '3%': (5.507, 0.171), '5%': (4.764, 0.352)},
    '紫外老化':         {'0%': (5.110, 0.000), '1%': (5.180, 0.474), '3%': (5.608, 0.126), '5%': (5.638, 0.155)},
    '湿热老化':         {'0%': (4.551, 0.474), '1%': (5.041, 0.031), '3%': (5.127, 0.035), '5%': (5.250, 0.072)},
    '紫外湿热耦合老化': {'0%': (4.485, 0.184), '1%': (5.244, 0.072), '3%': (5.819, 0.080), '5%': (5.908, 0.167)},
}

IMPACT_TOUGHNESS = {  # 表7-5 冲击韧性 (kJ/m²)
    '初始':             {'0%': (8.430, 1.105), '1%': (7.822, 1.433), '3%': (7.686, 0.554), '5%': (6.793, 0.510)},
    '紫外老化':         {'0%': (8.109, 1.529), '1%': (8.413, 1.231), '3%': (7.701, 0.425), '5%': (7.480, 0.570)},
    '湿热老化':         {'0%': (9.448, 0.291), '1%': (8.799, 0.674), '3%': (9.350, 0.422), '5%': (7.690, 1.003)},
    '紫外湿热耦合老化': {'0%': (8.539, 0.593), '1%': (8.690, 0.902), '3%': (9.014, 0.815), '5%': (7.254, 0.204)},
}


def make_grouped_bar_4groups(data_dict, output_path, ylabel, y_unit, title):
    """生成4组分组的柱状图（初始、紫外、湿热、耦合）"""
    fig, ax = plt.subplots(figsize=(10, 6))

    n_groups = len(SIO2_ORDER)
    n_bars = len(GROUP_STYLES)
    group_width = 0.70
    bar_width = group_width / n_bars
    x = np.arange(n_groups)

    for i, style in enumerate(GROUP_STYLES):
        label = style['label']
        if label not in data_dict:
            continue
        d = data_dict[label]
        means = [d[sio2][0] for sio2 in SIO2_ORDER]
        stds = [d[sio2][1] for sio2 in SIO2_ORDER]

        offset = (i - (n_bars - 1) / 2) * bar_width
        bars = ax.bar(x + offset, means, bar_width,
                     yerr=stds, capsize=4,
                     facecolor=style['color'],
                     edgecolor=style['edgecolor'],
                     linewidth=1.0,
                     hatch=style['hatch'],
                     label=label,
                     error_kw={'elinewidth': 1.0, 'capthick': 1.0, 'ecolor': 'black'})

        # 计算所有数据点的最大值（全局），用于统一的标签偏移基准
        all_means_global = []
        for sty in GROUP_STYLES:
            lbl = sty['label']
            if lbl in data_dict:
                for s in SIO2_ORDER:
                    all_means_global.append(data_dict[lbl][s][0])
        y_max_global = max(all_means_global) if all_means_global and max(all_means_global) > 0 else 1

        for j, (mean, std) in enumerate(zip(means, stds)):
            # 每组使用交替偏移防止相邻柱的标签重叠
            # 组0(初始):0.01, 组1(紫外):0.07, 组2(湿热):0.01, 组3(耦合):0.07
            label_offsets = [0.01, 0.07, 0.01, 0.07]
            y_pos = mean + std + y_max_global * label_offsets[i % 4]
            ax.text(x[j] + offset, y_pos, f'{mean:.1f}',
                   ha='center', va='bottom', fontsize=13)

    ax.set_xticks(x)
    ax.set_xticklabels([f'SiO\u2082 {k}' for k in SIO2_ORDER], fontsize=14)
    ax.set_ylabel(f'{ylabel} ({y_unit})', fontsize=14)
    ax.set_xlabel('纳米SiO\u2082含量', fontsize=14)
    ax.set_title(title, fontsize=14, fontweight='bold', pad=15)
    ax.legend(fontsize=14, loc='lower left', framealpha=0.9)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # 自动调整y轴
    all_max_vals = []
    for style in GROUP_STYLES:
        label = style['label']
        if label in data_dict:
            for sio2 in SIO2_ORDER:
                all_max_vals.append(data_dict[label][sio2][0] + data_dict[label][sio2][1])
    if all_max_vals:
        # y轴上限需要考虑最高标签位置 + 0.07*y_max_global的偏移
        ax.set_ylim(bottom=0, top=max(all_max_vals) * 1.28)

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close()


def make_retention_chart(data_dict, output_path):
    """生成力学性能保留率柱状图（图7-4）"""
    fig, ax = plt.subplots(figsize=(10, 6))

    init_data = data_dict['初始']
    aging_labels = ['紫外老化', '湿热老化', '紫外湿热耦合老化']

    n_groups = len(SIO2_ORDER)
    n_bars = len(aging_labels)
    group_width = 0.65
    bar_width = group_width / n_bars
    x = np.arange(n_groups)

    # 保留率图只有3组（排除初始），使用中灰到深灰
    ret_styles = [
        {'color': '0.60', 'hatch': '///',    'edgecolor': 'black', 'label': '紫外老化'},
        {'color': '0.35', 'hatch': '\\\\\\', 'edgecolor': 'black', 'label': '湿热老化'},
        {'color': '0.15', 'hatch': '...',    'edgecolor': 'black', 'label': '紫外湿热耦合老化'},
    ]

    for i, style in enumerate(ret_styles):
        label = style['label']
        if label not in data_dict:
            continue
        d = data_dict[label]
        retentions = []
        ret_stds = []
        for sio2 in SIO2_ORDER:
            init_val = init_data[sio2][0]
            aged_mean = d[sio2][0]
            aged_std = d[sio2][1]
            if init_val > 0:
                retentions.append(aged_mean / init_val * 100)
                ret_stds.append(aged_std / init_val * 100)
            else:
                retentions.append(0)
                ret_stds.append(0)

        offset = (i - (n_bars - 1) / 2) * bar_width
        bars = ax.bar(x + offset, retentions, bar_width,
                     yerr=ret_stds, capsize=4,
                     facecolor=style['color'],
                     edgecolor=style['edgecolor'],
                     linewidth=1.0,
                     hatch=style['hatch'],
                     label=label,
                     error_kw={'elinewidth': 1.0, 'capthick': 1.0, 'ecolor': 'black'})

        for j, (ret, rs) in enumerate(zip(retentions, ret_stds)):
            if ret > 0:
                ax.text(x[j] + offset, ret + rs + 0.5, f'{ret:.1f}%',
                       ha='center', va='bottom', fontsize=14)

    ax.axhline(y=100, color='black', linestyle='--', linewidth=1, alpha=0.5)
    ax.set_xticks(x)
    ax.set_xticklabels([f'SiO\u2082 {k}' for k in SIO2_ORDER], fontsize=14)
    ax.set_ylabel('强度保留率 (%)', fontsize=14)
    ax.set_xlabel('纳米SiO\u2082含量', fontsize=14)
    ax.set_title('老化后BF/PLA/纳米SiO\u2082复合材料的力学性能保留率',
                fontsize=14, fontweight='bold', pad=15)
    ax.legend(fontsize=14, loc='best', framealpha=0.9)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_ylim(bottom=0, top=115)

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close()




if __name__ == '__main__':
    base_dir = '/Users/shiberlin/Desktop/毕业论文'
    output_dir = os.path.join(base_dir, 'charts', 'thesis')
    os.makedirs(output_dir, exist_ok=True)

    print("=" * 60)
    print("生成第七章结论章图表...")
    print("=" * 60)

    # 图7-1: 拉伸强度
    fpath = os.path.join(output_dir, '第七章_拉伸强度_分组柱状图.png')
    make_grouped_bar_4groups(TENSILE_STRENGTH, fpath,
                            '拉伸强度', 'MPa',
                            '不同老化条件下BF/PLA/纳米SiO\u2082复合材料的拉伸强度')
    print(f"  已保存: 第七章_拉伸强度_分组柱状图.png")

    # 图7-2: 弯曲强度
    fpath = os.path.join(output_dir, '第七章_弯曲强度_分组柱状图.png')
    make_grouped_bar_4groups(FLEXURAL_STRENGTH, fpath,
                            '弯曲强度', 'MPa',
                            '不同老化条件下BF/PLA/纳米SiO\u2082复合材料的弯曲强度')
    print(f"  已保存: 第七章_弯曲强度_分组柱状图.png")

    # 图7-3: 冲击强度
    fpath = os.path.join(output_dir, '第七章_冲击强度_分组柱状图.png')
    make_grouped_bar_4groups(IMPACT_STRENGTH, fpath,
                            '冲击强度', 'kJ/m\u00B2',
                            '不同老化条件下BF/PLA/纳米SiO\u2082复合材料的冲击强度')
    print(f"  已保存: 第七章_冲击强度_分组柱状图.png")

    # 图7-4: 保留率（用拉伸强度计算）
    fpath = os.path.join(output_dir, '第七章_强度保留率_分组.png')
    make_retention_chart(TENSILE_STRENGTH, fpath)
    print(f"  已保存: 第七章_强度保留率_分组.png")

    # 图7-9: 弯曲模量（不同老化条件下，全部SiO₂含量）
    fpath = os.path.join(output_dir, '第七章_弯曲模量_分组柱状图.png')
    make_grouped_bar_4groups(FLEXURAL_MODULUS, fpath,
                            '弯曲模量', 'GPa',
                            '不同老化条件下BF/PLA/纳米SiO\u2082复合材料的弯曲模量')
    print(f"  已保存: 第七章_弯曲模量_分组柱状图.png")

    # 图7-10: 冲击韧性（不同老化条件下，全部SiO₂含量）
    fpath = os.path.join(output_dir, '第七章_冲击韧性_分组柱状图.png')
    make_grouped_bar_4groups(IMPACT_TOUGHNESS, fpath,
                            '冲击韧性', 'kJ/m\u00B2',
                            '不同老化条件下BF/PLA/纳米SiO\u2082复合材料的冲击韧性')
    print(f"  已保存: 第七章_冲击韧性_分组柱状图.png")

    print(f"\n共生成 6 张第七章图表")
    print(f"输出目录: {output_dir}")
