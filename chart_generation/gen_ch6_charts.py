#!/usr/bin/env python3
"""
gen_ch6_charts.py
生成第六章（紫外湿热耦合老化）的柱状图。
数据来源：stats_v2.json
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import json, os

plt.rcParams['font.sans-serif'] = ['SimHei', 'Heiti SC', 'Arial Unicode MS', 'STHeiti', 'Hiragino Sans GB']
plt.rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 14
plt.rcParams['mathtext.fontset'] = 'custom'
plt.rcParams['mathtext.rm'] = 'Times New Roman'
plt.rcParams['mathtext.it'] = 'Times New Roman:italic'
plt.rcParams['mathtext.bf'] = 'Times New Roman:bold'

BASE = '/Users/shiberlin/Desktop/毕业论文'
SIO2_ORDER = ['0%', '1%', '3%', '5%']

with open(os.path.join(BASE, 'charts/thesis/stats_v2.json'), 'r', encoding='utf-8') as f:
    stats = json.load(f)

# 黑白打印友好：渐变灰度 + 填充图案
COLORS = ['0.85', '0.60', '0.35', '0.15']
HATCHES = ['', '///', '\\\\\\', '...']


def make_single_bar_chart(data_dict, output_path, ylabel, y_unit, title):
    """生成4根柱的柱状图（一个老化条件下4种SiO2含量对比）"""
    fig, ax = plt.subplots(figsize=(8, 5))

    x = np.arange(len(SIO2_ORDER))
    bar_width = 0.55

    means = []
    stds = []
    for sio2 in SIO2_ORDER:
        if sio2 in data_dict:
            means.append(data_dict[sio2]['max_load_mean'])
            stds.append(data_dict[sio2]['max_load_std'])
        else:
            means.append(0)
            stds.append(0)

    bars = ax.bar(x, means, bar_width,
                  yerr=stds, capsize=5,
                  color=COLORS,
                  edgecolor='black',
                  linewidth=1.0,
                  hatch=HATCHES,
                  error_kw={'elinewidth': 1.0, 'capthick': 1.0, 'ecolor': 'black'})

    # 数值标注
    y_max_all = max(means) if max(means) > 0 else 1
    for j, (mean, std) in enumerate(zip(means, stds)):
        if mean > 0:
            y_pos = mean + std + y_max_all * 0.02
            ax.text(x[j], y_pos, f'{mean:.1f}',
                   ha='center', va='bottom', fontsize=13)

    ax.set_xticks(x)
    ax.set_xticklabels([f'SiO\u2082 {k}' for k in SIO2_ORDER], fontsize=14)
    ax.set_ylabel(f'{ylabel} ({y_unit})', fontsize=14)
    ax.set_xlabel('纳米SiO\u2082含量', fontsize=14)
    ax.set_title(title, fontsize=14, fontweight='bold', pad=15)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # y轴上限
    all_vals = [m + s for m, s in zip(means, stds)]
    if all_vals:
        ax.set_ylim(bottom=0, top=max(all_vals) * 1.22)

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close()
    print(f"  已保存: {os.path.basename(output_path)}")


if __name__ == '__main__':
    output_dir = os.path.join(BASE, 'charts', 'thesis')
    os.makedirs(output_dir, exist_ok=True)

    print("=" * 60)
    print("生成第六章紫外湿热耦合老化柱状图...")
    print("=" * 60)

    ts = stats['拉伸']['紫外湿热耦合老化组']

    # 图6-2: 拉伸强度
    fpath = os.path.join(output_dir, '第六章_拉伸_强度柱状图.png')
    make_single_bar_chart(ts, fpath,
                          '拉伸强度', 'N',
                          '紫外湿热耦合老化后不同SiO\u2082含量复合材料的拉伸强度')

    print(f"\n完成")
