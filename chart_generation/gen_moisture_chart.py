#!/usr/bin/env python3
"""
gen_moisture_chart.py
生成吸湿率折线图（图7-11效果），无误差棒。
数据来源：毕业论文_修改版.docx 表7-6
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
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

# ========== 数据：表7-6 不同SiO₂含量复合材料在不同时间点的吸湿率(%) ==========
time_points = [0, 6, 18, 42, 66, 114, 138]

data = {
    '0%':  [0.00, 0.23, 0.34, 1.25, 1.59, 1.71, 1.94],
    '1%':  [0.00, 0.66, 1.64, 1.75, 1.86, 2.08, 2.08],
    '3%':  [0.00, 2.28, 2.39, 2.84, 3.07, 3.41, 4.10],
    '5%':  [0.00, 0.32, 0.32, 0.86, 0.96, 0.96, 1.50],
}

SIO2_ORDER = ['0%', '1%', '3%', '5%']

# 黑白打印友好：不同线型 + 不同标记
line_styles = [
    {'marker': 'o', 'linestyle': '-',  'color': 'black', 'label': '0% SiO₂'},
    {'marker': 's', 'linestyle': '--', 'color': 'black', 'label': '1% SiO₂'},
    {'marker': '^', 'linestyle': '-.', 'color': 'black', 'label': '3% SiO₂'},
    {'marker': 'D', 'linestyle': ':',  'color': 'black', 'label': '5% SiO₂'},
]

fig, ax = plt.subplots(figsize=(10, 6))

for i, sio2 in enumerate(SIO2_ORDER):
    values = data[sio2]
    style = line_styles[i]
    ax.plot(time_points, values,
            marker=style['marker'],
            linestyle=style['linestyle'],
            color=style['color'],
            linewidth=1.5,
            markersize=8,
            markerfacecolor='white',
            markeredgecolor='black',
            markeredgewidth=1.5,
            label=style['label'])

ax.set_xlabel('吸湿时间 (h)', fontsize=14)
ax.set_ylabel('吸湿率 (%)', fontsize=14)
ax.set_title('不同SiO\u2082含量BF/PLA复合材料的吸湿率随时间变化曲线',
             fontsize=14, fontweight='bold', pad=15)

ax.legend(fontsize=12, loc='lower right', framealpha=0.9)
ax.grid(True, axis='y', alpha=0.3, linestyle='--')
ax.set_xlim(left=0)
ax.set_ylim(bottom=0, top=5.0)

ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

plt.tight_layout()

# 保存到 charts/ 目录（与旧图7-11同目录）
output_dir = '/Users/shiberlin/Desktop/毕业论文/charts'
os.makedirs(output_dir, exist_ok=True)
fpath = os.path.join(output_dir, '图7-11_不同老化条件下吸湿率变化曲线图.png')
plt.savefig(fpath, dpi=200, bbox_inches='tight')
plt.close()

print(f"已保存: {fpath}")
print(f"大小: {os.path.getsize(fpath)} bytes")
