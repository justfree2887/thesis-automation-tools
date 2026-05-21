#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成第七章红外光谱图（图7-5至图7-8）— 坐标偏移 + 黑白打印友好样式"""

import csv
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

# ── 字体设置 ──────────────────────────────────────────────
plt.rcParams['font.sans-serif'] = ['SimHei', 'Heiti SC', 'Arial Unicode MS', 'STHeiti', 'Hiragino Sans GB']
plt.rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 14
plt.rcParams['mathtext.fontset'] = 'custom'
plt.rcParams['mathtext.rm'] = 'Times New Roman'
plt.rcParams['mathtext.it'] = 'Times New Roman:italic'
plt.rcParams['mathtext.bf'] = 'Times New Roman:bold'

# ── 路径 ──────────────────────────────────────────────────
DATA_DIR = '/Users/shiberlin/Desktop/毕业论文/红外光谱'
OUT_DIR  = '/Users/shiberlin/Desktop/毕业论文/charts'

# ── 配置 ──────────────────────────────────────────────────
SIO2_LEVELS = ['0', '1', '3', '5']
CONDITIONS  = ['初始', '湿热', '紫外', '紫外湿热']

# 黑白打印友好的4种线型
LINE_CONFIGS = {
    '0': {'ls': '-',   'lw': 1.2, 'label': '0% SiO$_2$'},
    '1': {'ls': '--',  'lw': 1.2, 'label': '1% SiO$_2$'},
    '3': {'ls': '-.',  'lw': 1.2, 'label': '3% SiO$_2$'},
    '5': {'ls': (0, (3, 1, 1, 1)), 'lw': 1.8, 'label': '5% SiO$_2$'},  # 虚-点-点线（比单纯点线更清晰）
}

# 纵轴偏移量
OFFSET_STEP = 25

# 关键特征峰
KEY_PEAKS = {
    '初始':      [(3340, '3340'), (1750, '1750'), (1180, '1180'), (1085, '1085')],
    '湿热':      [(3340, '3340'), (1750, '1750'), (1710, '1710'), (1600, '1600'), (1510, '1510')],
    '紫外':      [(1750, '1750'), (1710, '1710')],
    '紫外湿热':  [(3340, '3340'), (1750, '1750'), (1710, '1710'), (1100, '1100')],
}


def load_ftir_csv(filepath):
    """读取CSV：两列（波数, 透射率），无表头"""
    wavenums, trans = [], []
    with open(filepath, 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            wavenums.append(float(row[0]))
            trans.append(float(row[1]))
    return np.array(wavenums), np.array(trans)


def make_ftir_chart(condition, output_path):
    """生成一张FTIR光谱叠加图（4个SiO₂含量，纵轴偏移）"""
    fig = plt.figure(figsize=(7.5, 6.0), dpi=300)
    ax = fig.add_axes([0.11, 0.08, 0.82, 0.85])

    # 先加载所有数据，计算全局y轴范围
    all_data = {}
    for sio2 in SIO2_LEVELS:
        fpath = os.path.join(DATA_DIR, f'{sio2}{condition}.CSV')
        wn, tr = load_ftir_csv(fpath)
        offset = OFFSET_STEP * SIO2_LEVELS.index(sio2)
        all_data[sio2] = (wn, tr, offset)

    # y轴范围：基于实际数据 + 偏移量计算
    y_min_global = min(tr.min() + offset for (_, tr, offset) in all_data.values())
    y_max_global = max(tr.max() + offset for (_, tr, offset) in all_data.values())
    y_margin = (y_max_global - y_min_global) * 0.05
    y_bottom = y_min_global - y_margin
    y_top = y_max_global + y_margin

    # 绘制
    for sio2 in SIO2_LEVELS:
        wn, tr, offset = all_data[sio2]
        cfg = LINE_CONFIGS[sio2]
        ax.plot(wn, tr + offset,
                linestyle=cfg['ls'],
                linewidth=cfg['lw'],
                color='black',
                label=cfg['label'])

        # 基线虚线
        ax.axhline(y=offset, color='gray', lw=0.3, ls='-', alpha=0.3)

    # ── 坐标轴 ──
    ax.set_xlim(4000, 400)   # 高波数在左（FTIR惯例）
    ax.set_ylim(y_bottom, y_top)
    ax.set_xlabel('波数 / cm$^{-1}$', fontsize=14)
    ax.set_ylabel('透射率 / %', fontsize=14)
    ax.tick_params(direction='in', top=True, right=True, labelsize=14)

    # ── 关键特征峰标注 ──
    peaks = KEY_PEAKS.get(condition, [])
    for wn_val, wn_label in peaks:
        ax.axvline(x=wn_val, color='gray', lw=0.4, ls='--', alpha=0.5)
        ax.text(wn_val, y_top - (y_top - y_bottom) * 0.02, wn_label,
                fontsize=14, ha='center', va='top', color='gray')

    # ── 图例（紧贴坐标轴右下角内侧）──
    ax.legend(loc='lower right', bbox_to_anchor=(1.0, 0.0),
              fontsize=14, frameon=True, edgecolor='black',
              fancybox=False, handlelength=3.5,
              borderpad=0.4, labelspacing=0.4)

    # ── 保存 ──
    fig.savefig(output_path, dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close(fig)
    print(f'已生成: {output_path}  (y范围: {y_bottom:.0f} ~ {y_top:.0f})')


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for condition in CONDITIONS:
        fname = f'第七章_{condition}_红外光谱.png'
        fpath = os.path.join(OUT_DIR, fname)
        make_ftir_chart(condition, fpath)
    print('\n全部完成！')


if __name__ == '__main__':
    main()
