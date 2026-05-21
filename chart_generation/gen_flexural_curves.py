#!/usr/bin/env python3
"""
生成论文弯曲应力-挠度曲线图
- 将弯曲力-位移数据转换为弯曲应力-挠度曲线
- 弯曲应力: sigma_f = 3FL/(2bh^2) = 0.6 * F(MPa)
- X轴: 挠度 (mm)
- 适配黑白打印
"""

import os
import re
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import rcParams
import openpyxl
from collections import defaultdict

# ============ 中文字体配置 ============
rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
rcParams['axes.unicode_minus'] = False
rcParams['font.size'] = 12

# ============ 弯曲试样几何参数 ============
FLEXURAL_SPAN = 64.0      # mm (跨距L)
FLEXURAL_WIDTH = 10.0     # mm (宽度b)
FLEXURAL_THICKNESS = 4.0  # mm (厚度h)
# 三点弯曲应力: sigma_f = 3FL/(2bh^2)
FLEXURAL_CONV = (3.0 * FLEXURAL_SPAN) / (2.0 * FLEXURAL_WIDTH * FLEXURAL_THICKNESS**2)  # = 0.6

# ============ 路径配置 ============
BASE_DIR = '/Users/shiberlin/Desktop/毕业论文'
FLEX_DIR = os.path.join(BASE_DIR, '\u5f2f\u66f2')  # 弯曲
CHARTS_DIR = os.path.join(BASE_DIR, 'charts')

# ============ 老化条件标签 ============
AGING_LABELS = {
    'initial': '\u521d\u59cb\u7ec4',            # 初始组
    'uv_5d': '\u7d2b\u5916\u8001\u53165\u5929',  # 紫外老化5天
    'uv_10d': '\u7d2b\u5916\u8001\u531610\u5929', # 紫外老化10天
    'hydro_5d': '\u6e7f\u70ed\u8001\u53165\u5929', # 湿热老化5天
    'hydro_10d': '\u6e7f\u70ed\u8001\u531610\u5929', # 湿热老化10天
    'coupled_5d': '\u7d2b\u5916\u6e7f\u70ed\u8026\u5408\u8001\u53165\u5929', # 紫外湿热耦合老化5天
}

# 线条样式（黑白打印适配）
LINE_STYLES = {
    0: {'color': '#000000', 'linestyle': '-', 'linewidth': 1.5, 'marker': 'o', 'markersize': 3, 'markevery': 600},
    1: {'color': '#333333', 'linestyle': '--', 'linewidth': 1.5, 'marker': 's', 'markersize': 3, 'markevery': 600},
    3: {'color': '#666666', 'linestyle': '-.', 'linewidth': 1.5, 'marker': '^', 'markersize': 3, 'markevery': 600},
    5: {'color': '#999999', 'linestyle': ':', 'linewidth': 1.5, 'marker': 'd', 'markersize': 3, 'markevery': 600},
}


def extract_sio2(filename):
    m = re.search(r'(\d+)%', filename)
    return int(m.group(1)) if m else None


def extract_sample_num(filename):
    nums = re.findall(r'(\d)(?!\d*%)', filename.replace('.xlsx', ''))
    return int(nums[-1]) if nums else None


def read_raw_data(filepath):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    forces = []
    positions = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] is not None and row[2] is not None:
            forces.append(float(row[1]))
            positions.append(float(row[2]))
    return np.array(forces), np.array(positions)


def find_load_start(forces, threshold_ratio=0.02, min_consecutive=50):
    n = len(forces)
    if n < min_consecutive:
        return 0
    window = 20
    smooth = np.convolve(forces, np.ones(window)/window, mode='same')
    max_f = np.max(smooth)
    if max_f < 1:
        return 0
    threshold = max_f * threshold_ratio
    start_idx = 0
    below_count = 0
    for i in range(n):
        if smooth[i] < threshold:
            below_count += 1
        else:
            if below_count >= min_consecutive:
                start_idx = max(0, i - below_count - 10)
                break
            below_count = 0
    if start_idx >= n * 0.8:
        start_idx = 0
    return start_idx


def process_flexural_file(filepath):
    """处理单个弯曲数据文件，返回挠度(mm)和弯曲应力(MPa)"""
    forces, positions = read_raw_data(filepath)
    if len(forces) < 100:
        return None, None, None

    start_idx = find_load_start(forces)
    forces = forces[start_idx:]
    positions = positions[start_idx:]

    if len(forces) < 100:
        return None, None, None

    # 弯曲应力
    stress = forces * FLEXURAL_CONV  # MPa
    # 挠度
    deflection = positions - positions[0]  # mm

    # 断裂点(最大应力点)
    max_idx = np.argmax(stress)

    return deflection, stress, max_idx


def categorize_flexural_files():
    """遍历弯曲目录分类所有数据文件"""
    categories = defaultdict(list)

    for root, dirs, files in os.walk(FLEX_DIR):
        for f in files:
            if not f.endswith('.xlsx') or f.startswith('~$'):
                continue
            filepath = os.path.join(root, f)
            rel = os.path.relpath(root, FLEX_DIR)

            sio2 = extract_sio2(f)
            if sio2 is None:
                continue
            sample_num = extract_sample_num(f)

            # 判断老化条件
            has_uv = '\u7d2b\u5916' in f   # 紫外
            has_sw = '\u6e7f\u70ed' in f   # 湿热
            has_cp = '5+5' in f

            if has_cp:
                aging_condition = 'coupled_5d'
            elif has_uv and has_sw:
                aging_condition = 'coupled_5d'
            elif has_uv:
                if '10' in rel or '10' in f:
                    aging_condition = 'uv_10d'
                else:
                    aging_condition = 'uv_5d'
            elif has_sw:
                if '10' in rel or '10' in f:
                    aging_condition = 'hydro_10d'
                else:
                    aging_condition = 'hydro_5d'
            else:
                # 初始组
                aging_condition = 'initial'

            key = (aging_condition, sio2)
            categories[key].append((filepath, sample_num))

    return categories


def select_representative(file_data_list):
    """从平行样品中选择最接近中位数的代表性曲线"""
    results = []
    for filepath, sample_num in file_data_list:
        deflection, stress, max_idx = process_flexural_file(filepath)
        if deflection is not None:
            results.append((filepath, sample_num, deflection, stress, max_idx))

    if not results:
        return None

    max_stresses = [np.max(r[3]) for r in results]
    median_stress = np.median(max_stresses)
    best_idx = np.argmin([abs(s - median_stress) for s in max_stresses])
    return results[best_idx]


def generate_flexural_figure(categories, aging_keys, output_name, title_prefix):
    """
    为指定老化条件生成弯曲应力-挠度曲线图
    每个子图显示4个SiO2水平的代表性曲线
    """
    # 收集有数据的条件
    valid_keys = []
    for ak in aging_keys:
        for sio2 in [0, 1, 3, 5]:
            if (ak, sio2) in categories and len(categories[(ak, sio2)]) > 0:
                valid_keys.append(ak)
                break
    valid_keys = sorted(set(valid_keys), key=lambda x: aging_keys.index(x) if x in aging_keys else 99)

    n = len(valid_keys)
    if n == 0:
        print(f"  No flexural data for {output_name}")
        return

    fig, axes = plt.subplots(1, n, figsize=(6 * n, 5))
    if n == 1:
        axes = [axes]

    for ax_idx, ak in enumerate(valid_keys):
        ax = axes[ax_idx]

        for sio2 in [0, 1, 3, 5]:
            key = (ak, sio2)
            if key not in categories:
                continue

            result = select_representative(categories[key])
            if result is None:
                continue

            filepath, sample_num, deflection, stress, max_idx = result

            style = LINE_STYLES.get(sio2, LINE_STYLES[0])
            label = f'{sio2}% SiO2'

            ax.plot(deflection, stress,
                   color=style['color'], linestyle=style['linestyle'],
                   linewidth=style['linewidth'],
                   label=label)

            # 标记断裂点(最大应力点)
            ax.scatter([deflection[max_idx]], [stress[max_idx]],
                      marker='x', color=style['color'], s=50, zorder=5)

        ax.legend(fontsize=9, loc='upper left', framealpha=0.8)
        ax.set_xlabel('\u6320\u5ea6 / mm', fontsize=12)  # 挠度 / mm
        ax.set_ylabel('\u5f2f\u66f2\u5e94\u529b / MPa', fontsize=12)  # 弯曲应力 / MPa
        ax.set_title(AGING_LABELS.get(ak, ak), fontsize=13, fontweight='bold')
        ax.tick_params(labelsize=11)
        ax.grid(True, alpha=0.3, linestyle='--')
        ax.set_xlim(left=0)
        ax.set_ylim(bottom=0)

    plt.tight_layout()
    output_path = os.path.join(CHARTS_DIR, output_name)
    fig.savefig(output_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Saved: {output_name}")


def main():
    print("=" * 60)
    print("弯曲应力-挠度曲线生成")
    print("=" * 60)

    # 1. 分类数据
    print("\n[1/2] 分类弯曲数据...")
    categories = categorize_flexural_files()

    stats = defaultdict(int)
    for (ak, sio2), files in categories.items():
        stats[ak] += len(files)
    print(f"  共 {sum(stats.values())} 个弯曲数据文件")
    for ak in sorted(stats, key=lambda x: list(AGING_LABELS.keys()).index(x) if x in AGING_LABELS else 99):
        print(f"    {AGING_LABELS.get(ak, ak):20s} | {stats[ak]} files")

    # 2. 生成曲线
    print("\n[2/2] 生成弯曲曲线...")

    # 第三章: 初始组
    generate_flexural_figure(categories, ['initial'],
        'Fig3-2_initial_flexural.png', '初始组')

    # 第四章: 紫外老化 5d + 10d
    generate_flexural_figure(categories, ['uv_5d', 'uv_10d'],
        'Fig4-2_uv_flexural.png', '紫外老化')

    # 第五章: 湿热老化 5d + 10d
    generate_flexural_figure(categories, ['hydro_5d', 'hydro_10d'],
        'Fig5-2_hydro_flexural.png', '湿热老化')

    # 第六章: 耦合老化 5d (没有弯曲10d耦合数据)
    generate_flexural_figure(categories, ['coupled_5d'],
        'Fig6-2_coupled_flexural.png', '耦合老化')

    # ===== 输出数据摘要 =====
    print("\n" + "=" * 60)
    print("弯曲性能数据摘要")
    print("=" * 60)

    for ak in ['initial', 'uv_5d', 'uv_10d', 'hydro_5d', 'hydro_10d', 'coupled_5d']:
        label = AGING_LABELS.get(ak, ak)
        print(f"\n{label}:")
        for sio2 in [0, 1, 3, 5]:
            key = (ak, sio2)
            if key not in categories:
                continue

            max_stresses = []
            deflections_at_break = []
            for filepath, sample_num in categories[key]:
                deflection, stress, max_idx = process_flexural_file(filepath)
                if deflection is not None and max_idx is not None:
                    max_stresses.append(stress[max_idx])
                    deflections_at_break.append(deflection[max_idx])

            if max_stresses:
                avg = np.mean(max_stresses)
                std = np.std(max_stresses)
                avg_def = np.mean(deflections_at_break)
                std_def = np.std(deflections_at_break)
                print(f"  SiO2={sio2}%: sigma_max={avg:.2f}+-{std:.2f} MPa, "
                      f"deflection={avg_def:.2f}+-{std_def:.2f} mm")

    print("\n完成！")


if __name__ == '__main__':
    main()
