#!/usr/bin/env python3
"""
generate_thesis_charts.py
为毕业论文第三到第六章生成所需的图表
- 4种老化条件 × 4种SiO₂含量 × (拉伸/弯曲) × (试验曲线/强度柱状图/位移柱状图)
- 黑白打印机友好
- 每个SiO₂含量独立一张图（单图模式）
- 输出到 charts/thesis/
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd
import os
import re
import json

# ========== 中文字体配置 ==========
plt.rcParams['font.sans-serif'] = ['SimHei', 'Heiti SC', 'Arial Unicode MS', 'STHeiti', 'Hiragino Sans GB']
plt.rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 14  # 四号字=14pt
plt.rcParams['mathtext.fontset'] = 'custom'
plt.rcParams['mathtext.rm'] = 'Times New Roman'
plt.rcParams['mathtext.it'] = 'Times New Roman:italic'
plt.rcParams['mathtext.bf'] = 'Times New Roman:bold'

# ========== 黑白打印机友好配置 ==========
grayscale = {
    '0%': 0.0,    # 纯黑
    '1%': 0.25,   # 深灰
    '3%': 0.55,   # 中灰
    '5%': 0.25,   # 深灰（原0.85太白，黑白打印无法分辨）
}

marker_styles = {
    '0%': 's',     # 正方形
    '1%': 'o',     # 圆形
    '3%': '^',     # 上三角形
    '5%': 'D',     # 菱形
}

line_styles = {
    '0%': '-',             # 实线
    '1%': '--',            # 虚线
    '3%': '-.',            # 点划线
    '5%': (0, (2, 1, 0.5, 1)),  # 密疏点线（比单纯点线更清晰）
}

hatch_patterns = {
    '0%': '',       # 无填充
    '1%': '///',    # 斜线
    '3%': 'xxx',    # 叉号
    '5%': '...',    # 点
}

# 章节映射
CHAPTER_MAP = {
    '初始组': ('第三章', '初始组'),
    '紫外老化组': ('第四章', '紫外老化'),
    '湿热老化组': ('第五章', '湿热老化'),
    '紫外湿热耦合老化组': ('第六章', '紫外湿热耦合老化'),
}

SIO2_ORDER = ['0%', '1%', '3%', '5%']

# ========== 工具函数 ==========
def remove_outliers_3sigma(data_list, threshold=3):
    if len(data_list) < 3:
        return data_list, []
    arr = np.array(data_list)
    mean = np.mean(arr)
    std = np.std(arr, ddof=1)
    if std == 0:
        return data_list, []
    mask = np.abs(arr - mean) <= threshold * std
    filtered = arr[mask].tolist()
    removed_indices = np.where(~mask)[0].tolist()
    return filtered, removed_indices

def extract_aging_condition(filename):
    """从文件名提取老化条件（用于分类到对应章节）"""
    if '初始' in filename:
        return '初始组'
    elif '紫外湿热' in filename:
        return '紫外湿热耦合老化组'
    elif '紫外' in filename:
        return '紫外老化组'
    elif '湿热' in filename:
        return '湿热老化组'
    else:
        return None

def extract_sio2_content(filename):
    for pct in SIO2_ORDER:
        if pct in filename:
            return pct
    return None

def read_excel_data(filepath):
    try:
        df = pd.read_excel(filepath)
        df['PositionValue'] = df['PositionValue'] - df['PositionValue'].iloc[0]
        return df
    except Exception as e:
        print(f"  读取失败 {os.path.basename(filepath)}: {e}")
        return None

def interp_curves(samples, n_pts=400):
    if not samples:
        return None, None, None
    x_min = max(s['PositionValue'].min() for s in samples)
    x_max = min(s['PositionValue'].max() for s in samples)
    if x_max <= x_min:
        return None, None, None
    x_common = np.linspace(x_min, x_max, n_pts)
    y_all = []
    for s in samples:
        mask = (s['PositionValue'] >= x_min) & (s['PositionValue'] <= x_max)
        if mask.sum() < 2:
            continue
        y_interp = np.interp(x_common,
                             s['PositionValue'][mask].values,
                             s['LoadValue'][mask].values)
        y_all.append(y_interp)
    if len(y_all) == 0:
        return None, None, None
    y_arr = np.array(y_all)
    return x_common, y_arr.mean(axis=0), y_arr.std(axis=0)

# ========== 数据处理 ==========
def process_test_type(base_dir, test_type):
    """读取并分类所有数据"""
    test_dir = os.path.join(base_dir, test_type)
    if not os.path.exists(test_dir):
        return {}

    result = {}

    for folder_name in os.listdir(test_dir):
        folder_path = os.path.join(test_dir, folder_name)
        if not os.path.isdir(folder_path) or '.workbuddy' in folder_path:
            continue

        for fname in os.listdir(folder_path):
            if not fname.endswith('.xlsx') or fname.startswith('.'):
                continue

            aging_cond = extract_aging_condition(fname)
            sio2_pct = extract_sio2_content(fname)

            if aging_cond is None or sio2_pct is None:
                continue

            if aging_cond not in result:
                result[aging_cond] = {}
            if sio2_pct not in result[aging_cond]:
                result[aging_cond][sio2_pct] = {'dataframes': [], 'max_loads': [], 'max_disps': []}

            filepath = os.path.join(folder_path, fname)
            df = read_excel_data(filepath)
            if df is None:
                continue

            result[aging_cond][sio2_pct]['dataframes'].append(df)
            result[aging_cond][sio2_pct]['max_loads'].append(df['LoadValue'].max())
            result[aging_cond][sio2_pct]['max_disps'].append(df['PositionValue'].max())

    return result

def filter_and_aggregate(data_dict):
    """剔除异常值并计算统计量"""
    result = {}

    for aging_cond, sio2_data in data_dict.items():
        result[aging_cond] = {}

        for sio2_pct, data in sio2_data.items():
            if len(data['max_loads']) == 0:
                continue

            filtered_loads, removed_loads = remove_outliers_3sigma(data['max_loads'])
            filtered_disps, removed_disps = remove_outliers_3sigma(data['max_disps'])

            all_removed = set(removed_loads + removed_disps)
            valid_indices = [i for i in range(len(data['dataframes'])) if i not in all_removed]

            if len(valid_indices) == 0:
                continue

            samples = [data['dataframes'][i] for i in valid_indices]
            x_common, y_mean, y_std = interp_curves(samples)

            valid_loads = [data['max_loads'][i] for i in valid_indices]
            valid_disps = [data['max_disps'][i] for i in valid_indices]

            result[aging_cond][sio2_pct] = {
                'x': x_common,
                'y_mean': y_mean,
                'y_std': y_std,
                'max_load_mean': np.mean(valid_loads),
                'max_load_std': np.std(valid_loads, ddof=1) if len(valid_loads) > 1 else 0,
                'max_disp_mean': np.mean(valid_disps),
                'max_disp_std': np.std(valid_disps, ddof=1) if len(valid_disps) > 1 else 0,
                'n_samples': len(valid_indices)
            }

    return result

# ========== 图表生成函数 ==========

def make_single_curve_chart(data_dict, aging_cond, test_type, sio2_pct, output_path):
    """为某个SiO₂含量生成单张试验曲线图（单条曲线+误差带）"""
    if sio2_pct not in data_dict:
        return

    d = data_dict[sio2_pct]
    if d['x'] is None:
        return

    fig, ax = plt.subplots(figsize=(7, 5))

    gray_val = grayscale[sio2_pct]
    color = str(gray_val)

    # 绘制均值曲线
    ax.plot(d['x'], d['y_mean'],
            label=f'SiO₂ {sio2_pct} (n={d["n_samples"]})',
            color=color,
            linewidth=2,
            linestyle=line_styles[sio2_pct],
            marker=marker_styles[sio2_pct],
            markersize=4,
            markevery=25,
            markerfacecolor='white',
            markeredgecolor=color)

    # 绘制误差带
    ax.fill_between(d['x'],
                    d['y_mean'] - d['y_std'],
                    d['y_mean'] + d['y_std'],
                    color=color, alpha=0.2)

    ax.set_xlabel('位移 (mm)', fontsize=14)
    ax.set_ylabel('载荷 (N)', fontsize=14)
    title_prefix = CHAPTER_MAP.get(aging_cond, (aging_cond, aging_cond))[1]
    ax.set_title(f'{title_prefix}纳米SiO₂含量{sio2_pct}的{test_type}试验曲线',
                fontsize=14, fontweight='bold', pad=12)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.legend(fontsize=14, loc='best', framealpha=0.9)

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close()

def make_all_in_one_curve_chart(data_dict, aging_cond, test_type, output_path):
    """为某老化条件的所有SiO₂含量在一张图上绘制试验曲线（对比用）"""
    fig, ax = plt.subplots(figsize=(9, 6))

    has_data = False
    for pct in SIO2_ORDER:
        if pct not in data_dict or data_dict[pct]['x'] is None:
            continue
        d = data_dict[pct]
        has_data = True
        gray_val = grayscale[pct]
        color = str(gray_val)

        ax.plot(d['x'], d['y_mean'],
                label=f'SiO₂ {pct} (n={d["n_samples"]})',
                color=color,
                linewidth=2,
                linestyle=line_styles[pct],
                marker=marker_styles[pct],
                markersize=4,
                markevery=25,
                markerfacecolor='white',
                markeredgecolor=color)

        alpha_val = 0.12 + gray_val * 0.2
        ax.fill_between(d['x'],
                        d['y_mean'] - d['y_std'],
                        d['y_mean'] + d['y_std'],
                        color=color, alpha=alpha_val)

    if not has_data:
        plt.close()
        return

    ax.set_xlabel('位移 (mm)', fontsize=14)
    ax.set_ylabel('载荷 (N)', fontsize=14)
    title_prefix = CHAPTER_MAP.get(aging_cond, (aging_cond, aging_cond))[1]
    ax.set_title(f'{title_prefix}PLA/竹粉/纳米SiO₂复合材料的{test_type}试验曲线',
                fontsize=14, fontweight='bold', pad=15)
    ax.legend(fontsize=14, loc='best', framealpha=0.9)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close()

def make_bar_chart_single(data_dict, aging_cond, test_type, sio2_pct, output_path, ylabel, y_unit, is_displacement=False):
    """为某个SiO₂含量生成单张柱状图（与其他老化时间对比时使用）"""
    pass  # 这种场景由 make_aging_change_charts.py 处理

def make_bar_chart_all(data_dict, aging_cond, test_type, output_path, ylabel, y_unit, is_displacement=False):
    """为某老化条件生成所有SiO₂含量的柱状图"""
    fig, ax = plt.subplots(figsize=(8, 5.5))

    x_pos = np.arange(len(SIO2_ORDER))
    bar_width = 0.55

    means = []
    stds = []
    ns = []

    for key in SIO2_ORDER:
        if key in data_dict:
            if is_displacement:
                means.append(data_dict[key]['max_disp_mean'])
                stds.append(data_dict[key]['max_disp_std'])
            else:
                means.append(data_dict[key]['max_load_mean'])
                stds.append(data_dict[key]['max_load_std'])
            ns.append(data_dict[key]['n_samples'])
        else:
            means.append(0)
            stds.append(0)
            ns.append(0)

    if all(m == 0 for m in means):
        plt.close()
        return

    bars = ax.bar(x_pos, means, bar_width,
                  yerr=stds,
                  capsize=5,
                  color=[str(grayscale[k]) for k in SIO2_ORDER],
                  edgecolor='black',
                  linewidth=1.2,
                  error_kw={'elinewidth': 1.5, 'capthick': 1.5, 'ecolor': 'black'})

    for bar, key in zip(bars, SIO2_ORDER):
        if key in data_dict and data_dict[key].get('n_samples', 0) > 0:
            bar.set_hatch(hatch_patterns[key])

    for i, (mean, std) in enumerate(zip(means, stds)):
        if mean > 0:
            y_max = max(means) if max(means) > 0 else 1
            y_pos = mean + std + y_max * 0.03
            ax.text(i, y_pos, f'{mean:.1f}',
                   ha='center', va='bottom', fontsize=14)

    ax.set_xticks(x_pos)
    ax.set_xticklabels([f'SiO₂ {k}' for k in SIO2_ORDER], fontsize=14)
    ax.set_ylabel(f'{ylabel} ({y_unit})', fontsize=14)
    ax.set_xlabel('纳米SiO₂含量', fontsize=14)
    title_prefix = CHAPTER_MAP.get(aging_cond, (aging_cond, aging_cond))[1]
    ax.set_title(f'{title_prefix}PLA/竹粉/纳米SiO₂复合材料的{test_type}{ylabel}',
                fontsize=14, fontweight='bold', pad=15)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    y_max_val = max([m + s for m, s in zip(means, stds)]) if means else 0
    if y_max_val > 0:
        ax.set_ylim(bottom=0, top=y_max_val * 1.18)

    legend_handles = []
    for key in SIO2_ORDER:
        if key in data_dict and data_dict[key].get('n_samples', 0) > 0:
            handle = mpatches.Patch(
                facecolor=str(grayscale[key]),
                edgecolor='black',
                hatch=hatch_patterns[key],
                label=f'{key} (n={data_dict[key]["n_samples"]})'
            )
            legend_handles.append(handle)

    if legend_handles:
        ax.legend(handles=legend_handles, fontsize=14, loc='best', framealpha=0.9)

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close()

# ========== 生成力学性能保留率图表 ==========
def make_retention_chart(processed_data, aging_cond, test_type, output_path, is_displacement=False):
    """
    生成力学性能保留率柱状图（仅对老化章节：第四、五、六章）
    保留率 = 老化后/初始 × 100%
    """
    if aging_cond == '初始组':
        return  # 初始组无保留率

    if '初始组' not in processed_data or aging_cond not in processed_data:
        return

    initial = processed_data['初始组']
    aged = processed_data[aging_cond]

    fig, ax = plt.subplots(figsize=(8, 5.5))

    means = []
    stds = []
    valid_keys = []

    for key in SIO2_ORDER:
        if key not in aged or key not in initial:
            continue

        if is_displacement:
            init_val = initial[key]['max_disp_mean']
            aged_mean = aged[key]['max_disp_mean']
            aged_std = aged[key]['max_disp_std']
        else:
            init_val = initial[key]['max_load_mean']
            aged_mean = aged[key]['max_load_mean']
            aged_std = aged[key]['max_load_std']

        if init_val > 0:
            retention = (aged_mean / init_val) * 100
            # 传播误差
            retention_std = (aged_std / init_val) * 100
            means.append(retention)
            stds.append(retention_std)
            valid_keys.append(key)

    if not means:
        plt.close()
        return

    x_pos = np.arange(len(valid_keys))
    bar_width = 0.55

    bars = ax.bar(x_pos, means, bar_width,
                  yerr=stds,
                  capsize=5,
                  color=[str(grayscale[k]) for k in valid_keys],
                  edgecolor='black',
                  linewidth=1.2,
                  error_kw={'elinewidth': 1.5, 'capthick': 1.5, 'ecolor': 'black'})

    for bar, key in zip(bars, valid_keys):
        bar.set_hatch(hatch_patterns[key])

    for i, (mean, std) in enumerate(zip(means, stds)):
        y_pos = mean + std + 1
        ax.text(i, y_pos, f'{mean:.1f}%',
               ha='center', va='bottom', fontsize=14)

    # 添加100%参考线
    ax.axhline(y=100, color='black', linestyle='--', linewidth=1, alpha=0.5)

    ax.set_xticks(x_pos)
    ax.set_xticklabels([f'SiO₂ {k}' for k in valid_keys], fontsize=14)
    ylabel = '位移保留率' if is_displacement else '强度保留率'
    ax.set_ylabel(f'{ylabel} (%)', fontsize=14)
    ax.set_xlabel('纳米SiO₂含量', fontsize=14)
    title_prefix = CHAPTER_MAP.get(aging_cond, (aging_cond, aging_cond))[1]
    ax.set_title(f'{title_prefix}PLA/竹粉/纳米SiO₂复合材料的{test_type}{ylabel}',
                fontsize=14, fontweight='bold', pad=15)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    y_max_val = max([m + s for m, s in zip(means, stds)]) if means else 0
    if y_max_val > 0:
        ax.set_ylim(bottom=0, top=max(y_max_val * 1.18, 110))

    legend_handles = []
    for key in valid_keys:
        handle = mpatches.Patch(
            facecolor=str(grayscale[key]),
            edgecolor='black',
            hatch=hatch_patterns[key],
            label=key
        )
        legend_handles.append(handle)

    ax.legend(handles=legend_handles, fontsize=14, loc='best', framealpha=0.9)

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close()

# ========== 主程序 ==========
if __name__ == '__main__':
    base_dir = '/Users/shiberlin/Desktop/毕业论文'
    output_dir = os.path.join(base_dir, 'charts', 'thesis')
    os.makedirs(output_dir, exist_ok=True)

    # 存储数值数据用于后续论文修改
    all_stats = {}
    all_processed = {}

    test_types = ['拉伸', '弯曲']

    for test_type in test_types:
        print(f"\n{'='*60}")
        print(f"处理{test_type}试验数据")
        print('='*60)

        raw_data = process_test_type(base_dir, test_type)

        if not raw_data:
            print(f"  无数据")
            continue

        print(f"  老化条件: {list(raw_data.keys())}")
        print(f"  剔除异常值...")
        processed_data = filter_and_aggregate(raw_data)

        if not processed_data:
            print(f"  处理后无有效数据")
            continue

        all_processed[test_type] = processed_data

        for aging_cond in ['初始组', '紫外老化组', '湿热老化组', '紫外湿热耦合老化组']:
            if aging_cond not in processed_data:
                print(f"  {aging_cond}: 无数据")
                continue

            data = processed_data[aging_cond]
            chapter_info = CHAPTER_MAP[aging_cond]
            print(f"\n  生成 {chapter_info[0]} ({aging_cond}) 的图表...")

            # ===== 1. 单图模式：每个SiO₂含量一张试验曲线图 =====
            for sio2_pct in SIO2_ORDER:
                if sio2_pct not in data or data[sio2_pct]['x'] is None:
                    continue
                fname = f'{chapter_info[0]}_{test_type}_曲线_{sio2_pct}.png'
                fpath = os.path.join(output_dir, fname)
                make_single_curve_chart(data, aging_cond, test_type, sio2_pct, fpath)
                print(f"    已保存: {fname}")

            # ===== 2. 合并曲线图（所有SiO₂在一张图上） =====
            fname = f'{chapter_info[0]}_{test_type}_曲线_全.png'
            fpath = os.path.join(output_dir, fname)
            make_all_in_one_curve_chart(data, aging_cond, test_type, fpath)
            print(f"    已保存: {fname}")

            # ===== 3. 强度柱状图 =====
            fname = f'{chapter_info[0]}_{test_type}_强度柱状图.png'
            fpath = os.path.join(output_dir, fname)
            make_bar_chart_all(data, aging_cond, test_type, fpath,
                             ylabel=f'{test_type}强度', y_unit='N', is_displacement=False)
            print(f"    已保存: {fname}")

            # ===== 4. 位移柱状图 =====
            fname = f'{chapter_info[0]}_{test_type}_位移柱状图.png'
            fpath = os.path.join(output_dir, fname)
            make_bar_chart_all(data, aging_cond, test_type, fpath,
                             ylabel=f'{test_type}位移', y_unit='mm', is_displacement=True)
            print(f"    已保存: {fname}")

            # ===== 5. 力学性能保留率（仅老化章节） =====
            if aging_cond != '初始组':
                for is_disp in [False, True]:
                    ylabel_r = '位移保留率' if is_disp else '强度保留率'
                    suffix = '位移保留率' if is_disp else '强度保留率'
                    fname = f'{chapter_info[0]}_{test_type}_{suffix}.png'
                    fpath = os.path.join(output_dir, fname)
                    make_retention_chart(processed_data, aging_cond, test_type, fpath, is_displacement=is_disp)
                    print(f"    已保存: {fname}")

        # ===== 收集统计数据 =====
        all_stats[test_type] = {}
        for aging_cond, data in processed_data.items():
            all_stats[test_type][aging_cond] = {}
            for sio2_pct, d in data.items():
                all_stats[test_type][aging_cond][sio2_pct] = {
                    'max_load_mean': round(float(d['max_load_mean']), 2),
                    'max_load_std': round(float(d['max_load_std']), 2),
                    'max_disp_mean': round(float(d['max_disp_mean']), 2),
                    'max_disp_std': round(float(d['max_disp_std']), 2),
                    'n_samples': int(d['n_samples'])
                }

    # 保存统计数据为JSON供后续论文修改使用
    stats_path = os.path.join(output_dir, 'stats.json')
    with open(stats_path, 'w', encoding='utf-8') as f:
        json.dump(all_stats, f, ensure_ascii=False, indent=2)
    print(f"\n统计数据已保存到: {stats_path}")

    # 打印图表统计
    chart_files = [f for f in os.listdir(output_dir) if f.endswith('.png')]
    print(f"\n{'='*60}")
    print(f"共生成 {len(chart_files)} 张图表")
    print(f"输出目录: {output_dir}")
    print('='*60)
