#!/usr/bin/env python3
"""
gen_v2_charts.py
1. 分别处理初始组、5天紫外/湿热、10天紫外/湿热、耦合老化的数据
2. 生成包含初始+5天+10天的分组柱状图（用于第四章和第五章）
3. 保存 stats_v2.json
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
import copy

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

# 分组柱状图的视觉样式（黑白打印友好 - 灰度渐变 + 不同填充）
GROUP_STYLES = {
    '初始':     {'color': '0.85',  'hatch': '',       'edgecolor': 'black'},  # 浅灰，无填充
    '5天老化':  {'color': '0.50',  'hatch': '///',    'edgecolor': 'black'},  # 中灰，斜线
    '10天老化': {'color': '0.15',  'hatch': '...',    'edgecolor': 'black'},  # 深灰，点状
}

# ========== 工具函数（复用） ==========
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

def compute_stats(raw_data, tensile_gauge=50, bending_gauge=64, test_type='拉伸'):
    """从原始数据（含dataframes, max_loads, max_disps）计算统计量"""
    gauge = tensile_gauge if test_type == '拉伸' else bending_gauge
    result = {}
    for sio2_pct, data in raw_data.items():
        if len(data['max_loads']) == 0:
            continue
        filtered_loads, removed_loads = remove_outliers_3sigma(data['max_loads'])
        filtered_disps, removed_disps = remove_outliers_3sigma(data['max_disps'])
        all_removed = set(removed_loads + removed_disps)
        valid_indices = [i for i in range(len(data['dataframes'])) if i not in all_removed]
        if len(valid_indices) == 0:
            continue
        valid_loads = [data['max_loads'][i] for i in valid_indices]
        valid_disps = [data['max_disps'][i] for i in valid_indices]
        # 将位移转换为延伸率(%): elongation = displacement / gauge * 100
        valid_elongations = [d / gauge * 100 for d in valid_disps]
        result[sio2_pct] = {
            'max_load_mean': round(float(np.mean(valid_loads)), 2),
            'max_load_std': round(float(np.std(valid_loads, ddof=1) if len(valid_loads) > 1 else 0), 2),
            'max_elongation_mean': round(float(np.mean(valid_elongations)), 2),
            'max_elongation_std': round(float(np.std(valid_elongations, ddof=1) if len(valid_elongations) > 1 else 0), 2),
            'max_disp_mean': round(float(np.mean(valid_disps)), 2),   # 保留原始位移供参考
            'max_disp_std': round(float(np.std(valid_disps, ddof=1) if len(valid_disps) > 1 else 0), 2),
            'n_samples': int(len(valid_indices)),
        }
    return result


# ========== 数据处理 ==========
def process_folder(folder_path, aging_label, file_filter=None):
    """处理单个文件夹中的xlsx文件，按aging_label分组"""
    result = {}
    if not os.path.exists(folder_path):
        return result
    for fname in sorted(os.listdir(folder_path)):
        if not fname.endswith('.xlsx') or fname.startswith('.'):
            continue
        if file_filter and not file_filter(fname):
            continue
        sio2 = extract_sio2_content(fname)
        if sio2 is None:
            continue
        if sio2 not in result:
            result[sio2] = {'dataframes': [], 'max_loads': [], 'max_disps': []}
        fpath = os.path.join(folder_path, fname)
        df = read_excel_data(fpath)
        if df is None:
            continue
        result[sio2]['dataframes'].append(df)
        result[sio2]['max_loads'].append(float(df['LoadValue'].max()))
        result[sio2]['max_disps'].append(float(df['PositionValue'].max()))
    return result


def process_all_data(base_dir):
    """处理所有数据，返回按老化条件×试验类型×SiO2含量分组的统计量"""
    all_stats = {}
    for test_type in ['拉伸', '弯曲']:
        print(f"\n处理{test_type}试验数据...")
        all_stats[test_type] = {}

        # 1. 初始组
        init_dir = os.path.join(base_dir, test_type, '初始拉伸' if test_type == '拉伸' else '弯曲初始')
        if os.path.exists(init_dir):
            raw = process_folder(init_dir, '初始组')
            all_stats[test_type]['初始组'] = compute_stats(raw, test_type=test_type)
            print(f"  初始组: {len(all_stats[test_type]['初始组'])} 个含量")

        # 2. 5天老化 - 紫外
        dir_5d = os.path.join(base_dir, test_type, '5天老化')
        if os.path.exists(dir_5d):
            raw_uv = process_folder(dir_5d, '紫外', lambda f: '紫外' in f and '湿热' not in f)
            all_stats[test_type]['紫外老化5天'] = compute_stats(raw_uv, test_type=test_type)
            print(f"  紫外老化5天: {len(all_stats[test_type]['紫外老化5天'])} 个含量")

            # 3. 5天老化 - 湿热
            raw_heat = process_folder(dir_5d, '湿热', lambda f: '湿热' in f and '紫外' not in f)
            all_stats[test_type]['湿热老化5天'] = compute_stats(raw_heat, test_type=test_type)
            print(f"  湿热老化5天: {len(all_stats[test_type]['湿热老化5天'])} 个含量")

            # 4. 紫外湿热耦合老化（只有5+5天）
            raw_coupled = process_folder(dir_5d, '耦合', lambda f: '紫外湿热' in f)
            all_stats[test_type]['紫外湿热耦合老化组'] = compute_stats(raw_coupled, test_type=test_type)
            print(f"  耦合老化: {len(all_stats[test_type]['紫外湿热耦合老化组'])} 个含量")

        # 5. 10天老化 - 紫外
        dir_10d = os.path.join(base_dir, test_type, '10天老化')
        if os.path.exists(dir_10d):
            raw_uv10 = process_folder(dir_10d, '紫外10天', lambda f: '紫外老化10天' in f)
            all_stats[test_type]['紫外老化10天'] = compute_stats(raw_uv10, test_type=test_type)
            print(f"  紫外老化10天: {len(all_stats[test_type]['紫外老化10天'])} 个含量")

            # 6. 10天老化 - 湿热
            raw_heat10 = process_folder(dir_10d, '湿热10天', lambda f: '湿热老化10天' in f)
            all_stats[test_type]['湿热老化10天'] = compute_stats(raw_heat10, test_type=test_type)
            print(f"  湿热老化10天: {len(all_stats[test_type]['湿热老化10天'])} 个含量")

    return all_stats


# ========== 分组柱状图 ==========
def make_grouped_bar_chart(stats_dict, test_type, output_path, ylabel, y_unit,
                           aging_name, group_keys, group_labels, is_elongation=False):
    """
    生成分组柱状图
    group_keys: list of stats keys e.g. ['初始组', '紫外老化5天', '紫外老化10天']
    group_labels: list of display labels e.g. ['初始', '5天紫外老化', '10天紫外老化']
    is_elongation: True=使用延伸率(%), False=使用强度(N)
    """
    fig, ax = plt.subplots(figsize=(10, 6))

    n_groups = len(SIO2_ORDER)
    n_bars = len(group_keys)
    group_width = 0.65
    bar_width = group_width / n_bars

    x = np.arange(n_groups)

    # 黑白打印友好样式：灰度渐变 + 不同填充图案
    bar_colors = ['0.85', '0.50', '0.15']     # 浅灰、中灰、深灰
    bar_hatches = ['', '///', '...']           # 无、斜线、点状
    bar_edges = ['black', 'black', 'black']

    for i, (key, label) in enumerate(zip(group_keys, group_labels)):
        if key not in stats_dict[test_type]:
            print(f"  警告: {key} 不在统计数据中")
            continue

        stats_data = stats_dict[test_type][key]
        means = []
        stds = []
        for sio2 in SIO2_ORDER:
            if sio2 in stats_data:
                if is_elongation:
                    means.append(stats_data[sio2]['max_elongation_mean'])
                    stds.append(stats_data[sio2]['max_elongation_std'])
                else:
                    # 将力(N)转换为应力(MPa)
                    if test_type == '拉伸':
                        conv = 1.0 / 40.0  # σ=F/A, A=40mm²
                    else:  # 弯曲
                        conv = (3.0 * 64.0) / (2.0 * 10.0 * 4.0**2)  # σ=3FL/(2bh²)=0.6
                    means.append(stats_data[sio2]['max_load_mean'] * conv)
                    stds.append(stats_data[sio2]['max_load_std'] * conv)
            else:
                means.append(0)
                stds.append(0)

        offset = (i - (n_bars - 1) / 2) * bar_width
        bars = ax.bar(x + offset, means, bar_width,
                     yerr=stds, capsize=4,
                     facecolor=bar_colors[i],
                     edgecolor=bar_edges[i],
                     linewidth=1.0,
                     hatch=bar_hatches[i],
                     label=label,
                     error_kw={'elinewidth': 1.0, 'capthick': 1.0, 'ecolor': 'black'})

        # 数值标注
        for j, (mean, std) in enumerate(zip(means, stds)):
            if mean > 0:
                y_max_all = max(means) if max(means) > 0 else 1
                y_pos = mean + std + y_max_all * 0.02
                ax.text(x[j] + offset, y_pos, f'{mean:.1f}',
                       ha='center', va='bottom', fontsize=14)

    ax.set_xticks(x)
    ax.set_xticklabels([f'SiO\u2082 {k}' for k in SIO2_ORDER], fontsize=14)
    ax.set_ylabel(f'{ylabel} ({y_unit})', fontsize=14)
    ax.set_xlabel('纳米SiO\u2082含量', fontsize=14)
    ax.set_title(f'{aging_name}PLA/竹粉/纳米SiO\u2082复合材料的{test_type}{ylabel}',
                fontsize=14, fontweight='bold', pad=15)
    ax.legend(fontsize=14, loc='best', framealpha=0.9)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # 自动调整y轴上限
    all_means = []
    all_stds = []
    for key in group_keys:
        if key in stats_dict[test_type]:
            for sio2 in SIO2_ORDER:
                if sio2 in stats_dict[test_type][key]:
                    if is_elongation:
                        all_means.append(stats_dict[test_type][key][sio2]['max_elongation_mean'])
                        all_stds.append(stats_dict[test_type][key][sio2]['max_elongation_std'])
                    else:
                        conv = 1.0/40.0 if test_type == '拉伸' else (3.0*64.0)/(2.0*10.0*4.0**2)
                        all_means.append(stats_dict[test_type][key][sio2]['max_load_mean'] * conv)
                        all_stds.append(stats_dict[test_type][key][sio2]['max_load_std'] * conv)
    if all_means:
        y_max_val = max(m + s for m, s in zip(all_means, all_stds))
        if y_max_val > 0:
            ax.set_ylim(bottom=0, top=y_max_val * 1.22)

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close()


def make_grouped_retention_chart(stats_dict, test_type, output_path,
                                  aging_keys, aging_labels, aging_name,
                                  is_elongation=False):
    """
    生成分组保留率柱状图
    每个SiO2含量处有多个柱（5天保留率、10天保留率）
    """
    fig, ax = plt.subplots(figsize=(10, 6))

    init_data = stats_dict[test_type].get('初始组', {})
    n_groups = len(SIO2_ORDER)
    n_bars = len(aging_keys)
    group_width = 0.65
    bar_width = group_width / n_bars

    x = np.arange(n_groups)
    hatches = ['', '///', '...']

    valid_data_found = False

    for i, (key, label) in enumerate(zip(aging_keys, aging_labels)):
        if key not in stats_dict[test_type]:
            continue
        aged_data = stats_dict[test_type][key]

        retentions = []
        ret_stds = []
        for sio2 in SIO2_ORDER:
            if sio2 not in aged_data or sio2 not in init_data:
                retentions.append(0)
                ret_stds.append(0)
                continue
            if is_elongation:
                init_val = init_data[sio2]['max_elongation_mean']
                aged_mean = aged_data[sio2]['max_elongation_mean']
                aged_std = aged_data[sio2]['max_elongation_std']
            else:
                init_val = init_data[sio2]['max_load_mean']
                aged_mean = aged_data[sio2]['max_load_mean']
                aged_std = aged_data[sio2]['max_load_std']

            if init_val > 0:
                retentions.append(aged_mean / init_val * 100)
                ret_stds.append(aged_std / init_val * 100)
            else:
                retentions.append(0)
                ret_stds.append(0)

        if any(r > 0 for r in retentions):
            valid_data_found = True

        # 灰度渐变样式（保留率图只有2组：5天和10天）
        ret_colors = ['0.50', '0.15']
        ci = min(i, len(ret_colors) - 1)

        offset = (i - (n_bars - 1) / 2) * bar_width
        bars = ax.bar(x + offset, retentions, bar_width,
                     yerr=ret_stds, capsize=4,
                     facecolor=ret_colors[ci],
                     edgecolor='black',
                     linewidth=1.0,
                     hatch=hatches[i],
                     label=label,
                     error_kw={'elinewidth': 1.0, 'capthick': 1.0, 'ecolor': 'black'})

        for j, (ret, rs) in enumerate(zip(retentions, ret_stds)):
            if ret > 0:
                ax.text(x[j] + offset, ret + rs + 0.5, f'{ret:.1f}%',
                       ha='center', va='bottom', fontsize=14)

    if not valid_data_found:
        plt.close()
        return

    ax.axhline(y=100, color='black', linestyle='--', linewidth=1, alpha=0.5)

    ax.set_xticks(x)
    ax.set_xticklabels([f'SiO\u2082 {k}' for k in SIO2_ORDER], fontsize=14)
    ylabel = '延伸率保留率' if is_elongation else '强度保留率'
    ax.set_ylabel(f'{ylabel} (%)', fontsize=14)
    ax.set_xlabel('纳米SiO\u2082含量', fontsize=14)
    ax.set_title(f'{aging_name}PLA/竹粉/纳米SiO\u2082复合材料的{test_type}{ylabel}',
                fontsize=14, fontweight='bold', pad=15)
    ax.legend(fontsize=14, loc='best', framealpha=0.9)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_ylim(bottom=0, top=115)

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight')
    plt.close()


# ========== 主程序 ==========
if __name__ == '__main__':
    base_dir = '/Users/shiberlin/Desktop/毕业论文'
    output_dir = os.path.join(base_dir, 'charts', 'thesis')
    os.makedirs(output_dir, exist_ok=True)

    # 1. 处理所有数据
    print("=" * 60)
    print("处理数据...")
    print("=" * 60)
    all_stats = process_all_data(base_dir)

    # 保存统计量
    stats_path = os.path.join(output_dir, 'stats_v2.json')
    with open(stats_path, 'w', encoding='utf-8') as f:
        json.dump(all_stats, f, ensure_ascii=False, indent=2)
    print(f"\n统计数据已保存到: {stats_path}")

    # 打印10天数据摘要
    print("\n" + "=" * 60)
    print("10天老化数据摘要")
    print("=" * 60)
    for tt in ['拉伸', '弯曲']:
        for key in ['紫外老化10天', '湿热老化10天']:
            if key in all_stats[tt]:
                print(f"\n{tt} - {key}:")
                for sio2 in SIO2_ORDER:
                    if sio2 in all_stats[tt][key]:
                        d = all_stats[tt][key][sio2]
                        conv = 1.0/40.0 if tt == '拉伸' else (3.0*64.0)/(2.0*10.0*4.0**2)
                        print(f"  {sio2}: 强度={d['max_load_mean']*conv:.2f}±{d['max_load_std']*conv:.2f} MPa, "
                              f"延伸率={d['max_elongation_mean']:.2f}±{d['max_elongation_std']:.2f}% (n={d['n_samples']})")

    # 2. 生成分组柱状图
    print("\n" + "=" * 60)
    print("生成分组柱状图...")
    print("=" * 60)

    ch4_configs = [
        ('紫外老化', ['初始组', '紫外老化5天', '紫外老化10天'],
                      ['初始', '5天紫外老化', '10天紫外老化']),
    ]
    ch5_configs = [
        ('湿热老化', ['初始组', '湿热老化5天', '湿热老化10天'],
                      ['初始', '5天湿热老化', '10天湿热老化']),
    ]

    for chapter, configs in [('第四章', ch4_configs), ('第五章', ch5_configs)]:
        for aging_name, group_keys, group_labels in configs:
            for test_type in ['拉伸', '弯曲']:
                for is_elong, metric_label, metric_unit in [
                    (False, '强度', 'MPa'),
                    (True, '延伸率', '%')
                ]:
                    suffix = '延伸率' if is_elong else '强度'
                    fname = f'{chapter}_{test_type}_{suffix}_分组柱状图.png'
                    fpath = os.path.join(output_dir, fname)
                    make_grouped_bar_chart(all_stats, test_type, fpath,
                                          metric_label, metric_unit, aging_name,
                                          group_keys, group_labels, is_elong)
                    print(f"  已保存: {fname}")

            # 保留率图（5天 vs 10天）
            aging_keys_ret = [k for k in group_keys if k != '初始组']
            aging_labels_ret = [l for k, l in zip(group_keys, group_labels) if k != '初始组']
            for test_type in ['拉伸', '弯曲']:
                for is_elong in [False, True]:
                    suffix = '延伸率保留率' if is_elong else '强度保留率'
                    fname = f'{chapter}_{test_type}_{suffix}_分组.png'
                    fpath = os.path.join(output_dir, fname)
                    make_grouped_retention_chart(all_stats, test_type, fpath,
                                                 aging_keys_ret, aging_labels_ret,
                                                 aging_name, is_elong)
                    print(f"  已保存: {fname}")

    # 统计生成图表数
    chart_files = [f for f in os.listdir(output_dir) if f.endswith('.png') and '分组' in f]
    print(f"\n共生成 {len(chart_files)} 张分组图表")
    print(f"输出目录: {output_dir}")
