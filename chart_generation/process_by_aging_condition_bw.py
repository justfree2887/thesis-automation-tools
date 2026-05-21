#!/usr/bin/env python3
"""
按老化条件处理拉伸和弯曲试验数据并生成图表（黑白打印机友好版）
- 使用灰度、标记形状、线型、填充图案来区分数据
- 适合黑白打印机打印
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd
import os

# ========== 中文字体配置 ==========
plt.rcParams['font.sans-serif'] = ['SimHei', 'Heiti SC', 'Arial Unicode MS', 'STHeiti', 'Hiragino Sans GB', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 10.5  # 5号字约等于10.5pt

# ========== 黑白打印机友好的配置 ==========
# 使用灰度值（0=黑，1=白）
grayscale = {
    '0%': 0.0,    # 纯黑
    '1%': 0.25,   # 深灰
    '3%': 0.55,   # 中灰
    '5%': 0.15,   # 深灰（原0.85太白，黑白打印无法分辨）
}

# 标记形状
marker_styles = {
    '0%': 's',     # 正方形
    '1%': 'o',     # 圆形
    '3%': '^',     # 上三角形
    '5%': 'D',     # 菱形
}

# 线型
line_styles = {
    '0%': '-',     # 实线
    '1%': '--',    # 虚线
    '3%': '-.',    # 点划线
    '5%': (0, (2, 1, 0.5, 1)),  # 密疏点线（比单纯点线更清晰）
}

# 柱状图填充图案
hatch_patterns = {
    '0%': '',      # 无填充
    '1%': '///',   # 斜线
    '3%': 'xxx',   # 叉号
    '5%': '...',   # 点
}

# ========== 异常值剔除函数 ==========
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
    for pct in ['0%', '1%', '3%', '5%']:
        if pct in filename:
            return pct
    return None

def read_excel_data(filepath):
    try:
        df = pd.read_excel(filepath)
        df['PositionValue'] = df['PositionValue'] - df['PositionValue'].iloc[0]
        return df
    except Exception as e:
        print(f"  读取文件失败 {os.path.basename(filepath)}: {e}")
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

def process_test_type(base_dir, test_type):
    test_dir = os.path.join(base_dir, test_type)
    if not os.path.exists(test_dir):
        print(f"警告: {test_dir} 不存在")
        return {}
    
    result = {}
    
    for folder_name in os.listdir(test_dir):
        folder_path = os.path.join(test_dir, folder_name)
        if not os.path.isdir(folder_path):
            continue
        
        print(f"  处理文件夹: {folder_name}")
        
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
    result = {}
    
    for aging_cond, sio2_data in data_dict.items():
        result[aging_cond] = {}
        
        for sio2_pct, data in sio2_data.items():
            if len(data['max_loads']) == 0:
                continue
            
            print(f"    {aging_cond} - {sio2_pct}: {len(data['dataframes'])} 个样本")
            
            filtered_loads, removed_loads = remove_outliers_3sigma(data['max_loads'])
            if removed_loads:
                print(f"      剔除{len(removed_loads)}个异常样本 (载荷)")
            
            filtered_disps, removed_disps = remove_outliers_3sigma(data['max_disps'])
            if removed_disps:
                print(f"      剔除{len(removed_disps)}个异常样本 (位移)")
            
            all_removed = set(removed_loads + removed_disps)
            valid_indices = [i for i in range(len(data['dataframes'])) if i not in all_removed]
            
            if len(valid_indices) == 0:
                continue
            
            samples = [data['dataframes'][i] for i in valid_indices]
            x_common, y_mean, y_std = interp_curves(samples)
            
            if x_common is None:
                continue
            
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

# ========== 生成试验曲线图（黑白打印机友好） ==========
def make_curve_chart_bw(data_dict, aging_cond, test_type, output_path):
    fig, ax = plt.subplots(figsize=(9, 6))
    
    has_data = False
    for pct in ['0%', '1%', '3%', '5%']:
        if pct not in data_dict:
            continue
        
        d = data_dict[pct]
        x = d['x']
        y_mean = d['y_mean']
        y_std = d['y_std']
        
        has_data = True
        gray_val = grayscale[pct]
        color = str(gray_val)
        
        # 绘制均值曲线
        ax.plot(x, y_mean, 
                label=f'{pct} (n={d["n_samples"]})',
                color=color,
                linewidth=2,
                linestyle=line_styles[pct],
                marker=marker_styles[pct],
                markersize=4,
                markevery=20,
                markerfacecolor='white',
                markeredgecolor=color)
        
        # 绘制误差带（用不同透明度区分）
        alpha_val = 0.15 + gray_val * 0.25
        ax.fill_between(x, 
                        y_mean - y_std,
                        y_mean + y_std,
                        color=color,
                        alpha=alpha_val)
    
    if not has_data:
        plt.close()
        return
    
    ax.set_xlabel('位移 (mm)', fontsize=10.5)
    ax.set_ylabel('载荷 (N)', fontsize=10.5)
    ax.set_title(f'{aging_cond}PLA/竹粉/纳米SiO₂复合材料的{test_type}试验曲线', fontsize=12, fontweight='bold', pad=15)
    
    # 图例
    legend_handles = []
    for pct in ['0%', '1%', '3%', '5%']:
        if pct in data_dict:
            gray_val = grayscale[pct]
            color = str(gray_val)
            handle = plt.Line2D([0], [0], 
                               color=color, 
                               linewidth=2,
                               linestyle=line_styles[pct],
                               marker=marker_styles[pct],
                               markersize=6,
                               markerfacecolor='white',
                               markeredgecolor=color,
                               label=f'{pct} (n={data_dict[pct]["n_samples"]})')
            legend_handles.append(handle)
    
    ax.legend(handles=legend_handles, fontsize=9, loc='best', framealpha=0.9)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"已保存: {output_path}")

# ========== 生成柱状图（黑白打印机友好） ==========
def make_bar_chart_bw(data_dict, aging_cond, test_type, output_path, ylabel, y_unit, is_displacement=False):
    fig, ax = plt.subplots(figsize=(8, 5.5))
    
    sio2_keys = ['0%', '1%', '3%', '5%']
    x_pos = np.arange(len(sio2_keys))
    bar_width = 0.6
    
    means = []
    stds = []
    
    for key in sio2_keys:
        if key in data_dict:
            if is_displacement:
                means.append(data_dict[key]['max_disp_mean'])
                stds.append(data_dict[key]['max_disp_std'])
            else:
                means.append(data_dict[key]['max_load_mean'])
                stds.append(data_dict[key]['max_load_std'])
        else:
            means.append(0)
            stds.append(0)
    
    if all(m == 0 for m in means):
        plt.close()
        return
    
    # 绘制柱状图
    bars = ax.bar(x_pos, means, bar_width,
                  yerr=stds,
                  capsize=5,
                  color=[str(grayscale[k]) for k in sio2_keys],
                  edgecolor='black',
                  linewidth=1.2,
                  error_kw={'elinewidth': 1.5, 'capthick': 1.5, 'ecolor': 'black'})
    
    # 添加填充图案
    for bar, key in zip(bars, sio2_keys):
        if key in data_dict and data_dict[key].get('n_samples', 0) > 0:
            bar.set_hatch(hatch_patterns[key])
    
    # 在柱子上显示数值
    for i, (mean, std) in enumerate(zip(means, stds)):
        if mean > 0:
            y_max = max(means) if max(means) > 0 else 1
            y_pos = mean + std + y_max * 0.03
            ax.text(i, y_pos, f'{mean:.1f}', 
                   ha='center', va='bottom', fontsize=9)
    
    ax.set_xticks(x_pos)
    ax.set_xticklabels(sio2_keys, fontsize=10.5)
    ax.set_ylabel(f'{ylabel} ({y_unit})', fontsize=10.5)
    ax.set_xlabel('纳米SiO₂含量', fontsize=10.5)
    ax.set_title(f'{aging_cond}PLA/竹粉/纳米SiO₂复合材料的{test_type}{ylabel}', fontsize=12, fontweight='bold', pad=15)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    y_max_val = max([m + s for m, s in zip(means, stds)]) if means else 0
    if y_max_val > 0:
        ax.set_ylim(bottom=0, top=y_max_val * 1.18)
    
    # 图例
    legend_handles = []
    for key in sio2_keys:
        if key in data_dict and data_dict[key].get('n_samples', 0) > 0:
            handle = mpatches.Patch(
                facecolor=str(grayscale[key]),
                edgecolor='black',
                hatch=hatch_patterns[key],
                label=f'{key} (n={data_dict[key]["n_samples"]})'
            )
            legend_handles.append(handle)
    
    if legend_handles:
        ax.legend(handles=legend_handles, fontsize=9, loc='best', framealpha=0.9)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"已保存: {output_path}")

# ========== 主程序 ==========
if __name__ == '__main__':
    os.makedirs('charts', exist_ok=True)
    
    base_dir = '/Users/shiberlin/Desktop/毕业论文'
    test_types = ['拉伸', '弯曲']
    
    for test_type in test_types:
        print(f"\n{'='*60}")
        print(f"处理{test_type}试验数据（黑白打印机友好版）")
        print('='*60)
        
        raw_data = process_test_type(base_dir, test_type)
        
        if not raw_data:
            print(f"  无数据")
            continue
        
        print(f"  剔除异常值...")
        processed_data = filter_and_aggregate(raw_data)
        
        if not processed_data:
            print(f"  处理后无有效数据")
            continue
        
        for aging_cond in ['初始组', '紫外老化组', '湿热老化组', '紫外湿热耦合老化组']:
            if aging_cond not in processed_data:
                print(f"  {aging_cond}: 无数据")
                continue
            
            print(f"  生成{aging_cond}的图表...")
            
            # 生成试验曲线图
            make_curve_chart_bw(
                processed_data[aging_cond],
                aging_cond,
                test_type,
                f'charts/{aging_cond}_{test_type}_试验曲线_黑白版.png'
            )
            
            # 生成柱状图（强度）
            make_bar_chart_bw(
                processed_data[aging_cond],
                aging_cond,
                test_type,
                f'charts/{aging_cond}_{test_type}_强度_柱状图_黑白版.png',
                ylabel=f'{test_type}强度',
                y_unit='N',
                is_displacement=False
            )
            
            # 生成柱状图（位移）
            make_bar_chart_bw(
                processed_data[aging_cond],
                aging_cond,
                test_type,
                f'charts/{aging_cond}_{test_type}_位移_柱状图_黑白版.png',
                ylabel=f'{test_type}位移',
                y_unit='mm',
                is_displacement=True
            )
    
    print("\n" + "="*60)
    print("全部图表生成完成（黑白打印机友好版）!")
    print("="*60)
    print("\n区分方式：")
    print("  曲线图：不同线型(实线/虚线/点划线/点线) + 不同标记(方/圆/三角/菱形) + 不同灰度")
    print("  柱状图：不同灰度 + 不同填充图案(无/斜线/叉号/点)")
