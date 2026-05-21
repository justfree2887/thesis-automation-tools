#!/usr/bin/env python3
"""
拉伸和弯曲试验数据处理与可视化
- 读取初始、5天、10天的原始数据
- 剔除异常数据（3σ准则）
- 生成试验曲线图（含误差带）
- 生成实验结果柱状图（含误差棒）
- 使用黑体5号字体
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os
import re

# ========== 中文字体配置 ==========
plt.rcParams['font.sans-serif'] = ['SimHei', 'Heiti SC', 'Arial Unicode MS', 'STHeiti', 'Hiragino Sans GB', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 10.5  # 5号字约等于10.5pt

# ========== 配置 ==========
colors_sio2 = {
    '0%':  '#888780',
    '1%':  '#185FA5',
    '3%':  '#1D9E75',
    '5%':  '#D85A30',
}
marker_styles = {
    '0%':  's',
    '1%':  'o',
    '3%':  '^',
    '5%':  'D',
}

# ========== 异常值剔除函数 ==========
def remove_outliers_3sigma(data_list, threshold=3):
    """
    使用3σ准则剔除异常值
    返回：(筛选后数据列表, 被剔除的索引列表)
    """
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

def find_folder(base_path, keyword):
    """
    在base_path目录下模糊查找包含keyword的文件夹
    返回：匹配的文件夹完整路径，如果没找到则返回None
    """
    if not os.path.exists(base_path):
        return None
    
    for fname in os.listdir(base_path):
        if keyword in fname:
            return os.path.join(base_path, fname)
    
    return None

def process_excel_files(folder):
    """
    读取指定文件夹中的所有Excel文件
    返回：{sio2含量: {'dataframes': [df1, df2, ...], 'max_loads': [...], 'max_disps': [...]}}
    """
    if not os.path.exists(folder):
        print(f"警告: {folder} 不存在")
        return {}
    
    result = {
        '0%': {'dataframes': [], 'max_loads': [], 'max_disps': []},
        '1%': {'dataframes': [], 'max_loads': [], 'max_disps': []},
        '3%': {'dataframes': [], 'max_loads': [], 'max_disps': []},
        '5%': {'dataframes': [], 'max_loads': [], 'max_disps': []}
    }
    
    files = [f for f in os.listdir(folder) if f.endswith('.xlsx') and not f.startswith('.')]
    print(f"  找到 {len(files)} 个xlsx文件")
    
    for fname in files:
        # 判断SiO2含量
        pct = None
        for p in ['0%', '1%', '3%', '5%']:
            if p in fname:
                pct = p
                break
        if pct is None:
            continue
        
        # 读取Excel
        try:
            df = pd.read_excel(os.path.join(folder, fname))
            # 归一化位移
            df['PositionValue'] = df['PositionValue'] - df['PositionValue'].iloc[0]
            result[pct]['dataframes'].append(df)
            result[pct]['max_loads'].append(df['LoadValue'].max())
            result[pct]['max_disps'].append(df['PositionValue'].max())
        except Exception as e:
            print(f"读取文件失败 {fname}: {e}")
    
    return result

def filter_and_aggregate(data_dict):
    """
    对每个SiO2含量组进行异常值剔除，然后聚合数据
    返回：{sio2含量: {'x':, 'y_mean':, 'y_std':, 'max_load_mean':, ...}}
    """
    result = {}
    
    for pct, data in data_dict.items():
        if len(data['max_loads']) == 0:
            continue
        
        # 剔除最大载荷的异常值
        filtered_loads, removed_loads = remove_outliers_3sigma(data['max_loads'])
        if removed_loads:
            print(f"  {pct}: 剔除{len(removed_loads)}个异常样本 (载荷)")
        
        # 剔除最大位移的异常值
        filtered_disps, removed_disps = remove_outliers_3sigma(data['max_disps'])
        if removed_disps:
            print(f"  {pct}: 剔除{len(removed_disps)}个异常样本 (位移)")
        
        # 合并被剔除的索引
        all_removed = set(removed_loads + removed_disps)
        valid_indices = [i for i in range(len(data['dataframes'])) if i not in all_removed]
        
        if len(valid_indices) == 0:
            continue
        
        # 插值到统一x轴，计算均值和标准差
        samples = [data['dataframes'][i] for i in valid_indices]
        
        # 拉伸/弯曲曲线：载荷-位移
        x_common, y_mean, y_std = interp_curves(samples)
        
        if x_common is None:
            continue
        
        # 计算最大载荷和位移的统计量
        valid_loads = [data['max_loads'][i] for i in valid_indices]
        valid_disps = [data['max_disps'][i] for i in valid_indices]
        
        result[pct] = {
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

def interp_curves(samples, n_pts=400):
    """
    将多个样本曲线插值到统一x轴
    返回：(x_common, y_mean, y_std)
    """
    if not samples:
        return None, None, None
    
    # 找到共同的x范围
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

# ========== 生成试验曲线图（含误差带） ==========
def make_curve_chart(data_dict, title, filename, xlabel='位移 (mm)', ylabel='载荷 (N)'):
    """
    生成试验曲线图，包含误差带
    data_dict: {sio2含量: {'x':, 'y_mean':, 'y_std':}}
    """
    fig, ax = plt.subplots(figsize=(9, 6))
    
    for pct in ['0%', '1%', '3%', '5%']:
        if pct not in data_dict:
            continue
        
        d = data_dict[pct]
        x = d['x']
        y_mean = d['y_mean']
        y_std = d['y_std']
        
        # 绘制均值曲线
        ax.plot(x, y_mean, 
                label=f'{pct} (n={d["n_samples"]})',
                color=colors_sio2[pct],
                linewidth=2,
                marker=marker_styles[pct],
                markersize=4,
                markevery=20)
        
        # 绘制误差带
        ax.fill_between(x, 
                        y_mean - y_std,
                        y_mean + y_std,
                        color=colors_sio2[pct],
                        alpha=0.2)
    
    ax.set_xlabel(xlabel, fontsize=10.5)
    ax.set_ylabel(ylabel, fontsize=10.5)
    ax.set_title(title, fontsize=12, fontweight='bold', pad=15)
    ax.legend(fontsize=9, loc='best', framealpha=0.9)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    plt.tight_layout()
    plt.savefig(filename, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"已保存: {filename}")

# ========== 生成柱状图（含误差棒） ==========
def make_bar_chart(data_dict, title, filename, ylabel, y_unit, is_displacement=False):
    """
    生成实验结果柱状图
    data_dict: {sio2含量: {'max_load_mean':, 'max_load_std':}} 或位移数据
    """
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
    
    bars = ax.bar(x_pos, means, bar_width,
                  yerr=stds,
                  capsize=5,
                  color=[colors_sio2[k] for k in sio2_keys],
                  alpha=0.85,
                  edgecolor='white',
                  linewidth=1.2,
                  error_kw={'elinewidth': 1.5, 'capthick': 1.5, 'ecolor': '#444444'})
    
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
    ax.set_title(title, fontsize=12, fontweight='bold', pad=15)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    # 调整Y轴范围
    y_max_val = max([m + s for m, s in zip(means, stds)]) if means else 0
    if y_max_val > 0:
        ax.set_ylim(bottom=0, top=y_max_val * 1.18)
    
    plt.tight_layout()
    plt.savefig(filename, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"已保存: {filename}")

# ========== 主程序 ==========
if __name__ == '__main__':
    # 确保输出目录存在
    os.makedirs('charts', exist_ok=True)
    
    # 使用绝对路径
    base_dir = '/Users/shiberlin/Desktop/毕业论文'
    time_point_keywords = ['初始', '5天', '10天']
    test_types = ['拉伸', '弯曲']
    
    for test_type in test_types:
        print(f"\n{'='*60}")
        print(f"处理{test_type}试验数据")
        print('='*60)
        
        for time_keyword in time_point_keywords:
            print(f"\n--- {time_keyword} ---")
            
            # 查找文件夹（模糊匹配）
            base_path = os.path.join(base_dir, test_type)
            folder = find_folder(base_path, time_keyword)
            
            if folder is None:
                print(f"  未找到包含'{time_keyword}'的文件夹")
                continue
            
            print(f"  找到文件夹: {folder}")
            
            # 读取原始数据
            raw_data = process_excel_files(folder)
            
            if not raw_data or all(len(v['dataframes']) == 0 for v in raw_data.values()):
                print(f"  无数据")
                continue
            
            # 剔除异常值并聚合
            print(f"  剔除异常值...")
            processed_data = filter_and_aggregate(raw_data)
            
            if not processed_data:
                print(f"  处理后无有效数据")
                continue
            
            # 生成试验曲线图
            print(f"  生成曲线图...")
            make_curve_chart(
                processed_data,
                f'{time_keyword}条件下PLA/竹粉/纳米SiO₂复合材料的{test_type}试验曲线',
                f'charts/{time_keyword}_{test_type}_试验曲线.png',
                xlabel='位移 (mm)',
                ylabel='载荷 (N)'
            )
            
            # 生成柱状图（最大载荷/强度）
            print(f"  生成强度柱状图...")
            make_bar_chart(
                processed_data,
                f'{time_keyword}条件下PLA/竹粉/纳米SiO₂复合材料的{test_type}强度',
                f'charts/{time_keyword}_{test_type}_强度_柱状图.png',
                ylabel=f'{test_type}强度',
                y_unit='N',
                is_displacement=False
            )
            
            # 生成柱状图（最大位移）
            print(f"  生成位移柱状图...")
            make_bar_chart(
                processed_data,
                f'{time_keyword}条件下PLA/竹粉/纳米SiO₂复合材料的{test_type}位移',
                f'charts/{time_keyword}_{test_type}_位移_柱状图.png',
                ylabel=f'{test_type}位移',
                y_unit='mm',
                is_displacement=True
            )
    
    print("\n" + "="*60)
    print("全部图表生成完成!")
    print("="*60)
