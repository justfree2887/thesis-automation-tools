#!/usr/bin/env python3
"""
按老化条件处理拉伸和弯曲试验数据并生成图表
- 4种老化条件：初始组、紫外老化组、湿热老化组、紫外湿热耦合老化组
- 读取原始数据，剔除异常值（3σ准则）
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

def extract_aging_condition(filename):
    """
    从文件名中提取老化条件
    返回：'初始组', '紫外老化组', '湿热老化组', '紫外湿热耦合老化组', 或 None
    """
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
    """
    从文件名中提取SiO2含量
    返回：'0%', '1%', '3%', '5%', 或 None
    """
    for pct in ['0%', '1%', '3%', '5%']:
        if pct in filename:
            return pct
    return None

def read_excel_data(filepath):
    """
    读取Excel文件，返回处理后的DataFrame
    """
    try:
        df = pd.read_excel(filepath)
        # 归一化位移
        df['PositionValue'] = df['PositionValue'] - df['PositionValue'].iloc[0]
        return df
    except Exception as e:
        print(f"  读取文件失败 {os.path.basename(filepath)}: {e}")
        return None

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

def process_test_type(base_dir, test_type):
    """
    处理指定试验类型（拉伸或弯曲）的所有数据
    返回：{老化条件: {sio2含量: {'dataframes': [], 'max_loads': [], 'max_disps': []}}}
    """
    test_dir = os.path.join(base_dir, test_type)
    
    if not os.path.exists(test_dir):
        print(f"警告: {test_dir} 不存在")
        return {}
    
    # 数据结构：{老化条件: {sio2含量: {'dataframes': [], 'max_loads': [], 'max_disps': []}}}
    result = {}
    
    # 遍历所有子文件夹
    for folder_name in os.listdir(test_dir):
        folder_path = os.path.join(test_dir, folder_name)
        if not os.path.isdir(folder_path):
            continue
        
        print(f"  处理文件夹: {folder_name}")
        
        # 读取文件夹中的所有Excel文件
        for fname in os.listdir(folder_path):
            if not fname.endswith('.xlsx') or fname.startswith('.'):
                continue
            
            # 提取老化条件和SiO2含量
            aging_cond = extract_aging_condition(fname)
            sio2_pct = extract_sio2_content(fname)
            
            if aging_cond is None or sio2_pct is None:
                print(f"    跳过文件（无法识别条件）: {fname}")
                continue
            
            # 初始化数据结构
            if aging_cond not in result:
                result[aging_cond] = {}
            if sio2_pct not in result[aging_cond]:
                result[aging_cond][sio2_pct] = {'dataframes': [], 'max_loads': [], 'max_disps': []}
            
            # 读取数据
            filepath = os.path.join(folder_path, fname)
            df = read_excel_data(filepath)
            if df is None:
                continue
            
            # 存储数据和统计量
            result[aging_cond][sio2_pct]['dataframes'].append(df)
            result[aging_cond][sio2_pct]['max_loads'].append(df['LoadValue'].max())
            result[aging_cond][sio2_pct]['max_disps'].append(df['PositionValue'].max())
    
    return result

def filter_and_aggregate(data_dict):
    """
    对每个老化条件和SiO2含量组进行异常值剔除，然后聚合数据
    返回：{老化条件: {sio2含量: {'x':, 'y_mean':, 'y_std':, 'max_load_mean':, ...}}}
    """
    result = {}
    
    for aging_cond, sio2_data in data_dict.items():
        result[aging_cond] = {}
        
        for sio2_pct, data in sio2_data.items():
            if len(data['max_loads']) == 0:
                continue
            
            print(f"    {aging_cond} - {sio2_pct}: {len(data['dataframes'])} 个样本")
            
            # 剔除最大载荷的异常值
            filtered_loads, removed_loads = remove_outliers_3sigma(data['max_loads'])
            if removed_loads:
                print(f"      剔除{len(removed_loads)}个异常样本 (载荷)")
            
            # 剔除最大位移的异常值
            filtered_disps, removed_disps = remove_outliers_3sigma(data['max_disps'])
            if removed_disps:
                print(f"      剔除{len(removed_disps)}个异常样本 (位移)")
            
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

# ========== 生成试验曲线图（含误差带） ==========
def make_curve_chart(data_dict, aging_cond, test_type, output_path):
    """
    生成试验曲线图，包含误差带
    data_dict: {sio2含量: {'x':, 'y_mean':, 'y_std':}}
    """
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
    
    if not has_data:
        plt.close()
        return
    
    ax.set_xlabel('位移 (mm)', fontsize=10.5)
    ax.set_ylabel('载荷 (N)', fontsize=10.5)
    ax.set_title(f'{aging_cond}PLA/竹粉/纳米SiO₂复合材料的{test_type}试验曲线', fontsize=12, fontweight='bold', pad=15)
    ax.legend(fontsize=9, loc='best', framealpha=0.9)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"已保存: {output_path}")

# ========== 生成柱状图（含误差棒） ==========
def make_bar_chart(data_dict, aging_cond, test_type, output_path, ylabel, y_unit, is_displacement=False):
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
    
    # 检查是否有数据
    if all(m == 0 for m in means):
        plt.close()
        return
    
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
    ax.set_title(f'{aging_cond}PLA/竹粉/纳米SiO₂复合材料的{test_type}{ylabel}', fontsize=12, fontweight='bold', pad=15)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    # 调整Y轴范围
    y_max_val = max([m + s for m, s in zip(means, stds)]) if means else 0
    if y_max_val > 0:
        ax.set_ylim(bottom=0, top=y_max_val * 1.18)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"已保存: {output_path}")

# ========== 主程序 ==========
if __name__ == '__main__':
    # 确保输出目录存在
    os.makedirs('charts', exist_ok=True)
    
    # 使用绝对路径
    base_dir = '/Users/shiberlin/Desktop/毕业论文'
    test_types = ['拉伸', '弯曲']
    
    for test_type in test_types:
        print(f"\n{'='*60}")
        print(f"处理{test_type}试验数据")
        print('='*60)
        
        # 处理该试验类型的所有数据
        raw_data = process_test_type(base_dir, test_type)
        
        if not raw_data:
            print(f"  无数据")
            continue
        
        # 剔除异常值并聚合
        print(f"  剔除异常值...")
        processed_data = filter_and_aggregate(raw_data)
        
        if not processed_data:
            print(f"  处理后无有效数据")
            continue
        
        # 为每个老化条件生成图表
        for aging_cond in ['初始组', '紫外老化组', '湿热老化组', '紫外湿热耦合老化组']:
            if aging_cond not in processed_data:
                print(f"  {aging_cond}: 无数据")
                continue
            
            print(f"  生成{aging_cond}的图表...")
            
            # 生成试验曲线图
            make_curve_chart(
                processed_data[aging_cond],
                aging_cond,
                test_type,
                f'charts/{aging_cond}_{test_type}_试验曲线.png'
            )
            
            # 生成柱状图（最大载荷/强度）
            make_bar_chart(
                processed_data[aging_cond],
                aging_cond,
                test_type,
                f'charts/{aging_cond}_{test_type}_强度_柱状图.png',
                ylabel=f'{test_type}强度',
                y_unit='N',
                is_displacement=False
            )
            
            # 生成柱状图（最大位移）
            make_bar_chart(
                processed_data[aging_cond],
                aging_cond,
                test_type,
                f'charts/{aging_cond}_{test_type}_位移_柱状图.png',
                ylabel=f'{test_type}位移',
                y_unit='mm',
                is_displacement=True
            )
    
    print("\n" + "="*60)
    print("全部图表生成完成!")
    print("="*60)
