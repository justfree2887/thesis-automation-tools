#!/usr/bin/env python3
"""
对比紫外老化和湿热老化在不同时间点的力学性能
- 对比：紫外老化组 vs 湿热老化组
- 时间点：初始组、5天老化组、10天老化组
- 按SiO2含量分别制图（4个子图）
- 黑白打印机友好
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
plt.rcParams['font.size'] = 10.5

# ========== 黑白打印机友好的配置 ==========
# 老化条件：紫外=深色/实填充，湿热=浅色/图案填充
aging_colors = {
    '紫外': '0.0',      # 纯黑
    '湿热': '0.7',      # 浅灰
}
aging_hatch = {
    '紫外': '',         # 无填充
    '湿热': '///',      # 斜线
}

# 时间点
time_points_uv = ['初始组', '5天老化', '10天老化']
time_points_heat = ['初始组', '5天老化', '10天老化']
time_labels = ['初始', '5天', '10天']

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
        return None

def process_test_type(base_dir, test_type):
    """处理指定试验类型的所有数据，按老化条件和SiO2含量分组"""
    test_dir = os.path.join(base_dir, test_type)
    if not os.path.exists(test_dir):
        return {}
    
    result = {}  # {老化条件: {SiO2含量: {'max_loads': [], 'max_disps': []}}}
    
    for folder_name in os.listdir(test_dir):
        folder_path = os.path.join(test_dir, folder_name)
        if not os.path.isdir(folder_path):
            continue
        
        for fname in os.listdir(folder_path):
            if not fname.endswith('.xlsx') or fname.startswith('.'):
                continue
            
            aging_cond = extract_aging_condition(fname)
            sio2_pct = extract_sio2_content(fname)
            
            if aging_cond is None or sio2_pct is None:
                continue
            
            # 只处理初始组、紫外老化组、湿热老化组
            if aging_cond not in ['初始组', '紫外老化组', '湿热老化组']:
                continue
            
            if aging_cond not in result:
                result[aging_cond] = {}
            if sio2_pct not in result[aging_cond]:
                result[aging_cond][sio2_pct] = {'max_loads': [], 'max_disps': []}
            
            filepath = os.path.join(folder_path, fname)
            df = read_excel_data(filepath)
            if df is None:
                continue
            
            result[aging_cond][sio2_pct]['max_loads'].append(df['LoadValue'].max())
            result[aging_cond][sio2_pct]['max_disps'].append(df['PositionValue'].max())
    
    return result

def filter_and_aggregate(raw_data):
    """剔除异常值并聚合"""
    result = {}
    
    for aging_cond, sio2_data in raw_data.items():
        result[aging_cond] = {}
        
        for sio2_pct, data in sio2_data.items():
            if len(data['max_loads']) == 0:
                continue
            
            # 剔除异常值
            filtered_loads, removed_loads = remove_outliers_3sigma(data['max_loads'])
            filtered_disps, removed_disps = remove_outliers_3sigma(data['max_disps'])
            
            if removed_loads:
                print(f"      {aging_cond}-{sio2_pct}: 剔除{len(removed_loads)}个异常值(载荷)")
            
            # 使用筛选后的数据
            valid_indices = [i for i in range(len(data['max_loads'])) if i not in set(removed_loads + removed_disps)]
            
            if len(valid_indices) == 0:
                continue
            
            valid_loads = [data['max_loads'][i] for i in valid_indices]
            valid_disps = [data['max_disps'][i] for i in valid_indices]
            
            result[aging_cond][sio2_pct] = {
                'load_mean': np.mean(valid_loads),
                'load_std': np.std(valid_loads, ddof=1) if len(valid_loads) > 1 else 0,
                'disp_mean': np.mean(valid_disps),
                'disp_std': np.std(valid_disps, ddof=1) if len(valid_disps) > 1 else 0,
                'n': len(valid_loads)
            }
    
    return result

# ========== 生成对比柱状图（黑白打印机友好） ==========
def make_comparison_chart(data_dict, test_type, output_path, ylabel, y_unit, is_displacement=False):
    """
    生成紫外 vs 湿热在不同时间点的对比柱状图
    4个子图：每个SiO2含量一个
    """
    sio2_keys = ['0%', '1%', '3%', '5%']
    
    # 创建4个子图
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    axes = axes.flatten()
    
    for idx, sio2_pct in enumerate(sio2_keys):
        ax = axes[idx]
        
        # 准备数据
        x = np.arange(len(time_labels))
        bar_width = 0.35
        
        means_uv = []
        stds_uv = []
        means_heat = []
        stds_heat = []
        
        for tp in ['初始组', '5天老化', '10天老化']:
            # 紫外老化组数据
            if tp in data_dict and sio2_pct in data_dict[tp]:
                if is_displacement:
                    means_uv.append(data_dict[tp][sio2_pct]['disp_mean'])
                    stds_uv.append(data_dict[tp][sio2_pct]['disp_std'])
                else:
                    means_uv.append(data_dict[tp][sio2_pct]['load_mean'])
                    stds_uv.append(data_dict[tp][sio2_pct]['load_std'])
            else:
                means_uv.append(0)
                stds_uv.append(0)
            
            # 湿热老化组数据
            if tp in data_dict and sio2_pct in data_dict[tp]:
                if is_displacement:
                    means_heat.append(data_dict[tp][sio2_pct]['disp_mean'])
                    stds_heat.append(data_dict[tp][sio2_pct]['disp_std'])
                else:
                    means_heat.append(data_dict[tp][sio2_pct]['load_mean'])
                    stds_heat.append(data_dict[tp][sio2_pct]['load_std'])
            else:
                means_heat.append(0)
                stds_heat.append(0)
        
        # 绘制紫外老化组柱状图
        bars1 = ax.bar(x - bar_width/2, means_uv, bar_width,
                       yerr=stds_uv,
                       capsize=4,
                       color=aging_colors['紫外'],
                       edgecolor='black',
                       linewidth=1.2,
                       hatch=aging_hatch['紫外'],
                       error_kw={'elinewidth': 1.2, 'capthick': 1.2, 'ecolor': 'black'})
        
        # 绘制湿热老化组柱状图
        bars2 = ax.bar(x + bar_width/2, means_heat, bar_width,
                       yerr=stds_heat,
                       capsize=4,
                       color=aging_colors['湿热'],
                       edgecolor='black',
                       linewidth=1.2,
                       hatch=aging_hatch['湿热'],
                       error_kw={'elinewidth': 1.2, 'capthick': 1.2, 'ecolor': 'black'})
        
        # 在柱子上显示数值（仅显示非零值）
        for i, (mean, std) in enumerate(zip(means_uv, stds_uv)):
            if mean > 0:
                y_max = max(max(means_uv), max(means_heat)) if (means_uv or means_heat) else 1
                y_pos = mean + std + y_max * 0.02
                ax.text(i - bar_width/2, y_pos, f'{mean:.1f}', 
                       ha='center', va='bottom', fontsize=7)
        
        for i, (mean, std) in enumerate(zip(means_heat, stds_heat)):
            if mean > 0:
                y_max = max(max(means_uv), max(means_heat)) if (means_uv or means_heat) else 1
                y_pos = mean + std + y_max * 0.02
                ax.text(i + bar_width/2, y_pos, f'{mean:.1f}', 
                       ha='center', va='bottom', fontsize=7)
        
        ax.set_xticks(x)
        ax.set_xticklabels(time_labels, fontsize=9)
        ax.set_title(f'SiO₂ {sio2_pct}', fontsize=10, fontweight='bold')
        ax.grid(True, axis='y', alpha=0.3, linestyle='--')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        # Y轴范围
        all_means = means_uv + means_heat
        all_stds = stds_uv + stds_heat
        y_max_val = max([m + s for m, s in zip(all_means, all_stds)]) if all_means else 0
        if y_max_val > 0:
            ax.set_ylim(bottom=0, top=y_max_val * 1.2)
    
    # 总标题
    fig.suptitle(f'{test_type}性能对比：紫外老化 vs 湿热老化', fontsize=14, fontweight='bold', y=0.98)
    
    # 统一的Y轴标签
    fig.text(0.02, 0.5, f'{ylabel} ({y_unit})', va='center', rotation='vertical', fontsize=10.5)
    fig.text(0.5, 0.02, '老化时间', ha='center', fontsize=10.5)
    
    # 统一图例
    legend_handles = [
        mpatches.Patch(facecolor=aging_colors['紫外'], edgecolor='black', hatch=aging_hatch['紫外'], label='紫外老化'),
        mpatches.Patch(facecolor=aging_colors['湿热'], edgecolor='black', hatch=aging_hatch['湿热'], label='湿热老化')
    ]
    fig.legend(handles=legend_handles, loc='upper right', fontsize=10, framealpha=0.9)
    
    plt.tight_layout(rect=[0.03, 0.03, 0.97, 0.96])
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
        print(f"处理{test_type}试验数据（紫外 vs 湿热对比）")
        print('='*60)
        
        # 处理数据
        raw_data = process_test_type(base_dir, test_type)
        
        if not raw_data:
            print(f"  无数据")
            continue
        
        print(f"  剔除异常值...")
        processed_data = filter_and_aggregate(raw_data)
        
        if not processed_data:
            print(f"  处理后无有效数据")
            continue
        
        # 生成对比柱状图（强度）
        print(f"  生成强度对比图...")
        make_comparison_chart(
            processed_data,
            test_type,
            f'charts/紫外_vs_湿热_{test_type}强度对比_黑白版.png',
            ylabel=f'{test_type}强度',
            y_unit='N',
            is_displacement=False
        )
        
        # 生成对比柱状图（位移）
        print(f"  生成位移对比图...")
        make_comparison_chart(
            processed_data,
            test_type,
            f'charts/紫外_vs_湿热_{test_type}位移对比_黑白版.png',
            ylabel=f'{test_type}位移',
            y_unit='mm',
            is_displacement=True
        )
    
    print("\n" + "="*60)
    print("对比图表生成完成!")
    print("="*60)
    print("\n图表说明：")
    print("  - 4个子图：分别对应SiO₂含量 0%、1%、3%、5%")
    print("  - X轴：初始、5天、10天")
    print("  - 分组柱状图：紫外老化 vs 湿热老化")
    print("  - 黑白打印区分：紫外=纯黑无填充，湿热=浅灰斜线填充")
