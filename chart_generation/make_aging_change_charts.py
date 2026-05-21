#!/usr/bin/env python3
"""
针对湿热老化和紫外老化，分别制作力学性能变化柱状图
- 湿热老化：初始、5天、10天
- 紫外老化：初始、5天、10天
- 4个子图：每个SiO2含量一个
- 黑白打印机友好
- 剔除异常数据（3σ准则）
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd
import os
import re

# ========== 中文字体配置 ==========
plt.rcParams['font.sans-serif'] = ['SimHei', 'Heiti SC', 'Arial Unicode MS', 'STHeiti', 'Hiragino Sans GB', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 10.5

# ========== 配置 ==========
# 时间点
time_points = ['初始', '5天', '10天']

# 黑白打印机友好配置
# 时间点用不同灰度和填充图案区分
time_colors = {
    '初始': '0.0',      # 纯黑
    '5天': '0.4',      # 中灰
    '10天': '0.8',     # 浅灰
}
time_hatch = {
    '初始': '',         # 无填充
    '5天': '///',      # 斜线
    '10天': '...',     # 点
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

def extract_info_from_filename(filename, folder_name):
    """
    从文件名和文件夹名中提取信息
    返回：(aging_type, aging_time, sio2_pct, sample_num)
    """
    # 从文件夹名确定aging_time
    if '5天' in folder_name:
        aging_time = '5天'
    elif '10天' in folder_name:
        aging_time = '10天'
    elif '初始' in folder_name:
        aging_time = '初始'
    else:
        aging_time = None
    
    # 从文件名确定aging_type
    if '紫外湿热' in filename:
        aging_type = '紫外湿热耦合'
    elif '紫外' in filename:
        aging_type = '紫外'
    elif '湿热' in filename:
        aging_type = '湿热'
    elif '初始' in filename:
        aging_type = '初始'
    else:
        aging_type = None
    
    # 提取SiO2含量
    sio2_pct = None
    for pct in ['0%', '1%', '3%', '5%']:
        if pct in filename:
            sio2_pct = pct
            break
    
    # 提取样本编号
    sample_num = None
    match = re.search(r'(\d+)', filename)
    if match:
        sample_num = int(match.group(1))
    
    return aging_type, aging_time, sio2_pct, sample_num

def read_excel_data(filepath):
    try:
        df = pd.read_excel(filepath)
        df['PositionValue'] = df['PositionValue'] - df['PositionValue'].iloc[0]
        return df
    except Exception as e:
        return None

def process_test_type(base_dir, test_type):
    """
    处理指定试验类型的所有数据
    返回：{aging_type: {aging_time: {sio2_pct: {'max_loads': [], 'max_disps': []}}}}
    """
    test_dir = os.path.join(base_dir, test_type)
    if not os.path.exists(test_dir):
        return {}
    
    result = {}
    
    for folder_name in os.listdir(test_dir):
        folder_path = os.path.join(test_dir, folder_name)
        if not os.path.isdir(folder_path):
            continue
        
        for fname in os.listdir(folder_path):
            if not fname.endswith('.xlsx') or fname.startswith('.'):
                continue
            
            aging_type, aging_time, sio2_pct, sample_num = extract_info_from_filename(fname, folder_name)
            
            if aging_type is None or aging_time is None or sio2_pct is None:
                continue
            
            # 初始化数据结构
            if aging_type not in result:
                result[aging_type] = {}
            if aging_time not in result[aging_type]:
                result[aging_type][aging_time] = {}
            if sio2_pct not in result[aging_type][aging_time]:
                result[aging_type][aging_time][sio2_pct] = {'max_loads': [], 'max_disps': []}
            
            filepath = os.path.join(folder_path, fname)
            df = read_excel_data(filepath)
            if df is None:
                continue
            
            result[aging_type][aging_time][sio2_pct]['max_loads'].append(df['LoadValue'].max())
            result[aging_type][aging_time][sio2_pct]['max_disps'].append(df['PositionValue'].max())
    
    return result

def filter_and_aggregate(raw_data):
    """
    剔除异常值并聚合数据
    返回：{aging_type: {aging_time: {sio2_pct: {'load_mean':, 'load_std':, ...}}}}
    """
    result = {}
    
    for aging_type, time_data in raw_data.items():
        result[aging_type] = {}
        
        for aging_time, sio2_data in time_data.items():
            result[aging_type][aging_time] = {}
            
            for sio2_pct, data in sio2_data.items():
                if len(data['max_loads']) == 0:
                    continue
                
                # 剔除异常值
                filtered_loads, removed_loads = remove_outliers_3sigma(data['max_loads'])
                filtered_disps, removed_disps = remove_outliers_3sigma(data['max_disps'])
                
                if removed_loads:
                    print(f"      {aging_type}-{aging_time}-{sio2_pct}: 剔除{len(removed_loads)}个异常值")
                
                # 使用筛选后的数据
                all_removed = set(removed_loads + removed_disps)
                valid_indices = [i for i in range(len(data['max_loads'])) if i not in all_removed]
                
                if len(valid_indices) == 0:
                    continue
                
                valid_loads = [data['max_loads'][i] for i in valid_indices]
                valid_disps = [data['max_disps'][i] for i in valid_indices]
                
                result[aging_type][aging_time][sio2_pct] = {
                    'load_mean': np.mean(valid_loads),
                    'load_std': np.std(valid_loads, ddof=1) if len(valid_loads) > 1 else 0,
                    'disp_mean': np.mean(valid_disps),
                    'disp_std': np.std(valid_disps, ddof=1) if len(valid_disps) > 1 else 0,
                    'n': len(valid_loads)
                }
    
    return result

# ========== 生成变化柱状图（黑白打印机友好，单图模式） ==========
def make_single_chart(data_dict, aging_type, test_type, sio2_pct, output_path, ylabel, y_unit, is_displacement=False):
    """
    为某个SiO2含量生成单张柱状图
    data_dict: {aging_time: {sio2_pct: {'load_mean':, 'load_std':, ...}}}
    """
    # 准备数据
    means = []
    stds = []
    ns = []

    for tp in ['初始', '5天', '10天']:
        if tp in data_dict and sio2_pct in data_dict[tp]:
            if is_displacement:
                means.append(data_dict[tp][sio2_pct]['disp_mean'])
                stds.append(data_dict[tp][sio2_pct]['disp_std'])
            else:
                means.append(data_dict[tp][sio2_pct]['load_mean'])
                stds.append(data_dict[tp][sio2_pct]['load_std'])
            ns.append(data_dict[tp][sio2_pct]['n'])
        else:
            means.append(0)
            stds.append(0)
            ns.append(0)

    if all(m == 0 for m in means):
        return

    fig, ax = plt.subplots(figsize=(7, 5))

    x = np.arange(len(time_points))
    bar_width = 0.5

    for i, tp in enumerate(time_points):
        if means[i] > 0:
            ax.bar(i, means[i], bar_width,
                  yerr=stds[i],
                  capsize=5,
                  color=time_colors[tp],
                  edgecolor='black',
                  linewidth=1.2,
                  hatch=time_hatch[tp],
                  error_kw={'elinewidth': 1.2, 'capthick': 1.2, 'ecolor': 'black'})

            y_pos = means[i] + stds[i] + max(means) * 0.02
            ax.text(i, y_pos, f'{means[i]:.1f}',
                    ha='center', va='bottom', fontsize=9)

    ax.set_xticks(x)
    ax.set_xticklabels(time_points, fontsize=10.5)
    ax.set_xlabel('老化时间', fontsize=10.5)
    ax.set_ylabel(f'{ylabel} ({y_unit})', fontsize=10.5)
    ax.set_title(f'{aging_type}老化对纳米SiO₂含量{sio2_pct}的{test_type}{ylabel}的影响', fontsize=11, fontweight='bold', pad=12)
    ax.grid(True, axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # Y轴范围
    y_max_val = max([m + s for m, s in zip(means, stds)]) if any(m > 0 for m in means) else 0
    if y_max_val > 0:
        ax.set_ylim(bottom=0, top=y_max_val * 1.18)

    # 图例
    legend_handles = [
        mpatches.Patch(facecolor=time_colors['初始'], edgecolor='black', hatch=time_hatch['初始'], label='初始'),
        mpatches.Patch(facecolor=time_colors['5天'], edgecolor='black', hatch=time_hatch['5天'], label='5天'),
        mpatches.Patch(facecolor=time_colors['10天'], edgecolor='black', hatch=time_hatch['10天'], label='10天')
    ]
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
        print(f"处理{test_type}试验数据")
        print('='*60)

        # 处理数据
        raw_data = process_test_type(base_dir, test_type)

        if not raw_data:
            print(f"  无数据")
            continue

        # 将初始组数据复制到湿热和紫外老化组中作为对照
        if '初始' in raw_data and '初始' in raw_data['初始']:
            initial_data = raw_data['初始']['初始']
            for aging_type in ['湿热', '紫外']:
                if aging_type in raw_data:
                    if '初始' not in raw_data[aging_type]:
                        raw_data[aging_type]['初始'] = {}
                    for sio2_pct, data in initial_data.items():
                        raw_data[aging_type]['初始'][sio2_pct] = data
                    print(f"  已将初始组数据加入{aging_type}老化组作为对照")

        print(f"  剔除异常值...")
        processed_data = filter_and_aggregate(raw_data)

        if not processed_data:
            print(f"  处理后无有效数据")
            continue
        
        # 为每种老化类型和每个SiO2含量生成单图
        for aging_type in ['湿热', '紫外']:
            if aging_type not in processed_data:
                print(f"  {aging_type}老化: 无数据")
                continue
            
            print(f"  生成{aging_type}老化对比图...")
            
            for sio2_pct in ['0%', '1%', '3%', '5%']:
                # 检查是否有数据
                has_data = any(tp in processed_data[aging_type] and sio2_pct in processed_data[aging_type][tp] for tp in ['初始', '5天', '10天'])
                if not has_data:
                    continue
                
                # 生成强度变化图
                make_single_chart(
                    processed_data[aging_type],
                    aging_type,
                    test_type,
                    sio2_pct,
                    f'charts/{aging_type}老化_{test_type}强度变化_{sio2_pct}_黑白版.png',
                    ylabel=f'{test_type}强度',
                    y_unit='N',
                    is_displacement=False
                )
                
                # 生成位移变化图
                make_single_chart(
                    processed_data[aging_type],
                    aging_type,
                    test_type,
                    sio2_pct,
                    f'charts/{aging_type}老化_{test_type}位移变化_{sio2_pct}_黑白版.png',
                    ylabel=f'{test_type}位移',
                    y_unit='mm',
                    is_displacement=True
                )
    
    print("\n" + "="*60)
    print("变化对比图表生成完成!")
    print("="*60)
    print("\n图表说明：")
    print("  - 每张图包含4个子图：SiO₂含量 0%、1%、3%、5%")
    print("  - X轴：初始、5天、10天")
    print("  - 黑白打印区分：")
    print("    初始=纯黑无填充")
    print("    5天=中灰斜线填充")
    print("    10天=浅灰点填充")
