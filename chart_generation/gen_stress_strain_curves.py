#!/usr/bin/env python3
"""
生成论文应力-应变曲线图（解决Issue 3 & 4）
- 将原始力-位移数据转换为应力-应变曲线
- 标注弹性变形阶段与塑性变形阶段
- 使用正确的中文坐标轴标签
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

# ============ 试样几何参数 ============
TENSILE_AREA = 40.0       # mm² (10mm × 4mm)
TENSILE_GAUGE_LENGTH = 50.0  # mm
FLEXURAL_SPAN = 64.0      # mm
FLEXURAL_WIDTH = 10.0     # mm
FLEXURAL_THICKNESS = 4.0  # mm

# ============ 路径配置 ============
BASE_DIR = '/Users/shiberlin/Desktop/毕业论文'
CHARTS_DIR = os.path.join(BASE_DIR, 'charts')

# ============ 老化条件关键词（用于文件名匹配） ============
# 用Unicode码点来匹配中文关键词
AGING_KEYWORDS = {
    'initial': {'match': lambda s: not any(kw in s for kw in ['紫外', '湿热', '耦合', '5+5'])},
    'uv_5d': {'match': lambda s: '紫外' in s and '5' in s and '10' not in s and '5+5' not in s},
    'uv_10d': {'match': lambda s: '紫外' in s and '10' in s},
    'hydro_5d': {'match': lambda s: '湿热' in s and '5' in s and '10' not in s and '5+5' not in s},
    'hydro_10d': {'match': lambda s: '湿热' in s and '10' in s},
    'coupled_5d': {'match': lambda s: ('耦合' in s or '5+5' in s)},
    'coupled_10d': {'match': lambda s: ('耦合' in s and '10' in s)},
}

# ============ SiO₂含量提取 ============
def extract_sio2(filename):
    """从文件名提取SiO2含量"""
    m = re.search(r'(\d+)%', filename)
    if m:
        return int(m.group(1))
    return None

def extract_sample_num(filename):
    """从文件名提取样品编号"""
    # 最后一个数字（1-4）
    nums = re.findall(r'(\d)(?!\d*%)', filename.replace('.xlsx', ''))
    if nums:
        return int(nums[-1])
    return None

# ============ 数据读取与处理 ============
def read_raw_data(filepath):
    """读取原始力-位移数据"""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    
    times = []
    forces = []
    positions = []
    
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] is not None and row[2] is not None:
            times.append(float(row[0]) if row[0] is not None else 0)
            forces.append(float(row[1]))
            positions.append(float(row[2]))
    
    return np.array(times), np.array(forces), np.array(positions)

def find_load_start(forces, positions, threshold_ratio=0.02, min_consecutive=50):
    """
    自动检测实际加载起始点（去除预加载段）
    
    策略：寻找力值开始持续上升的点
    1. 计算力值的滑动窗口标准差
    2. 找到力值显著上升的起始位置
    """
    n = len(forces)
    if n < min_consecutive:
        return 0
    
    # 方法：计算力值的变化率，找到持续上升的起点
    # 先平滑
    window = 20
    forces_smooth = np.convolve(forces, np.ones(window)/window, mode='same')
    
    # 找最大力值
    max_force = np.max(forces_smooth)
    if max_force < 1:
        return 0
    
    # 从后往前找力值低于阈值（最大力×threshold_ratio）的点
    threshold = max_force * threshold_ratio
    
    # 从最后往前搜索，找到力值持续低于阈值的区域结束点
    start_idx = 0
    below_count = 0
    for i in range(n):
        if forces_smooth[i] < threshold:
            below_count += 1
        else:
            if below_count >= min_consecutive and start_idx == 0:
                start_idx = max(0, i - below_count - 10)
                break
            below_count = 0
    
    # 确保不从最后面开始
    if start_idx >= n * 0.8:
        start_idx = 0
    
    return start_idx

def force_to_stress_tensile(forces):
    """拉伸力转应力: σ(MPa) = F(N) / 40mm²"""
    return forces / TENSILE_AREA

def position_to_strain_tensile(positions, start_idx):
    """位移转应变: ε(%) = ΔL(mm) / 50mm × 100%"""
    if start_idx >= len(positions):
        return np.zeros_like(positions)
    delta_L = positions - positions[start_idx]
    return (delta_L / TENSILE_GAUGE_LENGTH) * 100.0

def force_to_stress_flexural(forces):
    """弯曲力转应力: σ_f = 3FL/(2bh²) = 0.6×F"""
    return (3.0 * FLEXURAL_SPAN * forces) / (2.0 * FLEXURAL_WIDTH * FLEXURAL_THICKNESS**2)

def identify_elastic_region(stress, strain, r_squared_threshold=0.98):
    """
    识别弹性变形区域
    使用线性回归逐步扩展，当R²低于阈值时停止
    """
    n = len(strain)
    if n < 20:
        return 0, n - 1
    
    # 从起始点开始，逐步扩展弹性区域
    best_end = min(20, n - 1)
    
    for end in range(20, min(n, int(n * 0.6))):
        x = strain[:end]
        y = stress[:end]
        
        if len(x) < 10:
            continue
        
        # 线性回归
        A = np.vstack([x, np.ones(len(x))]).T
        try:
            m, c = np.linalg.lstsq(A, y, rcond=None)[0]
            y_pred = m * x + c
            ss_res = np.sum((y - y_pred) ** 2)
            ss_tot = np.sum((y - np.mean(y)) ** 2)
            
            if ss_tot < 1e-10:
                continue
                
            r2 = 1 - ss_res / ss_tot
            
            if r2 < r_squared_threshold:
                break
            
            best_end = end
        except:
            break
    
    return 0, best_end

def process_single_file(filepath, test_type='tensile'):
    """处理单个数据文件，返回应力-应变曲线数据"""
    times, forces, positions = read_raw_data(filepath)
    
    if len(forces) < 100:
        return None, None, None, None, None
    
    # 去除预加载段
    start_idx = find_load_start(forces, positions)
    
    # 截取有效数据
    forces = forces[start_idx:]
    positions = positions[start_idx:]
    
    if len(forces) < 100:
        return None, None, None, None, None
    
    # 转换为应力-应变
    if test_type == 'tensile':
        stress = force_to_stress_tensile(forces)
        strain = position_to_strain_tensile(positions, 0)  # 已截取，从0开始
    else:
        stress = force_to_stress_flexural(forces)
        strain = positions - positions[0]  # 挠度(mm)
    
    # 识别弹性区域
    elastic_start, elastic_end = identify_elastic_region(stress, strain)
    
    # 找最大应力点（断裂点）
    max_idx = np.argmax(stress)
    
    return strain, stress, elastic_end, max_idx, (forces, positions)

# ============ 数据处理主流程 ============
def categorize_all_files():
    """遍历所有数据文件并分类"""
    categories = defaultdict(list)  # key: (test_type, aging_condition, sio2)
    
    for root, dirs, files in os.walk(BASE_DIR):
        # 跳过非数据目录
        if any(skip in root for skip in ['.workbuddy', 'charts', 'node_modules', '.docx_work']):
            continue
        
        for f in files:
            if not f.endswith('.xlsx') or f.startswith('~$'):
                continue
            
            filepath = os.path.join(root, f)
            rel = os.path.relpath(root, BASE_DIR)
            
            # 判断测试类型
            if '拉' in rel or '拉' in f:
                test_type = 'tensile'
            elif '弯' in rel or '弯' in f:
                test_type = 'flexural'
            else:
                continue
            
            # 提取SiO₂含量
            sio2 = extract_sio2(f)
            if sio2 is None:
                continue
            
            # 提取样品号
            sample_num = extract_sample_num(f)
            
            # 判断老化条件
            aging_condition = None
            if '初始' in rel or '初始' in f:
                aging_condition = 'initial'
            elif '耦合' in f or '5+5' in f:
                aging_condition = 'coupled_5d'
            elif '紫外' in f:
                if '10' in rel or '10' in f:
                    aging_condition = 'uv_10d'
                else:
                    aging_condition = 'uv_5d'
            elif '湿热' in f:
                if '10' in rel or '10' in f:
                    aging_condition = 'hydro_10d'
                else:
                    aging_condition = 'hydro_5d'
            
            if aging_condition is None:
                continue
            
            key = (test_type, aging_condition, sio2)
            categories[key].append((filepath, sample_num))
    
    return categories

# ============ 图表生成函数 ============
# 线条样式（适配黑白打印）
LINE_STYLES = {
    0: {'color': '#000000', 'linestyle': '-', 'linewidth': 1.5, 'marker': 'o', 'markersize': 3, 'markevery': 500},
    1: {'color': '#333333', 'linestyle': '--', 'linewidth': 1.5, 'marker': 's', 'markersize': 3, 'markevery': 500},
    3: {'color': '#666666', 'linestyle': '-.', 'linewidth': 1.5, 'marker': '^', 'markersize': 3, 'markevery': 500},
    5: {'color': '#999999', 'linestyle': ':', 'linewidth': 1.5, 'marker': 'd', 'markersize': 3, 'markevery': 500},
}

AGING_CONDITION_LABELS = {
    'initial': '初始组',
    'uv_5d': '紫外老化5天',
    'uv_10d': '紫外老化10天',
    'hydro_5d': '湿热老化5天',
    'hydro_10d': '湿热老化10天',
    'coupled_5d': '紫外湿热耦合老化5天',
    'coupled_10d': '紫外湿热耦合老化10天',
}

CHAPTER_MAP = {
    'initial': ('3', '初始组'),
    'uv_5d': ('4', '紫外老化5天'),
    'uv_10d': ('4', '紫外老化10天'),
    'hydro_5d': ('5', '湿热老化5天'),
    'hydro_10d': ('5', '湿热老化10天'),
    'coupled_5d': ('6', '紫外湿热耦合老化5天'),
    'coupled_10d': ('6', '紫外湿热耦合老化10天'),
}

def select_representative_curve(file_data_list):
    """
    从多个平行样品中选择最代表性的曲线
    策略：选择最大应力最接近中位数的样品
    """
    if not file_data_list:
        return None, None, None, None, None, None
    
    results = []
    for filepath, sample_num in file_data_list:
        strain, stress, elastic_end, max_idx, raw = process_single_file(filepath, 'tensile')
        if strain is not None:
            results.append((filepath, sample_num, strain, stress, elastic_end, max_idx, raw))
    
    if not results:
        return None, None, None, None, None, None
    
    # 选择最大应力最接近中位数的样品
    max_stresses = [np.max(r[3]) for r in results]
    median_stress = np.median(max_stresses)
    best_idx = np.argmin([abs(s - median_stress) for s in max_stresses])
    
    return results[best_idx]

def generate_tensile_figure_for_chapter(categories, chapter_conditions, output_name, title_prefix):
    """
    为某一章生成拉伸应力-应变曲线图
    每个子图显示4个SiO₂水平的代表性曲线
    """
    # 收集这一章需要的条件
    fig_conditions = []
    for cond_key in chapter_conditions:
        for cat_key, file_list in categories.items():
            test_type, aging_cond, sio2 = cat_key
            if test_type == 'tensile' and aging_cond == cond_key:
                fig_conditions.append(cond_key)
                break
    
    fig_conditions = sorted(set(fig_conditions), key=lambda x: chapter_conditions.index(x) if x in chapter_conditions else 99)
    n_conditions = len(fig_conditions)
    
    if n_conditions == 0:
        print(f"  No data for {output_name}")
        return
    
    # 创建图形
    fig, axes = plt.subplots(1, n_conditions, figsize=(6 * n_conditions, 5))
    if n_conditions == 1:
        axes = [axes]
    
    for ax_idx, cond_key in enumerate(fig_conditions):
        ax = axes[ax_idx]
        
        for sio2 in [0, 1, 3, 5]:
            cat_key = ('tensile', cond_key, sio2)
            file_list = categories.get(cat_key, [])
            
            if not file_list:
                continue
            
            # 选择代表性曲线
            result = select_representative_curve(file_list)
            if result is None or result[2] is None:
                continue
            
            filepath, sample_num, strain, stress, elastic_end, max_idx, raw = result
            
            style = LINE_STYLES.get(sio2, LINE_STYLES[0])
            label = f'{sio2}% SiO₂'
            
            ax.plot(strain, stress, 
                   color=style['color'], linestyle=style['linestyle'],
                   linewidth=style['linewidth'],
                   label=label)
            
        ax.legend(fontsize=9, loc='upper left', framealpha=0.8)
        
        ax.set_xlabel('应变 / %', fontsize=12)
        ax.set_ylabel('拉伸应力 / MPa', fontsize=12)
        ax.set_title(AGING_CONDITION_LABELS.get(cond_key, cond_key), fontsize=13, fontweight='bold')
        ax.tick_params(labelsize=11)
        ax.grid(True, alpha=0.3, linestyle='--')
        ax.set_xlim(left=0)
        ax.set_ylim(bottom=0)
    
    plt.tight_layout()
    output_path = os.path.join(CHARTS_DIR, output_name)
    fig.savefig(output_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Saved: {output_name}")

def generate_all_tensile_curves(categories):
    """生成所有章节的拉伸应力-应变曲线"""
    print("\n=== 生成拉伸应力-应变曲线 ===")
    
    # 第三章: 初始组
    generate_tensile_figure_for_chapter(
        categories,
        ['initial'],
        'Fig3-1_initial_stress_strain.png',
        '初始组'
    )
    
    # 第四章: 紫外老化（5天+10天）
    uv_conds = [k for k in AGING_CONDITION_LABELS if k.startswith('uv_')]
    if uv_conds:
        generate_tensile_figure_for_chapter(
            categories,
            uv_conds,
            'Fig4-1_uv_stress_strain.png',
            '紫外老化'
        )
    
    # 第五章: 湿热老化（5天+10天）
    hydro_conds = [k for k in AGING_CONDITION_LABELS if k.startswith('hydro_')]
    if hydro_conds:
        generate_tensile_figure_for_chapter(
            categories,
            hydro_conds,
            'Fig5-1_hydro_stress_strain.png',
            '湿热老化'
        )
    
    # 第六章: 耦合老化
    coupled_conds = [k for k in AGING_CONDITION_LABELS if k.startswith('coupled_')]
    if coupled_conds:
        generate_tensile_figure_for_chapter(
            categories,
            coupled_conds,
            'Fig6-1_coupled_stress_strain.png',
            '紫外湿热耦合老化'
        )

# ============ 单图多曲线对比（展示弹性-塑性阶段） ============
def generate_detailed_stress_strain(categories, cond_key, sio2=0, output_name=None):
    """
    为特定条件生成详细应力-应变曲线，展示所有平行样品
    用于展示弹性-塑性转变
    """
    cat_key = ('tensile', cond_key, sio2)
    file_list = categories.get(cat_key, [])
    
    if not file_list:
        print(f"  No data for {AGING_CONDITION_LABELS.get(cond_key)} SiO2={sio2}%")
        return
    
    fig, ax = plt.subplots(1, 1, figsize=(8, 6))
    
    colors = ['#1a1a1a', '#4a4a4a', '#7a7a7a', '#aaaaaa']
    
    for i, (filepath, sample_num) in enumerate(sorted(file_list, key=lambda x: x[1])):
        strain, stress, elastic_end, max_idx, raw = process_single_file(filepath, 'tensile')
        if strain is None:
            continue
        
        color = colors[i % len(colors)]
        ax.plot(strain, stress, color=color, linewidth=1.2, label=f'样品{sample_num}')
        
        # 标注弹性区域
        if elastic_end and elastic_end > 10:
            # 对第一个样品标注弹性模量线
            if i == 0:
                x_elastic = strain[:elastic_end]
                y_elastic = stress[:elastic_end]
                if len(x_elastic) > 5:
                    A = np.vstack([x_elastic, np.ones(len(x_elastic))]).T
                    m, c = np.linalg.lstsq(A, y_elastic, rcond=None)[0]
                    x_line = np.array([strain[0], strain[elastic_end]])
                    y_line = m * x_line + c
                    ax.plot(x_line, y_line, 'r--', linewidth=1.5, alpha=0.7, label='弹性模量拟合')
        
        # 标注断裂点
        if max_idx:
            ax.scatter([strain[max_idx]], [stress[max_idx]], 
                      marker='x', color=color, s=50, zorder=5)
    
    ax.set_xlabel('应变 / %', fontsize=14)
    ax.set_ylabel('拉伸应力 / MPa', fontsize=14)
    ax.set_title(f'{AGING_CONDITION_LABELS.get(cond_key, cond_key)} {sio2}% SiO₂ 应力-应变曲线', 
                fontsize=15, fontweight='bold')
    ax.legend(fontsize=11, loc='upper left')
    ax.tick_params(labelsize=12)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.set_xlim(left=0)
    ax.set_ylim(bottom=0)
    
    plt.tight_layout()
    if output_name is None:
        output_name = f'detail_{cond_key}_sio2{sio2}.png'
    output_path = os.path.join(CHARTS_DIR, output_name)
    fig.savefig(output_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Saved: {output_name}")

# ============ 主程序 ============
def main():
    print("=" * 60)
    print("应力-应变曲线生成脚本（解决Issue 3 & 4）")
    print("=" * 60)
    
    # 1. 分类所有数据文件
    print("\n[1/3] 分类数据文件...")
    categories = categorize_all_files()
    
    # 打印统计
    stats = defaultdict(int)
    for (test_type, aging_cond, sio2), files in categories.items():
        stats[(test_type, aging_cond)] += len(files)
    
    print(f"  共找到 {sum(stats.values())} 个数据文件")
    for (tt, ac), count in sorted(stats.items()):
        label = AGING_CONDITION_LABELS.get(ac, ac)
        print(f"    {tt:8s} | {label:20s} | {count} files")
    
    # 2. 生成各章拉伸应力-应变曲线
    print("\n[2/3] 生成应力-应变曲线...")
    generate_all_tensile_curves(categories)
    
    # 3. 生成详细对比图（展示弹性-塑性阶段）
    print("\n[3/3] 生成详细曲线（展示弹性-塑性阶段）...")
    
    # 为每个条件选SiO₂=0%做详细展示
    for cond_key in ['initial', 'uv_5d', 'uv_10d', 'hydro_5d', 'hydro_10d', 'coupled_5d']:
        generate_detailed_stress_strain(
            categories, cond_key, sio2=0,
            output_name=f'detail_stress_strain_{cond_key}_0pct.png'
        )
    
    print("\n" + "=" * 60)
    print("完成！所有曲线已保存到 charts/ 目录")
    print("=" * 60)
    
    # 输出数据摘要
    print("\n=== 拉伸性能数据摘要 ===")
    for cond_key in ['initial', 'uv_5d', 'uv_10d', 'hydro_5d', 'hydro_10d', 'coupled_5d']:
        label = AGING_CONDITION_LABELS.get(cond_key, cond_key)
        print(f"\n{label}:")
        for sio2 in [0, 1, 3, 5]:
            cat_key = ('tensile', cond_key, sio2)
            file_list = categories.get(cat_key, [])
            if not file_list:
                continue
            
            max_stresses = []
            elongations = []
            for filepath, sample_num in file_list:
                strain, stress, _, max_idx, _ = process_single_file(filepath, 'tensile')
                if strain is not None and max_idx is not None:
                    max_stresses.append(stress[max_idx])
                    elongations.append(strain[max_idx])
            
            if max_stresses:
                avg_stress = np.mean(max_stresses)
                std_stress = np.std(max_stresses)
                avg_elong = np.mean(elongations)
                std_elong = np.std(elongations)
                print(f"  SiO₂={sio2}%: σ_max={avg_stress:.2f}±{std_stress:.2f} MPa, "
                      f"ε_b={avg_elong:.2f}±{std_elong:.2f}%")

if __name__ == '__main__':
    main()
