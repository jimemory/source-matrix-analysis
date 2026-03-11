#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
 visualize raw data.xlsx 数据分析和可视化
对比各个 Type L1 的 Source 指标在过去 Gen 和 Gen12 Target 的表现
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')  # 使用非交互式后端
import matplotlib.pyplot as plt
import numpy as np

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

def load_data():
    """加载数据"""
    df = pd.read_excel('visualize raw data.xlsx', sheet_name='Sheet1')
    print("=== 数据加载成功 ===")
    print(f"数据形状: {df.shape}")
    print(f"\n数据预览:")
    print(df.to_string(index=False))
    return df

def analyze_gen_trends(df):
    """分析各 Type L1 在过去 Gen 的趋势"""
    print("\n" + "="*80)
    print("各 Type L1 在不同 Generation 的趋势分析")
    print("="*80)
    
    for type_l1 in df['Type L1'].unique():
        type_df = df[df['Type L1'] == type_l1].sort_values('Generation Portfolio')
        target = type_df['Gen12 Target'].iloc[0]
        
        print(f"\n【{type_l1}】")
        print(f"Gen12 Target: {target}")
        print(type_df[['Generation Portfolio', 'Source in BC scope/Spec (不去重/不去重)', 
                      'Source QTY/Spec (不去重/不去重)', 'Gen12 Target']].to_string(index=False))
        
        # 计算增长情况
        gen_list = type_df['Generation Portfolio'].tolist()
        bc_list = type_df['Source in BC scope/Spec (不去重/不去重)'].tolist()
        src_list = type_df['Source QTY/Spec (不去重/不去重)'].tolist()
        
        print(f"\n  BC/Spec 变化: {' -> '.join([f'{x:.2f}' for x in bc_list])}")
        print(f"  Source/Spec 变化: {' -> '.join([f'{x:.2f}' for x in src_list])}")
        
        # 最新 Gen 与 Target 对比
        latest_bc = bc_list[-1]
        latest_src = src_list[-1]
        print(f"\n  最新 FY2627 vs Gen12 Target:")
        print(f"    BC/Spec: {latest_bc:.2f} vs {target} (差距: {latest_bc - target:+.2f})")
        print(f"    Source/Spec: {latest_src:.2f} vs {target} (差距: {latest_src - target:+.2f})")

def create_visualization(df):
    """创建可视化图表 - 3个图表布局"""
    
    # 获取唯一的 Type L1
    type_l1_list = df['Type L1'].unique()
    gen_list = ['FY2425', 'FY2526', 'FY2627']
    
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle('Type L1 Source Analysis: Past Gen vs Gen12 Target', fontsize=16, fontweight='bold')
    
    colors = {'FY2425': '#3498db', 'FY2526': '#e74c3c', 'FY2627': '#2ecc71'}
    
    # ===== 图1: Source in BC scope / Spec (左上) =====
    ax1 = axes[0, 0]
    x = np.arange(len(type_l1_list))
    width = 0.25
    
    for i, gen in enumerate(gen_list):
        values = []
        targets = []
        for type_l1 in type_l1_list:
            type_df = df[(df['Type L1'] == type_l1) & (df['Generation Portfolio'] == gen)]
            if len(type_df) > 0:
                values.append(type_df['Source in BC scope/Spec (不去重/不去重)'].iloc[0])
                targets.append(type_df['Gen12 Target'].iloc[0])
            else:
                values.append(0)
                targets.append(0)
        
        bars = ax1.bar(x + i*width, values, width, label=gen, color=colors[gen], alpha=0.8)
        
        # 添加数值标签（水平摆放）
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax1.text(bar.get_x() + bar.get_width()/2., height,
                        f'{height:.2f}',
                        ha='center', va='bottom', fontsize=8)
    
    # 添加 Target 线
    for j, type_l1 in enumerate(type_l1_list):
        target = df[df['Type L1'] == type_l1]['Gen12 Target'].iloc[0]
        ax1.plot([j-0.4, j+0.65], [target, target], 'k--', linewidth=2, alpha=0.5)
        ax1.text(j+0.65, target, f'Target: {target}', va='center', fontsize=9)
    
    ax1.set_xlabel('Type L1', fontsize=12)
    ax1.set_ylabel('Source in BC scope / Spec', fontsize=12)
    ax1.set_title('Source in BC scope / Spec', fontsize=13, fontweight='bold')
    ax1.set_xticks(x + width)
    ax1.set_xticklabels(type_l1_list)
    ax1.legend()
    ax1.grid(axis='y', alpha=0.3)
    
    # ===== 图2: Total Source QTY / Spec (右上) =====
    ax2 = axes[0, 1]
    
    for i, gen in enumerate(gen_list):
        values = []
        for type_l1 in type_l1_list:
            type_df = df[(df['Type L1'] == type_l1) & (df['Generation Portfolio'] == gen)]
            if len(type_df) > 0:
                values.append(type_df['Source QTY/Spec (不去重/不去重)'].iloc[0])
            else:
                values.append(0)
        
        bars = ax2.bar(x + i*width, values, width, label=gen, color=colors[gen], alpha=0.8)
        
        # 添加数值标签（水平摆放）
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax2.text(bar.get_x() + bar.get_width()/2., height,
                        f'{height:.2f}',
                        ha='center', va='bottom', fontsize=8)
    
    # 添加 Target 线
    for j, type_l1 in enumerate(type_l1_list):
        target = df[df['Type L1'] == type_l1]['Gen12 Target'].iloc[0]
        ax2.plot([j-0.4, j+0.65], [target, target], 'k--', linewidth=2, alpha=0.5)
        ax2.text(j+0.65, target, f'Target: {target}', va='center', fontsize=9)
    
    ax2.set_xlabel('Type L1', fontsize=12)
    ax2.set_ylabel('Total Source QTY / Spec', fontsize=12)
    ax2.set_title('Total Source QTY / Spec', fontsize=13, fontweight='bold')
    ax2.set_xticks(x + width)
    ax2.set_xticklabels(type_l1_list)
    ax2.legend()
    ax2.grid(axis='y', alpha=0.3)
    
    # ===== 图3: Source / Spec Compare BC, with PCR and Gen12 Guidance (左下) =====
    ax3 = axes[1, 0]
    
    x = np.arange(len(type_l1_list))
    width = 0.25
    
    # 使用 FY2627 的最新数据
    bc_values = []
    src_values = []
    targets = []
    
    for type_l1 in type_l1_list:
        type_df = df[(df['Type L1'] == type_l1) & (df['Generation Portfolio'] == 'FY2627')]
        if len(type_df) > 0:
            bc_values.append(type_df['Source in BC scope/Spec (不去重/不去重)'].iloc[0])
            src_values.append(type_df['Source QTY/Spec (不去重/不去重)'].iloc[0])
            targets.append(type_df['Gen12 Target'].iloc[0])
        else:
            bc_values.append(0)
            src_values.append(0)
            targets.append(0)
    
    bars1 = ax3.bar(x - width, bc_values, width, label='BC/Spec (FY2627)', color='#3498db', alpha=0.8)
    bars2 = ax3.bar(x, src_values, width, label='Source/Spec (FY2627)', color='#e74c3c', alpha=0.8)
    bars3 = ax3.bar(x + width, targets, width, label='Gen12 Target', color='#2ecc71', alpha=0.8)
    
    # 添加数值标签
    for bars in [bars1, bars2, bars3]:
        for bar in bars:
            height = bar.get_height()
            ax3.text(bar.get_x() + bar.get_width()/2., height,
                    f'{height:.2f}',
                    ha='center', va='bottom', fontsize=8)
    
    ax3.set_xlabel('Type L1', fontsize=12)
    ax3.set_ylabel('Source QTY / Spec', fontsize=12)
    ax3.set_title('Source / Spec Compare BC, with PCR and Gen12 Guidance', fontsize=13, fontweight='bold')
    ax3.set_xticks(x)
    ax3.set_xticklabels(type_l1_list)
    ax3.legend()
    ax3.grid(axis='y', alpha=0.3)
    
    # ===== 图4: Spec Total (去重) Trend with Gen12 SBB Plan (右下) =====
    ax4 = axes[1, 1]
    
    x = np.arange(len(type_l1_list))
    width = 0.25
    
    # 获取每个 Type L1 的 Gen12 SBB Plan (用于虚线)
    sbb_plan_dict = {}
    for type_l1 in type_l1_list:
        type_df = df[df['Type L1'] == type_l1]
        sbb_plan_dict[type_l1] = type_df['Gen12 SBB Plan'].iloc[0]
    
    # 收集各 Generation 的 Spec Total (去重)
    spec_by_gen = {gen: [] for gen in gen_list}
    
    for type_l1 in type_l1_list:
        type_df = df[df['Type L1'] == type_l1].sort_values('Generation Portfolio')
        for gen in gen_list:
            gen_data = type_df[type_df['Generation Portfolio'] == gen]
            if len(gen_data) > 0:
                spec_by_gen[gen].append(gen_data['Spec Total (去重)'].iloc[0])
            else:
                spec_by_gen[gen].append(0)
    
    # 绘制各 Generation 的柱状图 (使用与前三个图相同的 colors 字典)
    for i, gen in enumerate(gen_list):
        offset = width * (i - 1)  # -width, 0, +width
        bars = ax4.bar(x + offset, spec_by_gen[gen], width, 
                      label=gen, color=colors[gen], alpha=0.8)
        
        # 添加数值标签（水平摆放）
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax4.text(bar.get_x() + bar.get_width()/2., height,
                        f'{int(height)}',
                        ha='center', va='bottom', fontsize=9, fontweight='bold')
    
    # 添加 Gen12 SBB Plan 虚线（与 Target 线样式一致）
    for i, type_l1 in enumerate(type_l1_list):
        sbb_plan = sbb_plan_dict[type_l1]
        ax4.hlines(y=sbb_plan, xmin=i-0.4, xmax=i+0.4, 
                  colors='#555555', linestyles='--', linewidth=2)
        ax4.text(i+0.35, sbb_plan, f'SBB Plan: {int(sbb_plan)}', 
                ha='right', va='bottom', fontsize=8, color='#555555')
    
    ax4.set_xlabel('Type L1', fontsize=12)
    ax4.set_ylabel('Spec Total (Dedup)', fontsize=12)
    ax4.set_title('Spec Total (Dedup) Trend vs Gen12 SBB Plan', fontsize=13, fontweight='bold')
    ax4.set_xticks(x)
    ax4.set_xticklabels(type_l1_list)
    ax4.legend()
    ax4.grid(axis='y', alpha=0.3)
    
    plt.tight_layout()
    plt.savefig('visualization_analysis.png', dpi=150, bbox_inches='tight')
    print("\n✅ 可视化图表已保存: visualization_analysis.png")
    plt.show()

def generate_summary_report(df):
    """生成总结报告"""
    print("\n" + "="*80)
    print("总结报告: Multiple Source 数量增长分析")
    print("="*80)
    
    print("\n📊 关键发现:\n")
    
    for type_l1 in df['Type L1'].unique():
        type_df = df[df['Type L1'] == type_l1].sort_values('Generation Portfolio')
        target = type_df['Gen12 Target'].iloc[0]
        
        # 获取最新数据
        latest_df = type_df[type_df['Generation Portfolio'] == 'FY2627']
        latest_bc = latest_df['Source in BC scope/Spec (不去重/不去重)'].iloc[0]
        latest_src = latest_df['Source QTY/Spec (不去重/不去重)'].iloc[0]
        
        # 计算增长情况
        first_bc = type_df['Source in BC scope/Spec (不去重/不去重)'].iloc[0]
        last_bc = type_df['Source in BC scope/Spec (不去重/不去重)'].iloc[-1]
        first_src = type_df['Source QTY/Spec (不去重/不去重)'].iloc[0]
        last_src = type_df['Source QTY/Spec (不去重/不去重)'].iloc[-1]
        
        bc_growth = ((last_bc - first_bc) / first_bc * 100) if first_bc > 0 else 0
        src_growth = ((last_src - first_src) / first_src * 100) if first_src > 0 else 0
        
        print(f"【{type_l1}】")
        print(f"  Gen12 Target: {target}")
        print(f"  FY2627 实际值:")
        print(f"    - BC/Spec: {latest_bc:.2f} ({'✅ 达标' if latest_bc <= target else '❌ 超标'})")
        print(f"    - Source/Spec: {latest_src:.2f} ({'✅ 达标' if latest_src <= target else '❌ 超标'})")
        print(f"  从 FY2425 到 FY2627 的增长:")
        print(f"    - BC/Spec: {bc_growth:+.1f}% ({'📈 增长' if bc_growth > 0 else '📉 下降'})")
        print(f"    - Source/Spec: {src_growth:+.1f}% ({'📈 增长' if src_growth > 0 else '📉 下降'})")
        
        # 趋势分析
        if bc_growth > 20 or src_growth > 20:
            print(f"  ⚠️  Warning: Multiple source 数量增长显著")
        elif bc_growth < -20 or src_growth < -20:
            print(f"  ✅ Good: Multiple source 数量得到有效控制")
        else:
            print(f"  ➡️  Stable: Multiple source 数量相对稳定")
        print()

def main():
    """主函数"""
    print("="*80)
    print("Visualize Raw Data 分析和可视化")
    print("="*80)
    
    # 1. 加载数据
    df = load_data()
    
    # 2. 分析趋势
    analyze_gen_trends(df)
    
    # 3. 创建可视化
    create_visualization(df)
    
    # 4. 生成总结报告
    generate_summary_report(df)
    
    print("\n" + "="*80)
    print("分析完成!")
    print("="*80)

if __name__ == "__main__":
    main()
