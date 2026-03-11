#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Source Matrix Analysis Tool
分析不同 Type L1 在各维度下的 source 数量分布情况
分类统计: Source Total / Source introduced by PCR / Source in BC scope
"""

import pandas as pd
import numpy as np
from pathlib import Path


def load_data(file_path):
    """加载Excel数据"""
    df = pd.read_excel(file_path)
    print(f"数据维度: {df.shape}")
    print(f"列名: {df.columns.tolist()[:10]}...")
    
    # PCR 列统计信息
    print(f"\nPCR 列统计:")
    print(f"  有值行数 (Source introduced by PCR): {df['PCR'].notna().sum()}")
    print(f"  无值行数 (Source in BC scope): {df['PCR'].isna().sum()}")
    
    return df


def get_unique_sources_by_description(df, group_by_cols=None):
    """
    基于 Part Description 前 30 个字符 + Supplier 去重后的唯一 source 计数

    去重规则：在同一个 Family Name + Supplier(Mandatory) 组合内，
    如果 Part Description 前 30 个字符相同，则视为同一个 source

    参数:
        df: DataFrame 数据
        group_by_cols: 定义"同一个产品"的分组列，默认为 ['Family Name']

    返回:
        dict: {'total': int, 'pcr': int, 'bc_scope': int}
    """
    if group_by_cols is None:
        group_by_cols = ['Family Name']

    # 检查必需的列是否存在
    desc_col = 'Part Description (No Need Input,Auto From Windchill)'
    supplier_col = 'Supplier(Mandatory)'

    if desc_col not in df.columns:
        # 如果 Description 列不存在，回退到 Part Number 计数
        total = df['Part Number(Mandatory)'].nunique()
        pcr_sources = df[df['PCR'].notna()]['Part Number(Mandatory)'].nunique() if df['PCR'].notna().any() else 0
        bc_scope_sources = df[df['PCR'].isna()]['Part Number(Mandatory)'].nunique() if df['PCR'].isna().any() else 0
        return {'total': total, 'pcr': pcr_sources, 'bc_scope': bc_scope_sources}

    # 创建 Part Description 前 30 个字符的临时列
    df = df.copy()
    df['_desc_prefix_30'] = df[desc_col].fillna('').astype(str).str[:30]

    # 去重：在同一个 Family Name + Supplier(Mandatory) 组合内，按 desc_prefix_30 去重
    unique_sources = set()
    unique_pcr_sources = set()
    unique_bc_sources = set()

    # 构建唯一键：(Family Name, Supplier, Description前30字符)
    for _, row in df.iterrows():
        # 构建分组键
        group_key = tuple(row[col] for col in group_by_cols if col in row)
        supplier = row.get(supplier_col, '') if supplier_col in row else ''
        desc_prefix = row['_desc_prefix_30']
        has_pcr = pd.notna(row.get('PCR'))

        # 唯一键：(group_key, supplier, desc_prefix)
        unique_key = (group_key, supplier, desc_prefix)

        # 统计所有 source
        unique_sources.add(unique_key)

        # 分别统计 PCR 和 BC scope
        if has_pcr:
            unique_pcr_sources.add(unique_key)
        else:
            unique_bc_sources.add(unique_key)

    return {
        'total': len(unique_sources),
        'pcr': len(unique_pcr_sources),
        'bc_scope': len(unique_bc_sources)
    }


def count_sources_by_pcr(df, use_description_dedup=True):
    """
    按 PCR 列分类统计 source 数量和 Spec 数量
    - Source Total: Part Description 前 30 字符去重后的唯一 source 数量
    - Source introduced by PCR: PCR 列有值的 Source 唯一值数量
    - Source in BC scope: PCR 列无值的 Source 唯一值数量
    - Spec Total: Spec(Mandatory) 唯一值数量
    """
    if use_description_dedup:
        # 使用基于 Part Description 前 30 字符的去重逻辑
        source_counts = get_unique_sources_by_description(df)
        total = source_counts['total']
        pcr_sources = source_counts['pcr']
        bc_scope_sources = source_counts['bc_scope']
    else:
        # 回退到原始 Part Number 去重逻辑
        total = df['Part Number(Mandatory)'].nunique()
        pcr_sources = df[df['PCR'].notna()]['Part Number(Mandatory)'].nunique() if df['PCR'].notna().any() else 0
        bc_scope_sources = df[df['PCR'].isna()]['Part Number(Mandatory)'].nunique() if df['PCR'].isna().any() else 0

    # 统计 Spec(Mandatory) 数量（保持不变）
    spec_total = df['Spec(Mandatory)'].nunique() if 'Spec(Mandatory)' in df.columns else 0

    return {
        'Source Total': total,
        'Source introduced by PCR': pcr_sources,
        'Source in BC scope': bc_scope_sources,
        'Spec Total': spec_total
    }


def count_prod_qty(df):
    """
    统计产品数量 (Family Name 的唯一值数量)
    """
    return df['Family Name'].nunique() if 'Family Name' in df.columns else 0


def analyze_by_family_name_bc_volume(df):
    """
    按 Family Name + BC Volume(pcs) 分析 source 数量
    BC Volume 需要按 Family Name 去重
    添加上下文列：From Factor, Sourcing Strategy, Plarform, Project Position, Dev Type, Generation Portfolio
    """
    print("\n" + "="*80)
    print("分析维度 1: Family Name + BC Volume(pcs)")
    print("="*80)
    
    results = []
    
    for type_l1 in df['Type L1'].dropna().unique():
        type_df = df[df['Type L1'] == type_l1]
        
        # 按 Family Name 分组，计算 BC Volume（去重）和 Source 数量
        family_stats = []
        for family in type_df['Family Name'].dropna().unique():
            family_df = type_df[type_df['Family Name'] == family]
            
            # Source 数量统计（按 PCR 分类）
            source_counts = count_sources_by_pcr(family_df)
            
            # BC Volume - 按 Family Name 去重（同一 Family 的 BC Volume 只计算一次）
            bc_volume = family_df['BC Volume(pcs)'].iloc[0] if pd.notna(family_df['BC Volume(pcs)'].iloc[0]) else 0
            
            # 获取上下文信息（取第一个非空值）
            form_factor = family_df['From Factor'].dropna().iloc[0] if len(family_df['From Factor'].dropna()) > 0 else ''
            sourcing_strategy = family_df['Sourcing Strategy'].dropna().iloc[0] if len(family_df['Sourcing Strategy'].dropna()) > 0 else ''
            platform = family_df['Plarform'].dropna().iloc[0] if len(family_df['Plarform'].dropna()) > 0 else ''
            project_position = family_df['Project Position'].dropna().iloc[0] if len(family_df['Project Position'].dropna()) > 0 else ''
            dev_type = family_df['Dev Type'].dropna().iloc[0] if len(family_df['Dev Type'].dropna()) > 0 else ''
            generation_portfolio = family_df['Generation Portfolio'].dropna().iloc[0] if len(family_df['Generation Portfolio'].dropna()) > 0 else ''
            
            family_stats.append({
                'Family Name': family,
                'BC Volume(pcs)': bc_volume,
                'Source Total': source_counts['Source Total'],
                'Source introduced by PCR': source_counts['Source introduced by PCR'],
                'Source in BC scope': source_counts['Source in BC scope'],
                'From Factor': form_factor,
                'Sourcing Strategy': sourcing_strategy,
                'Plarform': platform,
                'Project Position': project_position,
                'Dev Type': dev_type,
                'Generation Portfolio': generation_portfolio
            })
        
        family_df_stats = pd.DataFrame(family_stats)
        total_sources = family_df_stats['Source Total'].sum()
        total_pcr = family_df_stats['Source introduced by PCR'].sum()
        total_bc_scope = family_df_stats['Source in BC scope'].sum()
        total_bc_volume = family_df_stats['BC Volume(pcs)'].sum()
        
        results.append({
            'Type L1': type_l1,
            'Source Total': total_sources,
            'Source introduced by PCR': total_pcr,
            'Source in BC scope': total_bc_scope,
            'Total BC Volume(pcs)': total_bc_volume,
            'Family Count': len(family_df_stats)
        })
        
        print(f"\n【{type_l1}】")
        print(f"  Family 数量: {len(family_df_stats)}")
        print(f"  Source Total: {total_sources}")
        print(f"  Source introduced by PCR: {total_pcr}")
        print(f"  Source in BC scope: {total_bc_scope}")
        print(f"  总 BC Volume: {total_bc_volume:,.0f} pcs")
        print(f"  明细:")
        # 按指定顺序显示列
        display_cols = ['Family Name', 'BC Volume(pcs)', 'Source Total', 'Source introduced by PCR', 'Source in BC scope',
                       'From Factor', 'Sourcing Strategy', 'Plarform', 'Project Position', 'Dev Type', 'Generation Portfolio']
        print(family_df_stats[display_cols].to_string(index=False))
    
    return pd.DataFrame(results)


def analyze_by_platform(df):
    """按 Platform + Generation Portfolio 分析 source 数量（含 PCR 分类）"""
    print("\n" + "="*80)
    print("分析维度 2: Platform + Generation Portfolio")
    print("="*80)
    
    results = []
    
    for type_l1 in df['Type L1'].dropna().unique():
        type_df = df[df['Type L1'] == type_l1]
        
        platform_stats = []
        for platform in type_df['Plarform'].dropna().unique():
            platform_df = type_df[type_df['Plarform'] == platform]
            
            # 对每个 Platform，再按 Generation Portfolio 细分
            for gen in platform_df['Generation Portfolio'].dropna().unique():
                gen_platform_df = platform_df[platform_df['Generation Portfolio'] == gen]
                source_counts = count_sources_by_pcr(gen_platform_df)
                
                platform_stats.append({
                    'Platform': platform,
                    'Generation Portfolio': gen,
                    'Source Total': source_counts['Source Total'],
                    'Source introduced by PCR': source_counts['Source introduced by PCR'],
                    'Source in BC scope': source_counts['Source in BC scope']
                })
        
        platform_df_stats = pd.DataFrame(platform_stats).sort_values(['Platform', 'Generation Portfolio'])
        
        # 按 Type L1 汇总
        total_sources = platform_df_stats['Source Total'].sum()
        total_pcr = platform_df_stats['Source introduced by PCR'].sum()
        total_bc_scope = platform_df_stats['Source in BC scope'].sum()
        
        results.append({
            'Type L1': type_l1,
            'Source Total': total_sources,
            'Source introduced by PCR': total_pcr,
            'Source in BC scope': total_bc_scope,
            'Platform-Gen Count': len(platform_df_stats)
        })
        
        print(f"\n【{type_l1}】")
        print(f"  Platform-Generation 组合数量: {len(platform_df_stats)}")
        print(f"  Source Total: {total_sources}")
        print(f"  Source introduced by PCR: {total_pcr}")
        print(f"  Source in BC scope: {total_bc_scope}")
        print(f"  明细:")
        print(platform_df_stats.to_string(index=False))
    
    return pd.DataFrame(results)


def analyze_by_project_position(df):
    """按 Project Position + Generation Portfolio 分析 source 数量（含 PCR 分类）"""
    print("\n" + "="*80)
    print("分析维度 3: Project Position + Generation Portfolio")
    print("="*80)
    
    results = []
    
    for type_l1 in df['Type L1'].dropna().unique():
        type_df = df[df['Type L1'] == type_l1]
        
        position_stats = []
        for position in type_df['Project Position'].dropna().unique():
            position_df = type_df[type_df['Project Position'] == position]
            
            # 对每个 Position，再按 Generation Portfolio 细分
            for gen in position_df['Generation Portfolio'].dropna().unique():
                gen_position_df = position_df[position_df['Generation Portfolio'] == gen]
                source_counts = count_sources_by_pcr(gen_position_df)
                
                position_stats.append({
                    'Project Position': position,
                    'Generation Portfolio': gen,
                    'Source Total': source_counts['Source Total'],
                    'Source introduced by PCR': source_counts['Source introduced by PCR'],
                    'Source in BC scope': source_counts['Source in BC scope']
                })
        
        position_df_stats = pd.DataFrame(position_stats).sort_values(['Project Position', 'Generation Portfolio'])
        
        total_sources = position_df_stats['Source Total'].sum()
        total_pcr = position_df_stats['Source introduced by PCR'].sum()
        total_bc_scope = position_df_stats['Source in BC scope'].sum()
        
        results.append({
            'Type L1': type_l1,
            'Source Total': total_sources,
            'Source introduced by PCR': total_pcr,
            'Source in BC scope': total_bc_scope,
            'Position-Gen Count': len(position_df_stats)
        })
        
        print(f"\n【{type_l1}】")
        print(f"  Position-Generation 组合数量: {len(position_df_stats)}")
        print(f"  Source Total: {total_sources}")
        print(f"  Source introduced by PCR: {total_pcr}")
        print(f"  Source in BC scope: {total_bc_scope}")
        print(f"  明细:")
        print(position_df_stats.to_string(index=False))
    
    return pd.DataFrame(results)


def analyze_by_form_factor(df):
    """按 Form Factor + Generation Portfolio 分析 source 数量（含 PCR 分类）"""
    print("\n" + "="*80)
    print("分析维度 4: Form Factor + Generation Portfolio")
    print("="*80)
    
    results = []
    
    for type_l1 in df['Type L1'].dropna().unique():
        type_df = df[df['Type L1'] == type_l1]
        
        form_stats = []
        for form in type_df['From Factor'].dropna().unique():
            form_df = type_df[type_df['From Factor'] == form]
            
            # 对每个 Form Factor，再按 Generation Portfolio 细分
            for gen in form_df['Generation Portfolio'].dropna().unique():
                gen_form_df = form_df[form_df['Generation Portfolio'] == gen]
                source_counts = count_sources_by_pcr(gen_form_df)
                
                form_stats.append({
                    'Form Factor': form,
                    'Generation Portfolio': gen,
                    'Source Total': source_counts['Source Total'],
                    'Source introduced by PCR': source_counts['Source introduced by PCR'],
                    'Source in BC scope': source_counts['Source in BC scope']
                })
        
        form_df_stats = pd.DataFrame(form_stats).sort_values(['Form Factor', 'Generation Portfolio'])
        
        total_sources = form_df_stats['Source Total'].sum()
        total_pcr = form_df_stats['Source introduced by PCR'].sum()
        total_bc_scope = form_df_stats['Source in BC scope'].sum()
        
        results.append({
            'Type L1': type_l1,
            'Source Total': total_sources,
            'Source introduced by PCR': total_pcr,
            'Source in BC scope': total_bc_scope,
            'Form-Gen Count': len(form_df_stats)
        })
        
        print(f"\n【{type_l1}】")
        print(f"  Form-Generation 组合数量: {len(form_df_stats)}")
        print(f"  Source Total: {total_sources}")
        print(f"  Source introduced by PCR: {total_pcr}")
        print(f"  Source in BC scope: {total_bc_scope}")
        print(f"  明细:")
        print(form_df_stats.to_string(index=False))
    
    return pd.DataFrame(results)


def analyze_by_generation_portfolio(df):
    """按 Generation Portfolio 分析 source 数量（含 PCR 分类）"""
    print("\n" + "="*80)
    print("分析维度 5: Generation Portfolio")
    print("="*80)
    
    results = []
    
    for type_l1 in df['Type L1'].dropna().unique():
        type_df = df[df['Type L1'] == type_l1]
        
        gen_stats = []
        for gen in type_df['Generation Portfolio'].dropna().unique():
            gen_df = type_df[type_df['Generation Portfolio'] == gen]
            source_counts = count_sources_by_pcr(gen_df)
            
            gen_stats.append({
                'Generation Portfolio': gen,
                'Source Total': source_counts['Source Total'],
                'Source introduced by PCR': source_counts['Source introduced by PCR'],
                'Source in BC scope': source_counts['Source in BC scope']
            })
        
        gen_df_stats = pd.DataFrame(gen_stats).sort_values('Generation Portfolio')
        
        total_sources = gen_df_stats['Source Total'].sum()
        total_pcr = gen_df_stats['Source introduced by PCR'].sum()
        total_bc_scope = gen_df_stats['Source in BC scope'].sum()
        
        results.append({
            'Type L1': type_l1,
            'Source Total': total_sources,
            'Source introduced by PCR': total_pcr,
            'Source in BC scope': total_bc_scope,
            'Generation Count': len(gen_df_stats)
        })
        
        print(f"\n【{type_l1}】")
        print(f"  Generation Portfolio 数量: {len(gen_df_stats)}")
        print(f"  Source Total: {total_sources}")
        print(f"  Source introduced by PCR: {total_pcr}")
        print(f"  Source in BC scope: {total_bc_scope}")
        print(f"  明细:")
        print(gen_df_stats.to_string(index=False))
    
    return pd.DataFrame(results)


def analyze_by_sourcing_strategy(df):
    """按 Sourcing Strategy + Generation Portfolio 分析 source 数量（含 PCR 分类）"""
    print("\n" + "="*80)
    print("分析维度 6: Sourcing Strategy + Generation Portfolio")
    print("="*80)
    
    results = []
    
    for type_l1 in df['Type L1'].dropna().unique():
        type_df = df[df['Type L1'] == type_l1]
        
        strategy_stats = []
        for strategy in type_df['Sourcing Strategy'].dropna().unique():
            strategy_df = type_df[type_df['Sourcing Strategy'] == strategy]
            
            # 对每个 Strategy，再按 Generation Portfolio 细分
            for gen in strategy_df['Generation Portfolio'].dropna().unique():
                gen_strategy_df = strategy_df[strategy_df['Generation Portfolio'] == gen]
                source_counts = count_sources_by_pcr(gen_strategy_df)
                
                strategy_stats.append({
                    'Sourcing Strategy': strategy,
                    'Generation Portfolio': gen,
                    'Source Total': source_counts['Source Total'],
                    'Source introduced by PCR': source_counts['Source introduced by PCR'],
                    'Source in BC scope': source_counts['Source in BC scope']
                })
        
        strategy_df_stats = pd.DataFrame(strategy_stats).sort_values(['Sourcing Strategy', 'Generation Portfolio'])
        
        total_sources = strategy_df_stats['Source Total'].sum()
        total_pcr = strategy_df_stats['Source introduced by PCR'].sum()
        total_bc_scope = strategy_df_stats['Source in BC scope'].sum()
        
        results.append({
            'Type L1': type_l1,
            'Source Total': total_sources,
            'Source introduced by PCR': total_pcr,
            'Source in BC scope': total_bc_scope,
            'Strategy-Gen Count': len(strategy_df_stats)
        })
        
        print(f"\n【{type_l1}】")
        print(f"  Strategy-Generation 组合数量: {len(strategy_df_stats)}")
        print(f"  Source Total: {total_sources}")
        print(f"  Source introduced by PCR: {total_pcr}")
        print(f"  Source in BC scope: {total_bc_scope}")
        print(f"  明细:")
        print(strategy_df_stats.to_string(index=False))
    
    return pd.DataFrame(results)


def generate_pcr_summary(df):
    """生成 PCR 统计汇总数据，包含重叠分析，按 Type L1 + Generation Portfolio 维度"""
    pcr_summary = []
    
    for type_l1 in df['Type L1'].dropna().unique():
        type_df = df[df['Type L1'] == type_l1]
        
        # 再按 Generation Portfolio 细分
        for gen in type_df['Generation Portfolio'].dropna().unique():
            gen_df = type_df[type_df['Generation Portfolio'] == gen]
            source_counts = count_sources_by_pcr(gen_df)
            prod_qty = count_prod_qty(gen_df)
            
            # 计算重叠的 Source（基于 Family + Supplier + Part Description 前 30 字符）
            desc_col = 'Part Description (No Need Input,Auto From Windchill)'
            supplier_col = 'Supplier(Mandatory)'
            if desc_col in gen_df.columns:
                # 创建 Description 前缀用于去重
                gen_df_copy = gen_df.copy()
                gen_df_copy['_desc_prefix_30'] = gen_df_copy[desc_col].fillna('').astype(str).str[:30]

                # 获取 PCR 和 BC scope 的 source（按 Family Name + Supplier + desc_prefix_30）
                pcr_df = gen_df_copy[gen_df_copy['PCR'].notna()]
                bc_df = gen_df_copy[gen_df_copy['PCR'].isna()]

                pcr_sources = set(zip(pcr_df['Family Name'], pcr_df[supplier_col], pcr_df['_desc_prefix_30']))
                bc_sources = set(zip(bc_df['Family Name'], bc_df[supplier_col], bc_df['_desc_prefix_30']))
                overlap_sources = pcr_sources & bc_sources
                overlap_count = len(overlap_sources)
            else:
                # 回退到 Part Number 逻辑
                pcr_parts = set(gen_df[gen_df['PCR'].notna()]['Part Number(Mandatory)'].dropna().unique())
                bc_parts = set(gen_df[gen_df['PCR'].isna()]['Part Number(Mandatory)'].dropna().unique())
                overlap_parts = pcr_parts & bc_parts
                overlap_count = len(overlap_parts)

            # 计算百分比和比率（去重版本 - 汇总后计算）
            total = source_counts['Source Total']
            spec_total = source_counts['Spec Total']
            pcr_pct = (source_counts['Source introduced by PCR'] / total * 100) if total > 0 else 0
            bc_pct = (source_counts['Source in BC scope'] / total * 100) if total > 0 else 0
            overlap_pct = (overlap_count / total * 100) if total > 0 else 0
            source_per_spec_dedup = (total / spec_total) if spec_total > 0 else 0

            # 计算不去重版本（按 Family 累加，基于 Part Description 前 30 字符去重）
            sum_sources_nodedup = 0
            sum_specs_nodedup = 0
            sum_bc_sources_nodedup = 0
            for family in gen_df['Family Name'].dropna().unique():
                family_df = gen_df[gen_df['Family Name'] == family]
                family_source_counts = get_unique_sources_by_description(family_df)
                sum_sources_nodedup += family_source_counts['total']
                sum_bc_sources_nodedup += family_source_counts['bc_scope']
                sum_specs_nodedup += family_df['Spec(Mandatory)'].nunique()
            
            # 计算 Source QTY/Spec 比率
            ratio_dedup_dedup = (total / spec_total) if spec_total > 0 else 0
            ratio_nodedup_nodedup = (sum_sources_nodedup / sum_specs_nodedup) if sum_specs_nodedup > 0 else 0
            ratio_bc_nodedup_nodedup = (sum_bc_sources_nodedup / sum_specs_nodedup) if sum_specs_nodedup > 0 else 0
            
            pcr_summary.append({
                'Type L1': type_l1,
                'Generation Portfolio': gen,
                'Prod QTY': prod_qty,
                'Source QTY (去重)': total,
                'Source QTY (不去重)': sum_sources_nodedup,
                'Source introduced by PCR': source_counts['Source introduced by PCR'],
                'PCR %': f"{pcr_pct:.1f}%",
                'Source in BC scope': source_counts['Source in BC scope'],
                'BC %': f"{bc_pct:.1f}%",
                'Source in BC scope (不去重)': sum_bc_sources_nodedup,
                'Source in BC scope/Spec (不去重/不去重)': f"{ratio_bc_nodedup_nodedup:.2f}",
                'Overlap (PCR & BC)': overlap_count,
                'Overlap %': f"{overlap_pct:.1f}%",
                'Spec Total (去重)': spec_total,
                'Spec Total (不去重)': sum_specs_nodedup,
                'Source QTY/Spec (去重/去重)': f"{ratio_dedup_dedup:.2f}",
                'Source QTY/Spec (不去重/不去重)': f"{ratio_nodedup_nodedup:.2f}",
                '说明': f'{source_counts["Source introduced by PCR"]} + {source_counts["Source in BC scope"]} - {overlap_count} = {total}'
            })
    
    return pd.DataFrame(pcr_summary)


def generate_excel_report(df, output_path):
    """生成Excel分析报告（含 PCR 分类统计）"""
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 1. 总览（含 PCR 分类）
        summary_data = []
        for type_l1 in df['Type L1'].dropna().unique():
            type_df = df[df['Type L1'] == type_l1]
            source_counts = count_sources_by_pcr(type_df)
            
            summary_data.append({
                'Type L1': type_l1,
                'Source Total': source_counts['Source Total'],
                'Source introduced by PCR': source_counts['Source introduced by PCR'],
                'Source in BC scope': source_counts['Source in BC scope']
            })
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='总览', index=False)
        
        # 2. PCR 统计汇总（新增）- 清晰展示 PCR 占比
        pcr_summary_df = generate_pcr_summary(df)
        pcr_summary_df.to_excel(writer, sheet_name='PCR统计汇总', index=False)
        
        # 3. Family Name + BC Volume 详细分析（含 PCR 分类和上下文列）
        family_results = []
        for type_l1 in df['Type L1'].dropna().unique():
            type_df = df[df['Type L1'] == type_l1]
            
            for family in type_df['Family Name'].dropna().unique():
                family_df = type_df[type_df['Family Name'] == family]
                source_counts = count_sources_by_pcr(family_df)
                bc_volume = family_df['BC Volume(pcs)'].iloc[0] if pd.notna(family_df['BC Volume(pcs)'].iloc[0]) else 0
                
                # 获取上下文信息
                form_factor = family_df['From Factor'].dropna().iloc[0] if len(family_df['From Factor'].dropna()) > 0 else ''
                sourcing_strategy = family_df['Sourcing Strategy'].dropna().iloc[0] if len(family_df['Sourcing Strategy'].dropna()) > 0 else ''
                platform = family_df['Plarform'].dropna().iloc[0] if len(family_df['Plarform'].dropna()) > 0 else ''
                project_position = family_df['Project Position'].dropna().iloc[0] if len(family_df['Project Position'].dropna()) > 0 else ''
                dev_type = family_df['Dev Type'].dropna().iloc[0] if len(family_df['Dev Type'].dropna()) > 0 else ''
                generation_portfolio = family_df['Generation Portfolio'].dropna().iloc[0] if len(family_df['Generation Portfolio'].dropna()) > 0 else ''
                
                family_results.append({
                    'Type L1': type_l1,
                    'Family Name': family,
                    'BC Volume(pcs)': bc_volume,
                    'Source Total': source_counts['Source Total'],
                    'Source introduced by PCR': source_counts['Source introduced by PCR'],
                    'Source in BC scope': source_counts['Source in BC scope'],
                    'From Factor': form_factor,
                    'Sourcing Strategy': sourcing_strategy,
                    'Plarform': platform,
                    'Project Position': project_position,
                    'Dev Type': dev_type,
                    'Generation Portfolio': generation_portfolio
                })
        
        family_results_df = pd.DataFrame(family_results)
        family_cols = ['Type L1', 'Family Name', 'BC Volume(pcs)', 
                      'Source Total', 'Source introduced by PCR', 'Source in BC scope',
                      'From Factor', 'Sourcing Strategy', 'Plarform', 
                      'Project Position', 'Dev Type', 'Generation Portfolio']
        family_results_df = family_results_df[family_cols]
        family_results_df.to_excel(writer, sheet_name='Family_BCVolume', index=False)
        
        # 4. Platform + Generation Portfolio 详细分析（含 PCR 分类和 Prod QTY/Spec）
        platform_results = []
        for type_l1 in df['Type L1'].dropna().unique():
            type_df = df[df['Type L1'] == type_l1]
            
            for platform in type_df['Plarform'].dropna().unique():
                platform_df = type_df[type_df['Plarform'] == platform]
                
                for gen in platform_df['Generation Portfolio'].dropna().unique():
                    gen_platform_df = platform_df[platform_df['Generation Portfolio'] == gen]
                    source_counts = count_sources_by_pcr(gen_platform_df)
                    prod_qty = count_prod_qty(gen_platform_df)
                    spec_total = source_counts['Spec Total']
                    source_per_spec = (source_counts['Source Total'] / spec_total) if spec_total > 0 else 0
                    
                    platform_results.append({
                        'Type L1': type_l1,
                        'Platform': platform,
                        'Generation Portfolio': gen,
                        'Prod QTY': prod_qty,
                        'Source Total': source_counts['Source Total'],
                        'Source introduced by PCR': source_counts['Source introduced by PCR'],
                        'Source in BC scope': source_counts['Source in BC scope'],
                        'Spec Total': spec_total,
                        'Source QTY/Spec': f"{source_per_spec:.2f}"
                    })
        
        platform_results_df = pd.DataFrame(platform_results)
        platform_results_df.to_excel(writer, sheet_name='Platform', index=False)
        
        # 5. Project Position + Generation Portfolio 详细分析（含 PCR 分类和 Prod QTY/Spec）
        position_results = []
        for type_l1 in df['Type L1'].dropna().unique():
            type_df = df[df['Type L1'] == type_l1]
            
            for position in type_df['Project Position'].dropna().unique():
                position_df = type_df[type_df['Project Position'] == position]
                
                for gen in position_df['Generation Portfolio'].dropna().unique():
                    gen_position_df = position_df[position_df['Generation Portfolio'] == gen]
                    source_counts = count_sources_by_pcr(gen_position_df)
                    prod_qty = count_prod_qty(gen_position_df)
                    spec_total = source_counts['Spec Total']
                    source_per_spec = (source_counts['Source Total'] / spec_total) if spec_total > 0 else 0
                    
                    position_results.append({
                        'Type L1': type_l1,
                        'Project Position': position,
                        'Generation Portfolio': gen,
                        'Prod QTY': prod_qty,
                        'Source Total': source_counts['Source Total'],
                        'Source introduced by PCR': source_counts['Source introduced by PCR'],
                        'Source in BC scope': source_counts['Source in BC scope'],
                        'Spec Total': spec_total,
                        'Source QTY/Spec': f"{source_per_spec:.2f}"
                    })
        
        position_results_df = pd.DataFrame(position_results)
        position_results_df.to_excel(writer, sheet_name='Project_Position', index=False)
        
        # 6. Form Factor + Generation Portfolio 详细分析（含 PCR 分类和 Prod QTY/Spec）
        form_results = []
        for type_l1 in df['Type L1'].dropna().unique():
            type_df = df[df['Type L1'] == type_l1]
            
            for form in type_df['From Factor'].dropna().unique():
                form_df = type_df[type_df['From Factor'] == form]
                
                for gen in form_df['Generation Portfolio'].dropna().unique():
                    gen_form_df = form_df[form_df['Generation Portfolio'] == gen]
                    source_counts = count_sources_by_pcr(gen_form_df)
                    prod_qty = count_prod_qty(gen_form_df)
                    spec_total = source_counts['Spec Total']
                    source_per_spec = (source_counts['Source Total'] / spec_total) if spec_total > 0 else 0
                    
                    form_results.append({
                        'Type L1': type_l1,
                        'Form Factor': form,
                        'Generation Portfolio': gen,
                        'Prod QTY': prod_qty,
                        'Source Total': source_counts['Source Total'],
                        'Source introduced by PCR': source_counts['Source introduced by PCR'],
                        'Source in BC scope': source_counts['Source in BC scope'],
                        'Spec Total': spec_total,
                        'Source QTY/Spec': f"{source_per_spec:.2f}"
                    })
        
        form_results_df = pd.DataFrame(form_results)
        form_results_df.to_excel(writer, sheet_name='Form_Factor', index=False)
        
        # 7. Generation Portfolio 详细分析（含 PCR 分类和 Prod QTY/Spec）
        gen_results = []
        for type_l1 in df['Type L1'].dropna().unique():
            type_df = df[df['Type L1'] == type_l1]
            
            for gen in type_df['Generation Portfolio'].dropna().unique():
                gen_df = type_df[type_df['Generation Portfolio'] == gen]
                source_counts = count_sources_by_pcr(gen_df)
                prod_qty = count_prod_qty(gen_df)
                spec_total = source_counts['Spec Total']
                source_per_spec = (source_counts['Source Total'] / spec_total) if spec_total > 0 else 0
                
                gen_results.append({
                    'Type L1': type_l1,
                    'Generation Portfolio': gen,
                    'Prod QTY': prod_qty,
                    'Source Total': source_counts['Source Total'],
                    'Source introduced by PCR': source_counts['Source introduced by PCR'],
                    'Source in BC scope': source_counts['Source in BC scope'],
                    'Spec Total': spec_total,
                    'Source QTY/Spec': f"{source_per_spec:.2f}"
                })
        
        gen_results_df = pd.DataFrame(gen_results)
        gen_results_df.to_excel(writer, sheet_name='Generation_Portfolio', index=False)
        
        # 8. Sourcing Strategy + Generation Portfolio 详细分析（含 PCR 分类和 Prod QTY/Spec）
        strategy_results = []
        for type_l1 in df['Type L1'].dropna().unique():
            type_df = df[df['Type L1'] == type_l1]
            
            for strategy in type_df['Sourcing Strategy'].dropna().unique():
                strategy_df = type_df[type_df['Sourcing Strategy'] == strategy]
                
                for gen in strategy_df['Generation Portfolio'].dropna().unique():
                    gen_strategy_df = strategy_df[strategy_df['Generation Portfolio'] == gen]
                    source_counts = count_sources_by_pcr(gen_strategy_df)
                    prod_qty = count_prod_qty(gen_strategy_df)
                    spec_total = source_counts['Spec Total']
                    source_per_spec = (source_counts['Source Total'] / spec_total) if spec_total > 0 else 0
                    
                    strategy_results.append({
                        'Type L1': type_l1,
                        'Sourcing Strategy': strategy,
                        'Generation Portfolio': gen,
                        'Prod QTY': prod_qty,
                        'Source Total': source_counts['Source Total'],
                        'Source introduced by PCR': source_counts['Source introduced by PCR'],
                        'Source in BC scope': source_counts['Source in BC scope'],
                        'Spec Total': spec_total,
                        'Source QTY/Spec': f"{source_per_spec:.2f}"
                    })
        
        strategy_results_df = pd.DataFrame(strategy_results)
        strategy_results_df.to_excel(writer, sheet_name='Sourcing_Strategy', index=False)
        
        # 9. 交叉分析: Type L1 + 所有维度（含 PCR 分类和完整上下文列）
        cross_results = []
        for type_l1 in df['Type L1'].dropna().unique():
            type_df = df[df['Type L1'] == type_l1]
            
            for family in type_df['Family Name'].dropna().unique():
                family_df = type_df[type_df['Family Name'] == family]
                bc_volume = family_df['BC Volume(pcs)'].iloc[0] if pd.notna(family_df['BC Volume(pcs)'].iloc[0]) else 0
                
                # 获取上下文信息
                form_factor = family_df['From Factor'].dropna().iloc[0] if len(family_df['From Factor'].dropna()) > 0 else ''
                sourcing_strategy = family_df['Sourcing Strategy'].dropna().iloc[0] if len(family_df['Sourcing Strategy'].dropna()) > 0 else ''
                dev_type = family_df['Dev Type'].dropna().iloc[0] if len(family_df['Dev Type'].dropna()) > 0 else ''
                
                for platform in family_df['Plarform'].dropna().unique():
                    platform_family_df = family_df[family_df['Plarform'] == platform]
                    
                    for position in platform_family_df['Project Position'].dropna().unique():
                        pos_df = platform_family_df[platform_family_df['Project Position'] == position]
                        
                        for gen in pos_df['Generation Portfolio'].dropna().unique():
                            gen_pos_df = pos_df[pos_df['Generation Portfolio'] == gen]
                            source_counts = count_sources_by_pcr(gen_pos_df)
                            
                            cross_results.append({
                                'Type L1': type_l1,
                                'Family Name': family,
                                'BC Volume(pcs)': bc_volume,
                                'Source Total': source_counts['Source Total'],
                                'Source introduced by PCR': source_counts['Source introduced by PCR'],
                                'Source in BC scope': source_counts['Source in BC scope'],
                                'From Factor': form_factor,
                                'Sourcing Strategy': sourcing_strategy,
                                'Plarform': platform,
                                'Project Position': position,
                                'Generation Portfolio': gen,
                                'Dev Type': dev_type
                            })
        
        cross_df = pd.DataFrame(cross_results)
        cross_cols = ['Type L1', 'Family Name', 'BC Volume(pcs)', 
                     'Source Total', 'Source introduced by PCR', 'Source in BC scope',
                     'From Factor', 'Sourcing Strategy', 'Plarform', 
                     'Project Position', 'Generation Portfolio', 'Dev Type']
        cross_df = cross_df[cross_cols]
        cross_df.to_excel(writer, sheet_name='交叉分析', index=False)
        
        # 10. Spec 深入分析工作表 - 分析 Spec 复用度和两种计算方式对比
        spec_analysis_results = []
        for type_l1 in df['Type L1'].dropna().unique():
            type_df = df[df['Type L1'] == type_l1]
            
            for gen in type_df['Generation Portfolio'].dropna().unique():
                gen_df = type_df[type_df['Generation Portfolio'] == gen]

                # 基础统计
                total_families = gen_df['Family Name'].nunique()
                # 使用基于 Part Description 前 30 字符的去重逻辑
                total_sources = get_unique_sources_by_description(gen_df)['total']
                total_specs = gen_df['Spec(Mandatory)'].nunique()
                source_per_spec = total_sources / total_specs if total_specs > 0 else 0

                # 按 Family 计算平均值（使用 Description 去重）
                family_ratios = []
                for family in gen_df['Family Name'].dropna().unique():
                    family_df = gen_df[gen_df['Family Name'] == family]
                    f_sources = get_unique_sources_by_description(family_df)['total']
                    f_specs = family_df['Spec(Mandatory)'].nunique()
                    if f_specs > 0:
                        family_ratios.append(f_sources / f_specs)
                avg_family_ratio = sum(family_ratios) / len(family_ratios) if family_ratios else 0
                
                # Spec 复用度分析
                spec_family_counts = gen_df.groupby('Spec(Mandatory)')['Family Name'].nunique()
                max_coverage = spec_family_counts.max() if len(spec_family_counts) > 0 else 0
                min_coverage = spec_family_counts.min() if len(spec_family_counts) > 0 else 0
                avg_coverage = spec_family_counts.mean() if len(spec_family_counts) > 0 else 0
                
                # 找出最通用的 Spec (覆盖最多 Family)
                most_common_spec = spec_family_counts.idxmax() if len(spec_family_counts) > 0 else ''
                most_common_count = spec_family_counts.max() if len(spec_family_counts) > 0 else 0
                
                spec_analysis_results.append({
                    'Type L1': type_l1,
                    'Generation Portfolio': gen,
                    'Total Families': total_families,
                    'Total Sources': total_sources,
                    'Total Specs': total_specs,
                    'Source QTY/Spec (汇总去重)': f"{source_per_spec:.2f}",
                    'Source QTY/Spec (Family平均)': f"{avg_family_ratio:.2f}",
                    '差异': f"{abs(source_per_spec - avg_family_ratio):.2f}",
                    'Spec最大覆盖Family数': max_coverage,
                    'Spec最小覆盖Family数': min_coverage,
                    'Spec平均覆盖Family数': f"{avg_coverage:.1f}",
                    '最通用Spec': most_common_spec[:50] + '...' if len(str(most_common_spec)) > 50 else most_common_spec,
                    '最通用Spec覆盖Family数': most_common_count
                })
        
        spec_analysis_df = pd.DataFrame(spec_analysis_results)
        spec_analysis_df.to_excel(writer, sheet_name='Spec分析', index=False)
    
    print(f"\n✅ Excel 报告已生成: {output_path}")


def main():
    """主函数"""
    input_file = "source matrix with component type for AI.xlsx"
    output_file = "source_analysis_report.xlsx"
    
    print("="*80)
    print("Source Matrix 分析工具 (含 PCR 分类统计)")
    print("="*80)
    
    # 1. 加载数据
    df = load_data(input_file)
    
    # 2. 执行各维度分析
    analyze_by_family_name_bc_volume(df)
    analyze_by_platform(df)
    analyze_by_project_position(df)
    analyze_by_form_factor(df)
    analyze_by_generation_portfolio(df)
    analyze_by_sourcing_strategy(df)
    
    # 3. 生成Excel报告
    generate_excel_report(df, output_file)
    
    print("\n" + "="*80)
    print("分析完成!")
    print("="*80)


if __name__ == "__main__":
    main()
