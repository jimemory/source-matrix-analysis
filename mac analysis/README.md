# Source Matrix 分析工具

分析不同 `Type L1` 在各维度下的 source 数量分布情况，支持 PCR 分类统计和 Generation Portfolio 维度。

---

## 环境搭建

```bash
# 1. 创建虚拟环境
uv venv --python 3.11

# 2. 激活环境
source .venv/bin/activate

# 3. 安装依赖
uv pip install pandas openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
```

---

## 运行分析

```bash
python source_analysis.py
```

---

## Source 分类逻辑

| 分类 | 判断条件 | 说明 |
|------|----------|------|
| **Source Total** | Part Number 唯一值总数 | 所有 source 的总数量 |
| **Source introduced by PCR** | `PCR` 列有值的行 | 通过 PCR 流程引入的 source |
| **Source in BC scope** | `PCR` 列无值(None/NaN)的行 | BC 范围内的 source |

---

## 分析维度

| 维度 | 说明 | Generation Portfolio |
|------|------|---------------------|
| **Family Name + BC Volume** | 按产品族统计 source 数量，BC Volume 按 Family 去重 | ✅ 包含 |
| **Platform** | 按平台类型统计 source 数量 | ✅ 组合维度 |
| **Project Position** | 按产品定位统计 source 数量 | ✅ 组合维度 |
| **Form Factor** | 按产品形态统计 source 数量 | ✅ 组合维度 |
| **Generation Portfolio** | 按产品代际组合统计 source 数量 | ✅ 主维度 |
| **Sourcing Strategy** | 按采购策略统计 source 数量 | ✅ 组合维度 |

---

## 核心逻辑

- **Source 数量**: 统计 `Part Number(Mandatory)` 的唯一值数量
- **BC Volume 去重**: 同一 `Family Name` 的 BC Volume 只计算一次，避免重复
- **PCR 分类**: 根据 `PCR` 列是否有值进行分类统计
  - PCR 列有值 → Source introduced by PCR
  - PCR 列无值 → Source in BC scope
- **Generation Portfolio**: 所有维度分析都包含 Generation Portfolio 信息

---

## 输出文件

| 文件 | 说明 |
|------|------|
| `source_analysis_report.xlsx` | Excel 分析报告，包含多个工作表 |

### Excel 工作表说明

| 工作表 | 行数 | 列数 | 说明 |
|--------|------|------|------|
| `总览` | 4 | 4 | 各 Type L1 的 Source Total / Source introduced by PCR / Source in BC scope |
| **`PCR统计汇总`** | **13** | **18** | ⭐ **新增**: 按 **Type L1 + Generation Portfolio** 维度的 PCR 统计 + 重叠分析 + **完整的 8 列 Spec/Source/BC 数据** |
| `Family_BCVolume` | 230 | 12 | Family Name + BC Volume 详细分析（含 PCR 分类和上下文列） |
| `Platform` | 79 | **9** | Platform + Generation Portfolio 组合分析（含 PCR 分类 + **Prod QTY + Spec**） |
| `Project_Position` | 36 | **9** | Project Position + Generation Portfolio 组合分析（含 PCR 分类 + **Prod QTY + Spec**） |
| `Form_Factor` | 29 | **9** | Form Factor + Generation Portfolio 组合分析（含 PCR 分类 + **Prod QTY + Spec**） |
| `Generation_Portfolio` | 13 | **8** | Generation Portfolio 维度分析（含 PCR 分类 + **Prod QTY + Spec**） |
| `Sourcing_Strategy` | 33 | **9** | Sourcing Strategy + Generation Portfolio 组合分析（含 PCR 分类 + **Prod QTY + Spec**） |
| `交叉分析` | 228 | 12 | Type L1 + Family + Platform + Position + Generation 交叉分析 |
| **`Spec分析`** | **13** | **13** | ⭐ **新增**: Spec 复用度分析，对比两种 Source QTY/Spec 计算方式 |

### PCR统计汇总 Sheet

按 **Type L1 + Generation Portfolio** 维度展示 PCR 分布比例、重叠情况、产品数量和 **完整的 8 列 Spec/Source/BC 数据**：

| Type L1 | Gen | Prod QTY | Spec(去重) | Source(去重) | Spec(不去重) | Source(不去重) | BC(去重) | BC(不去重) | 比率(去重/去重) | 比率(不去重/不去重) | BC/Spec(不去重/不去重) |
|---------|-----|---------|-----------|-------------|-------------|---------------|---------|-----------|--------------|------------------|---------------------|
| Component LP5X | FY2425 | 38 | 9 | 31 | 121 | 366 | 25 | 210 | 3.44 | 3.02 | 1.74 |
| Component LP5X | FY2526 | 23 | 11 | 34 | 51 | 162 | 28 | 120 | 3.09 | 3.18 | 2.35 |
| Component LP5X | FY2627 | 27 | 8 | 24 | 96 | 192 | 17 | 140 | 3.00 | 2.00 | 1.46 |
| Module DIMM 5 | FY2425 | 16 | 6 | 41 | 70 | 239 | 21 | 174 | 6.83 | 3.41 | 2.49 |
| Module DIMM 5 | FY2526 | 58 | 11 | 74 | 119 | 431 | 42 | 248 | 6.73 | 3.62 | 2.08 |
| Module DIMM 5 | FY2627 | 28 | 5 | 51 | 97 | 282 | 44 | 181 | 10.20 | 2.91 | 1.87 |
| Component DDR5 | FY2425 | 3 | 1 | 5 | 3 | 5 | 3 | 3 | 5.00 | 1.67 | 1.00 |

#### 8 列核心数据说明

| 列名 | 说明 | 计算方式 | LP5X FY2425 |
|------|------|---------|------------|
| **Spec Total (去重)** | Spec 种类数（全局唯一） | Spec(Mandatory) 唯一值数量 | 9 |
| **Source QTY (去重)** | Source 种类数（全局唯一） | Part Number 唯一值数量 | 31 |
| **Spec Total (不去重)** | Spec 总次数（含重复） | Sum(各 Family 的 Spec 数量) | 121 |
| **Source QTY (不去重)** | Source 总次数（含重复） | Sum(各 Family 的 Source 数量) | 366 |
| **Source in BC scope** | BC scope Source 数（去重） | PCR 为空的 Part Number 唯一值数量 | 25 |
| **Source in BC scope (不去重)** | BC scope Source 数（不去重） | Sum(各 Family 的 BC Source 数量) | 210 |
| **Source QTY/Spec (去重/去重)** | 整体标准化比率 | Source QTY (去重) / Spec Total (去重) | 3.44 |
| **Source QTY/Spec (不去重/不去重)** | 产品级平均标准化比率 | Source QTY (不去重) / Spec Total (不去重) | 3.02 |
| **Source in BC scope/Spec (不去重/不去重)** | BC 级别平均标准化比率 | BC Source (不去重) / Spec Total (不去重) | 1.74 |

#### 关键解读

**去重 vs 不去重的差异说明**:

**Component LP5X - FY2425**:
- **Spec**: 9 (去重) vs 121 (不去重) = **13.4 倍差异**
  - 说明每个 Spec 平均被 13.4 个 Family 使用
  
- **Source**: 31 (去重) vs 366 (不去重) = **11.8 倍差异**
  - 说明每个 Source 平均被 11.8 个 Family 使用

- **BC Source**: 25 (去重) vs 210 (不去重) = **8.4 倍差异**
  - BC scope 内的 Source 复用度

**三种比率对比**:
- **去重/去重 (3.44)**: 全局视角，每个 Spec 对应 3.44 个 Source
- **不去重/不去重 (3.02)**: 产品级平均，每个 Spec 对应 3.02 个 Source
- **BC/Spec (不去重/不去重) (1.74)**: BC scope 内，每个 Spec 对应 1.74 个 Source

#### 为什么 Source Total ≠ PCR + BC scope？

**原因**: 同一个 Part Number 可能同时出现在 PCR 有值和无值的行中（重叠）

**解释**:
- 同一个 Part Number 可能用于多个产品（Family）
- 某些产品有 PCR 记录，某些产品没有 PCR 记录
- 因此：Source Total = PCR + BC scope - Overlap

**示例**: Component LP5X - FY2425
- Source introduced by PCR: 28
- Source in BC scope: 25  
- Overlap (同时在 PCR 和 BC 中): 22
- Source Total = 28 + 25 - 22 = **31** ✓

### Spec分析 Sheet

深入分析 Spec 复用度和两种 Source QTY/Spec 计算方式的差异：

| Type L1 | Gen | Families | Sources | Specs | 汇总去重 | Family平均 | 差异 | 最大覆盖 | 最通用Spec |
|---------|-----|---------|---------|-------|---------|-----------|------|---------|-----------|
| Component LP5X | FY2425 | 38 | 31 | 9 | 3.44 | 3.24 | 0.21 | 34 | Memory 4GB LPDDR5X 7500 |
| Module DIMM 5 | FY2526 | 58 | 74 | 11 | 6.73 | 9.16 | 2.43 | 53 | Memory 16GB DDR5 5600 |
| Component DDR5 | FY2425 | 3 | 5 | 1 | 5.00 | 5.00 | 0.00 | 3 | Memory 2GB DDR5 5600 |

#### 两种 Source QTY/Spec 计算方式对比

| 计算方式 | 计算方法 | 业务含义 | 适用场景 |
|---------|---------|---------|----------|
| **汇总去重** | Total Sources / Total Specs<br>(先汇总去重，再相除) | 反映**整体标准化程度**<br>同一Spec在多个Family中只算一次 | 评估代际/类型层面的<br>整体标准化水平 |
| **Family平均** | Average(Family Sources / Family Specs)<br>(先按Family计算，再求平均) | 反映**产品级标准化程度**<br>考虑Spec在不同产品间的复用 | 评估产品级别的<br>标准化设计水平 |

#### 为什么会有差异？

**核心原因**: 同一个 Spec 可能出现在多个 Family 中

**示例分析**: Component LP5X - FY2425
- 总 Family 数: 38
- 总 Sources (去重): 31
- 总 Specs (去重): 9
- **Spec 平均覆盖 13.4 个 Family**
- 最通用的 Spec (`Memory 4GB LPDDR5X 7500`) 出现在 **34 个 Family** 中

**两种计算方式差异**: 3.44 vs 3.24 = **0.21**

**差异说明**:
- 差异小 (如 Component LP5: 0.00) → Spec 复用度高，标准化程度一致
- 差异大 (如 Module DIMM 5 FY2526: 2.43) → Spec 在不同 Family 中分布不均，有的产品用了更多非通用 Spec

### 上下文列

`Family_BCVolume` 和 `交叉分析` 工作表包含以下上下文列：
- `From Factor`
- `Sourcing Strategy`
- `Plarform`
- `Project Position`
- `Dev Type`
- `Generation Portfolio`

---

## 分析结果摘要

### Type L1 Source 分布（含 PCR 分类 - 按 Generation Portfolio 汇总）

| Type L1 | Source Total | Source introduced by PCR | PCR % | Source in BC scope | BC % | Overlap | Overlap % |
|---------|-------------|-------------------------|-------|-------------------|------|---------|-----------|
| **Module DIMM 5** | 166 | 115 | 69.3% | 107 | 64.5% | 56 | 33.7% |
| **Component LP5X** | 95 | 60 | 63.2% | 76 | 80.0% | 41 | 43.2% |
| **Component LP5** | 7 | 7 | 100.0% | 4 | 57.1% | 4 | 57.1% |
| **Component DDR5** | 15 | 3 | 20.0% | 12 | 80.0% | 0 | 0.0% |

### Generation Portfolio 分布

| Type L1 | FY2425 | FY2526 | FY2627 | FY2728 |
|---------|--------|--------|--------|--------|
| Component LP5X | 31 | 34 | 24 | 6 |
| Component LP5 | 4 | 2 | 1 | - |
| Module DIMM 5 | 41 | 74 | 51 | - |
| Component DDR5 | 5 | 5 | 5 | - |

### 关键发现

#### PCR 分析
1. **PCR 占比最高**: Component LP5 (100%) 的 source 全部通过 PCR 引入
2. **PCR 占比最低**: Component DDR5 仅 20.0% 通过 PCR 引入，大部分在 BC scope 内
3. **重叠分析**: Component LP5X 的 FY2425 代际有 71.0% 的 source 同时出现在 PCR 和 BC scope 中

#### 产品数量 (Prod QTY) 分析
4. **Module DIMM 5 - FY2526** 产品数量最多 (58 个 Family)
5. **Component LP5X - FY2425** 有 38 个产品，但 Source QTY/Spec 比率较高 (3.44)

#### Source QTY/Spec 分析
6. **Module DIMM 5 - FY2627** 的 Source QTY/Spec 最高 (10.20)，说明每个 Spec 对应更多的 Source
7. **Component LP5** 的 Source QTY/Spec 最低 (1.00)，说明每个 Spec 只对应一个 Source
8. **Component DDR5** 的 Source QTY/Spec 为 5.00，标准化程度较低
9. **两种计算方式差异最大**: Module DIMM 5 - FY2526 (差异 2.43)，说明 Spec 在不同 Family 中分布不均匀

---

## 项目结构

```
.
├── source matrix with component type for AI.xlsx   # 输入数据
├── source_analysis.py                               # 分析脚本
├── source_analysis_report.xlsx                      # 输出报告
├── README.md                                        # 本文档
└── .venv/                                           # Python 虚拟环境
```

---

## 更新记录

| 日期 | 版本 | 说明 |
|------|------|------|
| 2026-03-10 | v1.0 | 初始版本，支持5个维度的 source 数量分析 |
| 2026-03-10 | v1.1 | 新增 Generation Portfolio 维度分析 |
| 2026-03-10 | v1.2 | Family Name 和交叉分析工作表新增上下文列 |
| 2026-03-10 | v1.3 | 新增 PCR 分类统计：Source Total / Source introduced by PCR / Source in BC scope |
| 2026-03-10 | v1.4 | 新增 PCR统计汇总 sheet；所有维度分析增加 Generation Portfolio 列 |
| 2026-03-10 | v1.5 | PCR统计汇总增加 Overlap 分析，解释 PCR + BC scope - Overlap = Total 的关系 |
| 2026-03-10 | v1.6 | PCR统计汇总增加 Generation Portfolio 维度，支持按代际分析 PCR 分布 |
| 2026-03-10 | v1.7 | 新增 Prod QTY (产品数量) 和 Source QTY/Spec (Spec 比率) 指标，所有工作表统一更新 |
| 2026-03-10 | v1.8 | 新增 Spec分析工作表，深入分析 Spec 复用度和两种 Source QTY/Spec 计算方式的差异 |
| 2026-03-10 | v1.9 | PCR统计汇总 sheet 新增完整的 8 列 Spec/Source/BC 数据（含 BC scope 不去重计数和比率） |
