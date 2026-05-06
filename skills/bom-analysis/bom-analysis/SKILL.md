---
name: bom-analysis
description: BOM 表解析与成本计算技能。当用户询问 BOM（物料清单）、产品成本、出成率、原料耗用、辅料耗用、包材耗用，或提供月份+SKU 代码进行查询时触发。输出格式为飞书表格。
---

# BOM Analysis Skill

分析 BOM 表、计算产品成本、计算出成率。数据来源为 `/Users/taojing/Documents/OpenClaw_Docs/Bom/` 目录下的 xlsx 文件。

## ⚠️ 强制执行检查清单（每次执行前必须逐条确认）

> 陶晶反馈 + 高斯复盘（2026-05-06）：两次出错均因未遵循本清单。

### 处理 BOM 出成率任务前 — 必须打勾

- [ ] **文件路径**：BOM 文件在 `/Users/taojing/Documents/OpenClaw_Docs/Bom/`，不是 `~/.openclaw-dev/media/inbound/`
- [ ] **目标文件确认**：正式合并表 = 15 Sheet（2025年1-12月 + 2026年1-3月），样本文件 = 3 Sheet（2026年1-3月），不要混淆
- [ ] **只过滤原料**：只统计 `YS` / `YHTL` 前缀的物料，辅材(FS)、包材(BS)、佐料(ZJ)、辅料(FD) 一律排除
- [ ] **原料大类映射**：通过 `media/inbound/原料大类映射表*.xlsx` 最新文件映射，映射表中无匹配则用原始名称
- [ ] **先去重后计算**：同一 (SKU, 原料大类) 先合并耗用，再计算出成率（`入库重量 ÷ 汇总耗用`），禁止逐行独立计算
- [ ] **生成后验证**：打印各物料前缀的记录数，确认只有 YS/YHTL

### 处理 BOM 成本计算任务前 — 必须打勾

- [ ] **月份一致性**：小包装和大包装数据必须取自同一月份
- [ ] **产品信息打印**：计算前先打印 SKU 代码、品名、规格、入库重量
- [ ] **禁止跨月复用**：不得复用跨月份的中间值

---

## 核心规则（铁则）

1. **月份一致性**：小包装和大包装数据必须取自同一月份
2. **月份确认**：计算前先打印月份
3. **打印产品基础信息**：先输出产品代码、品名、规格、入库重量
4. **禁止跨月复用**：不得复用跨月份的中间值

## 目录结构

- BOM 文件路径：`/Users/taojing/Documents/OpenClaw_Docs/Bom/YYYY年M月bom表.xlsx`
- 月份命名：如 `2026年1月bom表.xlsx`

## Sheet 结构

| Sheet                         | 内容   | 最小包装单位 |
| ----------------------------- | ------ | ------------ |
| Sheet 1（Sheet 2 for 大包装） | 小包装 | 盒/袋        |
| Sheet 2（大包装）             | 大包装 | 箱           |

## 交叉表结构（Row 1-4 为产品维度）

| 行     | 内容                                            |
| ------ | ----------------------------------------------- |
| Row 1  | 代码（产品编码，如 FC0000018-1）                |
| Row 2  | 品名（产品名称，如 芥末章鱼）                   |
| Row 3  | 规格（包装规格，如 70g/盒）                     |
| Row 4  | 入库重量（完工产品维度，单位 kg）               |
| Row 5  | 物料耗用表头：物料编码 / 耗用材料 / 数量 / 金额 |
| Row 6+ | 具体物料消耗明细                                |

**交叉表规则**：每行=物料，每2列=一个产品的（数量 + 金额）

## SKU 基础信息

同一列中：

- 代码 + 品名 + 规格 = 一个 SKU

**工序与产品代码规则：**

| 工序   | 产品代码格式                 | 示例        |
| ------ | ---------------------------- | ----------- |
| 小包装 | 代码 + "-1"                  | YQ0000004-1 |
| 大包装 | 公司真正的成品代码（无后缀） | YQ0000004   |

**注意**：询问 YQ0000004 产品成本时，默认指两道工序（包材+大包装）总成本。

## 物料编码前缀

| 前缀 | 类别               | 出成率统计 |
| ---- | ------------------ | ---------- |
| YS00 | 原料类（海鲜为主） | ✅ 计入    |
| YHTL | 鱼糜类             | ✅ 计入    |
| ZJ00 | 佐料类             | ❌ 排除    |
| FD00 | 辅料类             | ❌ 排除    |
| BS00 | 包材类             | ❌ 排除    |
| FS00 | 辅材类             | ❌ 排除    |
| FY00 | 辅助材料           | ❌ 排除    |

**小包装 Sheet 首字母分类：**

| 首字母 | 类别       |
| ------ | ---------- |
| Y      | 原料       |
| F 或 Z | 辅料       |
| B      | 小包装包材 |

**大包装包材**：以字母 B 开头，数量单位为个

## 成本计算逻辑

```
小包装 1kg 成本 = 小包装工序成本 ÷ 小包装工序入库重量
大包装 1kg 成本 = 大包装工序成本 ÷ 大包装工序入库重量
总 1kg 成本 = 小包装 1kg 成本 + 大包装 1kg 成本
SKU 单品成本 = 总 1kg 成本 × 规格（kg）
```

**注意**：不宜将小包装+大包装工序成本直接相加（两工序完工入库重量未必一致）。

## 出成率（仅适用小包装工序原料）

```
出成率 = 小包装完工入库重量 ÷ 该原料大类汇总耗用数量
```

- ✅ 仅适用于原料（YS / YHTL 前缀）
- ❌ 不适用于辅料（FS/ZJ/FD）和包材（BS）
- **正确计算方式**：先去重（同一 SKU + 原料大类下的多物料耗用求和），再计算 `入库 ÷ 汇总耗用`，每个 (SKU, 原料大类) 唯一一行
- **错误方式**：逐行独立计算（会导致同一 SKU + 原料大类出现多行，出成率各不相同）
- 按「耗用材料」口径汇总（非物料编码）

## 使用流程

1. 解析 BOM xlsx（使用 scripts/bom_reader.py）
2. 计算 1kg 成本（使用 scripts/cost_calculator.py）
3. 计算出成率（使用 scripts/yield_rate.py 或 yield_rate_batch.py）
4. 输出飞书表格格式

## Scripts

### bom_reader.py

```bash
python3 scripts/bom_reader.py --month 2026-01 --sheet small
python3 scripts/bom_reader.py --month 2026-01 --sheet large
```

### cost_calculator.py

```bash
python3 scripts/cost_calculator.py --month 2026-01 --sku YQ00109XS --process both
python3 scripts/cost_calculator.py --month 2026-01 --sku YQ00109XS --process small
python3 scripts/cost_calculator.py --month 2026-01 --sku YQ00109XS --process large
```

### yield_rate.py

```bash
python3 scripts/yield_rate.py --month 2026-01 --sku YQ00109XS --mode material
python3 scripts/yield_rate.py --month 2026-01 --sku YQ00109XS --mode large_category
```

### yield_rate_batch.py（推荐优先使用）

批量生成出成率合并表（表一 + 表二），内置正确算法（先去重后计算），自动读取映射表输出按原料大类口径。

```bash
# 生成当月出成率合并表
python3 scripts/yield_rate_batch.py --month 2026-04

# 跳过原料大类映射（仅生成按耗用材料口径）
python3 scripts/yield_rate_batch.py --month 2026-04 --no-mapping

# 指定输出路径
python3 scripts/yield_rate_batch.py --month 2026-04 --output /path/to/output.xlsx
```

**输出：**

- Sheet 1: `出成率（按耗用材料）` — SKU / 品名 / 规格 / 入库重量 / 原料名称 / 耗用数量 / 出成率
- Sheet 2: `出成率（按原料大类）` — SKU / 品名 / 规格 / 入库重量 / 原料大类 / 出成率
- Sheet 3: `汇总` — 月份 / SKU 总数 / 行数统计

**算法说明（关键）**：

1. 过滤出原料（YS/YHTL 前缀）
2. 按 (SKU, 原料大类) 分组，汇总耗用
3. 出成率 = 入库重量 ÷ 汇总耗用（每个组合唯一一行）
4. 若无映射匹配，原料名称即作为大类

**映射表匹配规则：**
取 `media/inbound/原料大类映射表*.xlsx` 中 st_mtime 最新的文件，适用于当月及历史所有月份。

## References

- 详细 BOM 结构说明：见 `references/bom_structure.md`
- 原料大类映射：见 `references/bom_structure.md`
