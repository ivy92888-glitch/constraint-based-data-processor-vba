# Constraint-Based Data Processor — VBA Toolkit

> 数据清洗 + 约束收敛 + 实体拆分 + 规则路由，一个 VBA 文件搞定

![](./examples/preview.svg)

---

## 📖 中文版（故事感）

有一天你打开一张"发货发货数据"表，往下翻三行就看到了一堆让你皱眉的问题：

- **同一行里混了 6 个实体编号**——`实体编号-1234567-001/实体编号-1234567-002/实体编号-1234567-003` 全堆在 C 列里，每个本来应该是独立的一行
- **仓库编码带了脏后缀**——`WH-EAST-30001` 里的 `-30001` 是邮编后缀不是仓库编码，直接拿去匹配当然会失败
- **邮编格式不统一**——`30001` 是 5 位（正确），但 `6380` 少了一位（应该补前导零成 `06380`）
- **PO 号数量和 实体编号 号对不上**——C 列 3 个编号，D 列只有 2 个 PO 号。拆了之后肯定会有一行没有 PO，后面申报会崩
- **总体积不在目标区间**——汇总 67.8（小了），目标是 68.0~68.1。67.8 预算分配订不到，必须精准调进去
- **快递单号没自动识别**——DHL/UPS/DPD 单号格式不一样，但你总不能每次 Ctrl+F 对比吧

手工处理这些问题：100 行数据 → 2-3 小时。因为每一步都有连锁反应：**拆一行 实体编号，要同时按件数比例分配重量和体积，还要二次汇总吸收四舍五入偏差，再去调下一行……** 还有那个 68.0~68.1 的约束收敛，你想不出怎么同时调所有行的长宽高才能让总体积刚好落进去。

`constraint-based-data-processor-vba` 把这些问题做成了 **5 个独立但可一键串联的 VBA 模块**。运行一次，原始表变成干净的标准发货数据表。

---

## 🧩 五个模块，各自解决一个具体问题

### 模块 1：清洗标准化（Clean & Normalize）

**干的事**：仓库编码去后缀 → 邮编补前导零 → 从 Lookup 表自动补全缺失编码。

**数字例子**：
> 仓库编码 `WH-EAST-30001` → 识别后缀 `-30001` → 删掉 → `WH-EAST`
> 
> 邮编 `6380` → 检测到不足 5 位 → 补前导零 → `06380`（强制文本格式避免 Excel 去掉前导零）

### 模块 2：实体拆分 + 比例重分配（Split & Redistribute）

**干的事**：一行里多个 `"/"` 分隔的标识符 → 拆成多行 → 重量/体积按件数比例重算 → 二次汇总吸收四舍五入偏差。

**数字例子**：
> 原始一行：C 列 `实体编号-1234567-001/实体编号-1234567-002/实体编号-1234567-003`（3 个标识符，各 1 件）
> 
> 拆成 3 行，重量/体积均分。拆分后每行单独汇总校验，把四舍五入产生的 0.01-0.03kg 偏差吸收掉。

### 模块 3：逻辑校验 + 防呆（Validate & Guard）

**干的事**：拆分前先做校验——C 列和 D 列的标识符数量必须一致。不一致 → 标红警告 + **停止拆分**（宁可停住让人工介入，也不要拆出一行没有 PO 号的坏数据）。

### 模块 4：目标约束收敛算法（Constraint Convergence — 最核心的算法）

**干的事**：**把总体积精准调到 `[VOLUME_MIN, VOLUME_MAX]` 区间**（默认 68.0 ~ 68.1 CBM）。

**核心思路**（迭代法，最多 200 次防止不收敛）：

```
当前总体积 > 68.1？
  └─ 是 → direction = -1（收缩）→ 所有行的 L/W/H 同时各减 Δ（默认 0.01）
当前总体积 < 68.0？
  └─ 是 → direction = +1（扩张）→ 所有行的 L/W/H 同时各加 Δ（默认 0.01）
每次调整后：
  1. 重算每行体积 = L × W × H
  2. 汇总总 → 检查是否进入 [68.0, 68.1]
  3. 同时约束：L/W/H 都不能低于 MIN_DIMENSION（默认 0.1）
```

**数字例子**：
> 当前总体积 = **68.35 CBM**（超出上限 68.1）
> 
> → direction = -1，Δ = -0.01
> → 迭代 10 次后 → 总体积 = **68.08 CBM** ✅ 进入区间
> → 所有行的 L/W/H 同时减少了 0.10

**为什么同时调所有行的三个维度？** 这样每行的比例关系保持不变（长方体形状不变），只是整体缩放。

### 模块 5：规则路由 + 模式匹配（Route & Classify）

**干的事**：快递单号正则匹配 → 自动识别哪家快递 → 写进路由列。

**规则表**（初始化时加载，按顺序匹配，第一个命中即返回）：

| 正则模式 | 路由编码 |
|---|---|
| `^1Z[A-Z0-9]{16}$` | UPS |
| `^\d{12,14}$` | DHL |
| `^[A-Z]{2}\d{9}$` | DPD |
| `^[A-Z0-9]{8,}$` | OTHER |

---

## 🔥 适用场景（至少 7 种不同行业）

### 场景 1 — 物流发货发货数据（原始场景）
发货表多个 实体编号 单号混一行 → 拆分 → 体积精准收敛 → 快递路由自动识别。

### 场景 2 — 制造业 BOM 物料拆分
一个组件行里有多个子零件编号（"/" 分隔）→ 按 BOM 比例拆成多行 → 重量按比例重算。

### 场景 3 — 财务发票校验
同一张发票里多个项目编号（"," 分隔）→ 拆成多行 → 金额按比例分摊 → 校验总额不变。

### 场景 4 — HR 工资条拆分
一个工号下多个兼职项目编号（"/" 分隔）→ 拆成多行 → 工时按比例分配 → 工资按比例重算。

### 场景 5 — 广告投放预算收敛
多组广告的花费汇总后需要精准落在预算区间 `[98, 100]` 万 → 约束收敛算法让每组花费同比调整。

### 场景 6 — 数据中心服务器资源收敛
多台服务器的 CPU 使用率需要整体落在 `[70%, 75%]` → 约束收敛算法让每台同比加/减分配。

### 场景 7 — 库存实体包裹尺寸标准化
多个实体包裹需要精准填充一个目标容器（内部尺寸 [68.0, 68.1] CBM）→ 约束收敛让每个实体包裹同比缩放。

---

## 🚀 快速上手

### 1. 导入代码
VBA 编辑器（Alt + F11）→ 导入 `src/ConstraintBasedProcessor.bas`。

### 2. 准备工作簿

| 工作表 | 必需列 | 干什么 | 被哪个模块用 |
|---|---|---|---|
| **Input** | A-L（默认布局） | 主数据表（待处理） | 全部 5 个模块 |
| **Lookup** | A-B | 编码补全查找表（可选） | 模块 1 |

> 所有 Sheet 名和列字母都在源码顶部 `Private Const` 里。改完保存即生效。详见 `docs/Configuration.md`。

### 3. 运行

```
一键跑全部 5 个模块:  RunFullPipeline()
单独运行：
  RunModule1() — 清洗标准化
  RunModule2() — 实体拆分
  RunModule3() — 逻辑校验
  RunModule4() — 约束收敛
  RunModule5() — 规则路由
```

### 4. 看结果
弹出提示框告诉你每个模块处理了多少行。约束收敛会显示最终体积是否落进目标区间。

---

## 🎯 约束收敛算法参数一览

| 常量 | 默认值 | 含义 | 什么时候改 |
|---|---|---|---|
| `VOLUME_MIN` | 68.0 | 目标体积区间下限 | 你的业务要求不同区间时 |
| `VOLUME_MAX` | 68.1 | 目标体积区间上限 | 同上 |
| `MIN_DIMENSION` | 0.1 | L/W/H 的最小约束（不允许低于此值） | 你的箱子最小尺寸有硬限制时 |
| 迭代次数上限 | 200 | 防止不收敛时的死循环 | 一般不需要改 |
| 每次调整步长 | 0.01 | 每次迭代调整的厘米数 | 如果收敛太慢可以改大（0.05），太粗改小（0.005） |

---

## ⚠️ 运行前必知

1. **约束收敛会同时修改所有行的 L/W/H 三个维度**——建议运行前备份
2. **实体拆分会增加行数**——每多一个 "/"+1 行
3. **逻辑校验会把不一致的行标红**——标红的行不会被拆分，建议人工处理后再跑
4. **Lookup 表是可选的**——没有也能跑，只是缺少编码自动补全这一步

---

## 📦 文件清单

```
constraint-based-data-processor-vba/
├── src/ConstraintBasedProcessor.bas   ← 唯一代码文件，5 模块 + 12 函数
├── examples/sample-data.xlsx          ← 虚构演示工作簿
├── docs/Configuration.md              ← 常量说明 + 表结构模板
├── README.md
├── LICENSE                            ← MIT
├── .gitignore
└── .gitattributes
```

---

## 📅 版本

- **v1.0.0** — 2026-11-15 — 初始发布：清洗标准化 + 实体拆分 + 逻辑校验 + 约束收敛算法 + 规则路由

---

## 🔗 其他项目

同系列仓库（clean-room 通用化 + MIT 许可）：
- [weighted-data-validator](https://github.com/ivy92888-glitch/weighted-data-validator) — 加权评分 + 规则引擎 + 关键词优先级解析
- [cross-sheet-data-aggregator](https://github.com/ivy92888-glitch/cross-sheet-data-aggregator) — 跨表合并 + 三级去重 + 实体标识符关联
- [smart-data-governance-vba](https://github.com/ivy92888-glitch/smart-data-governance-vba) — 去重合并 + 加权评分校准 + 智能字段清洗

---

## Quick Summary (English)

**constraint-based-data-processor-vba** is a single-file Excel VBA toolkit with
five data-transformation modules for row-level processing:

1. **Clean & Normalize** — Strip entity-code suffixes, pad postal codes to
   fixed length with leading zeros, backfill missing codes from a lookup
   table.
2. **Split & Redistribute** — Split rows with slash-separated identifiers
   into one row each; redistribute weight/volume proportionally by piece
   count; re-aggregate to absorb rounding drift.
3. **Validate & Guard** — Before splitting, cross-check count consistency
   across identifier columns; flag mismatches so operator intervenes.
4. **Constraint Convergence (core algorithm)** — Iterative uniform
   dimension adjustment (length/width/height ±Δ per row) until aggregate
   volume lands within a target range `[VOLUME_MIN, VOLUME_MAX]`.
   Respects MIN_DIMENSION floor. Max 200 iterations guard against
   non-convergence.
5. **Route & Classify** — Regex pattern-match tracking numbers to carrier
   codes (DHL/UPS/DPD/OTHER). First matching rule wins.

## Maintenance Note

This README is intentionally **Chinese-primary**. The English section is
a concise summary only. All parameters and constants live in source code —
**do not copy numbers into documentation**. Source is the single source
of truth.

---

> This README follows a Chinese-first style (故事感开头 → 人话功能总结 → 多场景覆盖 → 具体数字例子 → 砍边角功能) as preferred by the author. For English readers: the Quick Summary section covers the essentials.
