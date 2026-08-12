# 导出标准报表误粘死表内公式 — 根因与修复方案

日期：2026-08-11
涉及文件：`QSBar/ExportCommands.cs`

## 现象

「导出标准报表」把与外部链接无关的正常公式也粘死成数值，例如：

- `=F61*D61` —— 纯本表引用
- `='J_Drainage Channel'!M30` —— 表内跨工作表引用

两者引用的列都没有被隐藏、也没有被删除，粘死没有任何必要。

## 根因（已确证）

`ConvertReferencesToHiddenColsAndNames` 的阶段1 用**裸子串匹配**判断公式是否引用了将被删除的问题名称：

```csharp
foreach (string name in problemNames)
{
    if (formula.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
    {
        needsConversion = true;
        break;
    }
}
```

`problemNames` 来自 `wb.Names` 中所有取值含 `#REF!` / `:\` / `\\` / `#N/A` 的定义名称。

BOQ 文件通常由多来源清单拼合而来，会累积大量残缺名称，其中不乏**极短**的。子串匹配对短名称是灾难性的：一个名为 `A` 的坏名称，会让任何含字母 a/A 的公式全部中招——包括 `=SUM(...)` 里的列字母、跨表引用中的工作表名。

### 实测数据

样本 `.memory/Schedule of Prices.xlsx`（真实工程文件）：

| 指标 | 数值 |
|---|---|
| 定义名称总数 | 284 |
| 其中被判为 problemNames | 130 |
| 长度 ≤ 3 的 problemNames | `A` `AZ` `BV` `CP` `IC` `JY` `ODH` `SC` `UI` `VC` `VX` `ZZZ` `__` `_t1` `_t2` `ac` `kj` `vur` `yas` 等 31 个 |
| 工作表内公式单元格总数 | 928 |
| **被误粘死** | **144（15.5%）** |
| 其中由单个名称 `A` 命中 | 143 |

误杀样本：

- `=+SUM(AA8:AA31)` —— 本表求和，因列字母 `AA` 含 `A` 而死
- `='B101 Preamble'!B1` —— 跨表引用，因工作表名 `Preamble` 含 `a` 而死

用户报告的 `='J_Drainage Channel'!M30` 属同一类：工作表名 `Drainage` 里的 `a` 命中了名称 `A`。

### `=F61*D61` 的路径

代码里只有两条粘死路径。阶段2（隐藏列）已排除——D、F 列未隐藏也未被删除。阶段1 只有两个触发条件，`externalRefRegex` 要求公式文本含方括号而 `=F61*D61` 没有，因此**只可能是 problemNames 子串匹配**，即该工作簿中存在名为 `D` 或 `F` 的坏名称（单字母名称在 Excel 中合法，仅 `R` 和 `C` 保留）。与 `A` 的情形完全同构，同一处修复即可覆盖。

补充：F61 自身含外部引用、被粘死为数值是**正确行为**。粘死后 `=F61*D61` 仍能正常计算，不需要跟着死；代码中也不存在这种传播逻辑。

## 修复方案

对阶段1 的名称匹配做两处收窄。两处缺一不可。

### 修复点 1：先剥离公式中的引号字面量

工作表名和文本常量不是名称引用，不该参与名称匹配。匹配前先把 `'...'` 和 `"..."` 整体替换掉（Excel 中引号自身以 `''` / `""` 转义）：

```csharp
private static readonly Regex QuotedLiteral = new Regex(
    @"'(?:[^']|'')*'|""(?:[^""]|"""")*""",
    RegexOptions.Compiled);

string scrubbed = QuotedLiteral.Replace(formula, "''");
```

这一条单独就能救回 `='J_Drainage Channel'!M30` 和 `='B101 Preamble'!B1`。

### 修复点 2：名称匹配加标识符边界

名称引用必须是完整标识符，前后不能紧邻名称合法字符。Excel 名称的合法字符是字母、数字、下划线、点、反斜杠（.NET 的 `\w` 已覆盖 Unicode 字母数字下划线，含中日韩）：

```csharp
private const string NameBoundary = @"[\w.\\]";
// (?<![\w.\\]) + Regex.Escape(name) + (?![\w.\\])
```

这一条救回 `=+SUM(AA8:AA31)` 这类——`AA` 中的两个 `A` 前后都紧邻字母，不构成名称引用。

### 性能

现行实现是 `problemNames.Count × 公式数` 次 `IndexOf`。130 个名称 × 928 公式已是 12 万次；上万行的大表会到千万级，改成逐个正则会更慢。

实现时把所有 problemNames 合并成**一个**预编译正则，一次匹配：

```csharp
string pattern = @"(?<![\w.\\])(?:" +
    string.Join("|", problemNames.Select(Regex.Escape)) +
    @")(?![\w.\\])";
var nameRegex = new Regex(pattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
```

名称按长度降序排列，避免交替分支的短名称抢先匹配。`problemNames` 为空时跳过整个检查。

### 验证结果

在 `.memory/Schedule of Prices.xlsx` 上以脚本复刻现行逻辑与修复逻辑对比（脚本见 `.memory/temp/scripts/verify_fix.py`）：

```
formula cells        : 928
killed BEFORE fix    : 144
killed AFTER  fix    : 0
rescued              : 144
names still matching : []
```

回归检查——真正引用坏名称的公式必须仍被粘死，否则删名称后会变 `#NAME?`：

| 公式 | 修复前 | 修复后 |
|---|---|---|
| `=A*2` | 命中 `A` | 命中 `A` |
| `=SUM(A)` | 命中 `A` | 命中 `A` |
| `=IC+1` | 命中 `IC` | 命中 `IC` |
| `=ac` | 命中 `A`,`ac` | 命中 `ac` |
| `=B1+A` | 命中 `A` | 命中 `A` |
| `=F61*D61` | 未命中 | 未命中 |
| `='J_Drainage Channel'!M30` | 命中 `A` | 未命中 |

该杀的一个没放过，不该杀的一个没误伤。

## 次要缺陷（代码审查发现，未在样本中复现）

以下两项不是本次问题的病因，但同属精度缺陷，建议一并修。若要控制改动面，可只做上面的主修复。

### A. 外部引用正则过宽

```csharp
var externalRefRegex = new Regex(@"\[[^\]]*\]", ...);
```

匹配任意方括号。Excel 表格（ListObject）的结构化引用同样使用方括号——`=[@数量]*[@单价]`、`=SUM(表1[金额])`——会被误判为外部工作簿引用而粘死。

样本文件公式中不含任何方括号 token，未复现。若 BOQ 模板启用了 Excel 表格功能就会触发。

收窄为真实的外部引用形态：`\[\d+\]`（已解析的索引式）与 `\[[^\[\]]*\.xl[a-z]*\]`（路径式），并显式排除 `[@`。

### B. 阶段2 跨表引用只跳过紧邻 `!` 的第一个引用

```csharp
if (match.Index > 0 && formula[match.Index - 1] == '!')
    continue;
```

对 `='其他表'!M30` 有效，但对区域引用 `=SUM('其他表'!M30:M40)` 无效——`M40` 前面是 `:` 而非 `!`，会被当作本表引用。若本表 M 列恰好隐藏，该跨表公式会被误粘死。

改为整体识别跨表引用块（`'表名'!A1` 或 `表名!A1:B2`，含区域的后半段），从公式中整体剔除后再检查剩余的本表引用。

### C. 相关但独立的问题：跨表引用隐藏列漏粘死

阶段2 只处理「本表公式引用本表隐藏列」。若 Sheet2 有公式引用 Sheet1 的隐藏列，Sheet1 的列被 `DeleteHiddenColumns` 删除后，Sheet2 的公式会变 `#REF!`。

这是漏粘死（与本次的多粘死方向相反），属独立问题，本方案不处理，仅记录。

## 验证方法

修复后按 `CLAUDE.md` 流程编译 Release 并注册 COM，然后：

1. 用 `.memory/Schedule of Prices.xlsx` 执行「导出标准报表」，导出文件中 `=+SUM(AA8:AA31)`、`='B101 Preamble'!B1` 应保持为公式
2. 用报告问题的原文件执行导出，`=F61*D61` 与 `='J_Drainage Channel'!M30` 应保持为公式
3. 在测试文件中手工建一个取值为 `#REF!` 的定义名称 `A`，写入 `=A*2`，导出后该公式应被粘死为值（不能留 `#NAME?`）
4. 含真实外部链接的公式导出后仍应被粘死为值
5. 隐藏一列并写入引用该列的公式，导出后应粘死为值且该列被删除
