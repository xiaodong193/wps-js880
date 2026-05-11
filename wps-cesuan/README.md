# WPS 租金测算系统

> 基于 WPS Office JSA（JavaScript API）的专业财务计算系统

## 版本信息

- **系统版本**: 3.2026.5.1
- **更新日期**: 2026-05-01
- **作者**: 徐晓冬
- **架构**: 配置驱动 + 模块化设计 + JSA880框架优先
- **JSA880框架版本**: >= 3.9.4
- **环境要求**: WPS Office（Windows/macOS），内置V8引擎（支持ES6-ES2019）

---

## 目录

- [项目简介](#项目简介)
- [文件清单与加载顺序](#文件清单与加载顺序)
- [系统架构](#系统架构)
- [模块说明](#模块说明)
- [核心概念](#核心概念)
- [快速开始](#快速开始)
- [开发指南](#开发指南)

---

## 项目简介

基于 WPS Office JSA 的专业租赁测算系统，用于租赁业务的租金测算、现金流量表生成等核心财务功能。

### 核心特性

- **配置驱动**: 所有配置集中在 CONFIG 对象中，修改配置无需改动业务逻辑
- **模块化设计**: 清晰的职责划分，按加载顺序依赖注入
- **八种还款方式**: 等额本息（后付/先付）、等额本金（按天/按期）、本金比例（按天/按期）、按期付息、一次性还本付息
- **JSA880框架优先**: 数据处理优先使用 Array2D/DateUtils/RngUtils 框架函数，减少手写循环
- **R1C1 公式**: 使用相对引用公式，可复制粘贴
- **三期模式**: 区分首期、中间期、末期的不同计算逻辑
- **撤销/重做**: 统一撤销管理器支持操作回退
- **安全备份**: 数据修改前自动备份到隐藏工作表，关键操作 MsgBox 确认

---

## 文件清单与加载顺序

> ⚠️ **必须严格按照以下顺序加载，否则会出现依赖错误**

```
// ═══ 第零层：框架层（无业务依赖，必须最先加载） ═══
JSA880.js                    // JSA880框架（Array2D, DateUtils, RngUtils, IO, SuperMap）

// ═══ 第一层：基础层（依赖JSA880） ═══
mShared_constants.js        // 共享常量（颜色、字体、备份/快捷函数）

// ═══ 第二层：基础设施层（依赖基础层） ═══
mParameterManager.js        // 参数管理器（日期转换已集成DateUtils.fromExcelDate）
mErrorHandler.js            // 统一错误处理器
mUndoManager.js             // 撤销管理器

// ═══ 第三层：核心业务层（依赖基础设施层） ═══
mInitialization.js          // 初始化模块
mRentalCalculation.js       // 租金测算核心
mCashFlowGenerator.js       // 现金流量表生成器

// ═══ 第四层：主系统（依赖核心业务层） ═══
mMain.js                    // 主入口，工作流协调（集成自动备份+MsgBox确认）

// ═══ 第五层：扩展模块（按需加载） ═══
m调息.js                     // 利率调整模块（v3.0配置驱动，8种还款方式）
m银行承兑汇票模块.js          // 银行承兑汇票
m直租.js                     // 直租模块
m货币网利率更新.js            // 利率更新
m启动校验.js                  // 启动校验模块（原 m加载项.js）
```

---

## 系统架构

```
┌─────────────────────────────────────────────────────┐
│                    mMain.js（主入口）                  │
│              clsRentCalculationSystem                │
│         计算main() / 清除() / 调期() / ...           │
└──────────────┬──────────────────────┬───────────────┘
               │                      │
       ┌───────▼───────┐    ┌────────▼────────┐
       │ mInitialization│    │ mRentalCalculation│
       │ 参数区域初始化  │    │ 租金测算核心      │
       │ clsRentCalc-   │    │ clsRentalConfig   │
       │ ulationFillin- │    │ clsFormulaGenerator│
       │ Area           │    │ clsStyleManager   │
       └───────┬───────┘    │ clsRentalCalc-    │
               │            │ ulation           │
               │            │ clsRentalTable-   │
               │            │ RowManager        │
               │            └────────┬──────────┘
               │                     │
    ┌──────────▼─────────────────────▼──────────┐
    │           mCashFlowGenerator               │
    │         现金流量表 + 综合利率一览            │
    └──────────────────┬────────────────────────┘
                       │
    ┌──────────────────▼────────────────────────┐
    │         mParameterManager                  │
    │    参数管理器（所有模块共享，中心化配置）     │
    └──────────────────┬────────────────────────┘
                       │
    ┌──────┬───────────▼──────────┬──────────────┐
    │mShared│ mErrorHandler        │ mUndoManager │
    │_const-│ 统一错误处理          │ 撤销管理器   │
    │ants.js│                      │              │
    │常量/   │                      │              │
    │工具函数│                      │              │
    └──────┴──────────────────────┴──────────────┘
```

### 依赖关系图

```
mShared_constants ← mParameterManager ← mInitialization
                  ← mErrorHandler     ← mRentalCalculation ← mMain
                  ← mUndoManager      ← mCashFlowGenerator
```

---

## 模块说明

### mShared_constants.js — 共享常量模块

**职责**: 统一管理所有常量和工具函数

| 导出 | 类型 | 说明 |
|------|------|------|
| `XL` | Object | Excel/WPS 内置常量映射（对齐、边框、方向等） |
| `FORMAT_*` | string | 数字格式常量（Standard, Integer, Date 等） |
| `FONT_*` | string/number | 字体常量 |
| `COLOR_*` | number | 颜色常量（RGB 值） |
| `TABLE` | Object | 表格结构常量 |
| `ERROR_LEVELS` | Object | 错误级别 |
| `RGB(r,g,b)` | Function | RGB 颜色转换 |
| `应用格式(rng, type)` | Function | 应用数字格式 |
| `设置表格样式(rng)` | Function | 设置表格基本样式 |
| `设置字体样式(rng, opts)` | Function | 设置字体样式 |
| `设置背景颜色(rng, color)` | Function | 设置背景颜色 |
| `设置边框(rng, opts)` | Function | 设置边框 |
| `arrDataFromRngExtended(sheet, row, headers)` | Function | 增强型数组读取 |

### mParameterManager.js — 参数管理器

**职责**: 管理所有测算参数的读写、验证和缓存

| 核心方法 | 说明 |
|----------|------|
| `Initialize(sheetName)` | 初始化参数管理器 |
| `val(paramName)` | 获取参数值 |
| `addr(paramName, format)` | 获取参数单元格地址 |
| `param(paramName)` | 获取参数完整配置 |
| `ValidateParameter(paramName)` | 验证参数值 |

### mRentalCalculation.js — 租金测算核心（约2600行）

**职责**: 租金测算的配置、公式生成、样式管理、数据写入

包含 4 个内部类：

| 类名 | 职责 |
|------|------|
| `clsRentalConfig` | 配置管理（列定义、还款方式、合计行规则） |
| `clsFormulaGenerator` | 公式生成（6种还款方式的 R1C1 公式） |
| `clsStyleManager` | 样式管理（表头、格式、边框、颜色） |
| `clsRentalCalculation` | 主类（协调以上模块） |
| `clsRentalTableRowManager` | 行管理（插入/删除行，继承自主类） |

### mCashFlowGenerator.js — 现金流量表生成器

**职责**: 生成现金流量表和综合利率一览表

### mInitialization.js — 初始化模块

**职责**: 初始化工作表参数区域（标题、参数单元格、数据有效性、样式）

### mMain.js — 主入口

**职责**: 工作流协调，提供全局快捷调用函数

| 快捷函数 | 说明 |
|----------|------|
| `计算main()` | 执行完整测算流程（清除→租金表→现金流量表） |
| `清除()` | 清除原有数据 |
| `调期()` | 清除→生成表→生成月间隔 |
| `初始化系统()` | 输出系统信息 |
| `copySht()` | 复制工作表 |
| `调1期()` | 调整首末期自定义支付日 |
| `调2_每期适用利率()` | 使用每期适用利率生成测算表 |
| `调3_复杂融资租赁()` | 复杂融资租赁场景 |
| `执行Main()` | 自动填入参数并执行完整测算 |

### mUndoManager.js — 撤销管理器

**职责**: 提供撤销/重做功能，支持命令模式

### mErrorHandler.js — 统一错误处理器

**职责**: 统一错误分类、分级日志、错误统计

---

## 核心概念

### 配置驱动架构

系统采用配置驱动设计，核心思想是将配置与业务逻辑分离：

```javascript
// 配置对象（修改这里即可调整行为）
const WORKFLOW_CONFIG = {
    steps: { clear: {...}, generateRentTable: {...}, generateCashFlow: {...} },
    system: { moduleName: "租金测算系统", version: "2.2026.1.30" },
    errorHandling: { showAlert: true, continueOnError: false }
};

// 业务逻辑（无需修改）
class clsRentCalculationSystem { ... }
```

### 三期公式模式

租金测算公式区分三种期次类型：

| 类型 | 数组索引 | 说明 |
|------|----------|------|
| 首期 | `arr[1]` | 第1期，使用绝对起始日期 |
| 中间期 | `arr[2]` | 第2期到倒数第2期，使用相对引用 |
| 末期 | `arr[3]` | 最后一期，本金用差值补齐 |

### R1C1 公式引用

系统使用 R1C1 格式的公式，便于相对引用：

```javascript
// 绝对引用（指向 B5 单元格的利率）
arr[1][3] = `=ROUND(-PMT(R5C2/R11C2, R10C2, R4C1, 0), 2)`;

// 相对引用（指向上一行同列）
arr[2][1] = "=R[-1]C+1";
```

### 八种还款方式

| 还款方式 | 公式方法 | 说明 |
|----------|----------|------|
| 等额本息（后付） | `generateEqualPaymentFormulas` | PMT/PPMT 函数 type=0 |
| 等额本息（先付） | `generateEqualPaymentAdvanceFormulas` | PMT/PPMT type=1 |
| 等额本金（按天计息） | `generateEqualPrincipalDailyInterestFormulas` | 本金/总期数，利息按天 |
| 等额本金（按期计息） | `generateEqualPrincipalPeriodicInterestFormulas` | 本金/总期数，利息按期 |
| 本金比例（按期计息） | `generatePrincipalRatioPeriodicInterestFormulas` | 本金×比例，利息按期 |
| 本金比例（按天计息） | `generatePrincipalRatioDailyInterestFormulas` | 本金×比例，利息按天 |
| 按期付息 | `generateFormulasForMethod` (配置驱动) | 每期仅付利息，末期还本 |
| 一次性还本付息 | `generateFormulasForMethod` (配置驱动) | 末期一次性还本+全部利息 |

> v3.0 新增：按期付息、一次性还本付息。全部8种方式通过 `REPAYMENT_CONFIGS` 配置表驱动，新增还款方式仅需添加配置项。

---

## 快速开始

### 1. 在 WPS 中加载

1. 打开 WPS 表格
2. 按 `Alt+F11` 打开宏编辑器
3. 按照加载顺序依次导入 JS 文件
4. 运行 `执行Main()` 进行自动测算

### 2. 手动操作

```javascript
// 初始化系统
初始化系统();

// 执行完整测算（自动设置参数）
执行Main();

// 或分步操作
清除();          // 步骤1：清除数据
计算main();      // 步骤2：生成租金表 + 现金流量表
```

### 3. 使用每期适用利率

```javascript
// 生成使用每期适用利率的测算表
调2_每期适用利率();

// 修改 M 列利率值后，重新计算
Application.Worksheets("1租金测算表V1").Calculate();
```

---

## 开发指南

### 添加新的还款方式

在 `m调息.js` 的 `REPAYMENT_CONFIGS` 中添加配置项即可（配置驱动，无需修改公式生成器）：

```javascript
REPAYMENT_CONFIGS['新还款方式名称'] = {
    pmtType: 0,                    // PMT type: 0=后付, 1=先付
    interestBasis: 'periodic',     // 'periodic'=按期计息, 'daily'=按天计息
    usesPrincipalRatio: false,     // 是否使用本金比例
    firstDateFormula: function(p) { return '=EDATE(...'; },
    firstRentFormula: function(p) { return '=ROUND(...'; },
    firstPrincipalFormula: function(p) { return '=...'; },
    firstInterestFormula: function(p) { return '=...'; },
    middleInterestFormula: function(p) { return '=...'; },
    lastPrincipalFormula: function(p) { return '=...'; }
};
```
`generateFormulasForMethod()` 将根据配置自动生成完整的首/中/末期公式数组。

### 添加新的参数

1. 在 `mParameterManager.js` 的参数配置中添加参数定义
2. 在 `mInitialization.js` 的 `INITIALIZATION_CONFIG` 中添加区域配置
3. 如需数据有效性，添加 `dataValidation` 配置

### 注意事项

- **全局作用域**: WPS JSA 不支持 ES6 的 `export` 语法，所有定义都在全局作用域中
- **加载顺序**: 严格遵守文件加载顺序，依赖关系不可颠倒
- **单元格引用**: 公式统一使用 R1C1 格式，通过参数管理器获取单元格地址
- **错误处理**: 使用 `try-catch` 包裹所有关键操作，错误通过 `alert` 通知用户
- **控制台日志**: 大量 `console.log` 用于调试，生产环境可通过 `ERROR_CONFIG.currentLogLevel` 控制输出级别

---

## 文档索引

| 文档 | 说明 |
|------|------|
| `模块清单文档.md` | 模块功能清单 |
| `加载顺序说明.md` | 详细的加载顺序说明 |
| `每期适用利率使用指南.md` | 每期适用利率功能的使用方法 |
| `统一撤销管理器使用指南.md` | 撤销管理器的使用方法 |
| `依赖注入重构维护指南.md` | 依赖注入相关的维护说明 |