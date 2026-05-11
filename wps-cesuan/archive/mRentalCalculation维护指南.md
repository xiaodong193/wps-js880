# 租金测算系统 V2 维护指南

## 概述

本文件是租金测算系统的 V2 重构版本，完全基于 mParameterManager_v2.js 设计。

## 核心设计原则

1. **配置驱动**：所有配置集中在 clsRentalConfig
2. **单一职责**：每个类只负责一个功能域
3. **消除重复**：公式生成逻辑统一，差异参数化
4. **缓存优先**：充分利用参数管理器的 _addressCache
5. **完善注释**：所有复杂逻辑都有设计意图说明

## 如何添加新还款方式

### 步骤 1：在配置中添加还款方式

位置：`clsRentalConfig._initRepaymentMethods()`

```javascript
"新还款方式": {
    formulaMethod: "generateNewMethodFormulas",  // 对应的公式生成方法名
    usePrincipalRatio: false,                    // 是否使用本金比例列
    needCustomInterval: false,                   // 是否需要自定义月间隔列
    convertColumns: [1, 2],                     // 需要转换为数值的列索引
    clearColumns: [],                            // 需要清除内容的列索引
    headerRange: [1, 10],                       // 表头范围 [起始列, 结束列]
    addFrame: ["A:J"],                          // 需要添加框线的列范围
    extraColumns: []                             // 需要特殊处理的列索引
}
```

### 步骤 2：添加公式生成方法

位置：`clsFormulaGenerator` 类

```javascript
generateNewMethodFormulas() {
    try {
        const params = this._getFormulaParams();
        const arr = this._createFormulaArray();

        // 第1期公式
        arr[1][1] = "1";
        arr[1][2] = `=EDATE(${params.leaseStartDateCell}, ${params.paymentInterval})`;
        // ... 其他列公式

        // 第2期-倒数第2期公式
        arr[2][1] = "=R[-1]C+1";
        // ... 其他列公式

        // 最后一期公式
        arr[3][1] = arr[2][1];
        // ... 其他列公式

        return arr;
    } catch (error) {
        console.error(`[${this.MODULE_NAME}] 新还款方式公式生成失败: ${error.message}`);
        return null;
    }
}
```

### 步骤 3：在公式生成器中添加调用逻辑

位置：`clsFormulaGenerator.generateFormulas()` 方法

该方法会自动根据配置调用对应的公式生成方法，无需修改。

## 如何修改现有功能

### 修改还款方式的配置

**位置**：`clsRentalConfig._initRepaymentMethods()`

**说明**：修改配置对象的属性值

| 配置项 | 说明 | 示例 |
|--------|------|------|
| convertColumns | 需要转换为数值的列（从1开始） | `[1, 2]` 表示第1、2列 |
| clearColumns | 需要清除内容的列（从1开始） | `[12]` 表示第12列（本金比例） |
| headerRange | 需要额外创建表头的列范围 | `[12, 12]` 表示只创建第12列表头 |
| addFrame | 需要添加框线的列范围 | `["A:J", "L:L"]` |
| extraColumns | 需要特殊处理的列（如黄色背景） | `[12]` |

### 修改公式逻辑

**位置**：`clsFormulaGenerator` 中的对应方法

**说明**：修改 `arr[x][y]` 的公式字符串

**示例**：修改等额本息的租金公式

```javascript
// 原公式
arr[1][3] = `=ROUND(-PMT(${params.interestRateCell}/R11C2,${params.totalPeriodsCell},${params.principalCell},0),2)`;

// 修改为（添加手续费）
arr[1][3] = `=ROUND(-PMT(${params.interestRateCell}/R11C2,${params.totalPeriodsCell},${params.principalCell}*1.01,0),2)`;
```

### 修改样式

**位置**：`clsStyleManager` 中的对应方法

**说明**：修改样式设置（颜色、字体、边框等）

**示例**：修改表头背景色

```javascript
colRange.Interior.Color = m_COLOR_BLUE;  // 原蓝色
// 改为
colRange.Interior.Color = this.p.m_COLOR_GREEN;  // 绿色
```

### 修改列定义

**位置**：`clsRentalConfig._initColumnDefinitions()`

**说明**：修改列字母映射

**示例**：将期次列从 A 改为 B

```javascript
cols = {
    PERIOD: 'B',  // 原来是 'A'
    DATE: 'C',    // 原来是 'B'
    // ... 其他列也要相应调整
}
```

**注意**：修改列定义时，必须同步修改 `_initHeaders()` 中的表头顺序。

## 代码导航提示

| 需求 | 搜索关键词 | 位置 |
|------|-----------|------|
| 添加还款方式 | `_initRepaymentMethods` | clsRentalConfig |
| 修改公式 | `generateXxxFormulas` | clsFormulaGenerator |
| 修改样式 | `createTableHeaders`, `applyDataFormat` | clsStyleManager |
| 修改列定义 | `_initColumnDefinitions` | clsRentalConfig |
| 修改行操作 | `insertRows`, `deleteRows` | clsRentalTableRowManager |
| 修改合计行 | `_initTotalRowConfig`, `createTotalRow` | clsRentalConfig, clsStyleManager |

## 与参数管理器的集成

### 访问参数管理器

通过全局变量 `p` 访问参数管理器实例：

```javascript
// 读取参数值
const totalPeriods = this.p.val("TotalPeriods");
const principal = this.p.val("Principal");

// 获取单元格地址（支持缓存）
const cellA1 = this.p.addr("LeaseStartDate");
const cellR1C1 = this.p.addr("InterestRate", "R1C1");

// 访问工作表
const worksheet = this.p.m_worksheet;

// 访问颜色常量
const blueColor = this.p.m_COLOR_BLUE;
const whiteColor = this.p.m_COLOR_WHITE;
```

### 参数管理器的关键属性

| 属性 | 类型 | 说明 |
|------|------|------|
| p.m_worksheet | Worksheet | 目标工作表 |
| p.m_config | Object | 参数配置对象 |
| p._isInitialized | Boolean | 是否已初始化 |
| p._addressCache | Object | 地址缓存 |
| p.RentTableStartRow | Number | 租金测算表起始行 |
| p.CashFlowTablerowStart | Number | 现金流表起始行 |

## 常见维护场景

| 场景 | 修改位置 | 说明 |
|------|----------|------|
| 添加新列 | `_initHeaders`, `_initColumnDefinitions` | 修改表头和列定义 |
| 修改表头文本 | `_initHeaders` | 修改表头数组 |
| 修改列顺序 | `_initHeaders`, `_initColumnDefinitions` | 保持一致 |
| 修改合计行规则 | `_initTotalRowConfig`, `createTotalRow` | 修改合计行逻辑 |
| 修改颜色 | 使用参数管理器的颜色常量 | m_COLOR_xxx |
| 修改公式范围 | `clsFormulaGenerator` 中的公式字符串 | 注意 R1C1 格式 |
| 撤回上一步骤操作 ｜修改刚才操作步骤，回退｜撤销刚才运行的程序｜

## 公式模板说明

### 数组结构

```javascript
arr[formulaType][columnIndex]
```

- `formulaType`: 1=首期，2=中间期，3=末期
- `columnIndex`: 从1开始的列索引

### Excel 公式格式

使用 R1C1 格式的相对引用：

| 格式 | 说明 | 示例 |
|------|------|------|
| `R[-1]C` | 上一行同列 | 引用上期数据 |
| `RC[-1]` | 同行左一列 | 引用左侧列 |
| `R5C3` | 绝对引用（第5行第3列） | 引用固定单元格 |
| `SUM(R1C1:R10C1)` | 求和范围 | 从第1行到第10行的第1列 |

### 公式模板示例

```javascript
// 第1期公式
arr[1][1] = "1";  // 期次
arr[1][2] = `=EDATE(${params.leaseStartDateCell}, ${params.paymentInterval})`;  // 支付日
arr[1][3] = `=ROUND(-PMT(...),2)`;  // 租金

// 第2期公式（相对引用）
arr[2][1] = "=R[-1]C+1";  // 期次 = 上期 + 1
arr[2][2] = "=EDATE(R[-1]C, 6)";  // 支付日 = 上期 + 6个月
```

## 行操作说明

### 插入行流程

1. 验证输入参数
2. 判断是否需要重新生成公式
3. 在指定位置插入空行
4. 重新生成公式（如果需要）
5. 更新受影响行的公式
6. 更新剩余租金余额公式范围
7. 重分配本金比例（如果是本金比例法）
8. 更新 TotalPeriods 参数
9. 同步到工作表
10. 触发重新计算

### 删除行流程

1. 验证输入参数
2. 判断是否需要重新生成公式
3. 从指定位置删除行
4. 重新生成公式（如果需要）
5. 更新受影响行的公式
6. 更新剩余租金余额公式范围
7. 重分配本金比例（如果是本金比例法）
8. 更新 TotalPeriods 参数
9. 同步到工作表
10. 触发重新计算

### 受影响的还款方式

| 还款方式 | 是否需要重新生成公式 | 原因 |
|----------|---------------------|------|
| 等额本息（后付） | 否 | 使用相对引用 |
| 等额本息（先付） | 否 | 使用相对引用 |
| 等额本金（按天计息） | 是 | 公式包含总期数 |
| 等额本金（按期计息） | 是 | 公式包含总期数 |
| 本金比例（按天计息） | 是 | 公式包含总期数 |
| 本金比例（按期计息） | 是 | 公式包含总期数 |

## 注意事项

1. **列定义一致性**
   - 修改列定义时，确保表头和列定义保持一致
   - 表头数组的顺序必须与列字母映射对应

2. **公式生成方法**
   - 添加新还款方式时，确保公式生成方法存在
   - 方法名必须与配置中的 `formulaMethod` 一致

3. **R1C1 格式**
   - 修改公式时，注意 R1C1 格式的相对引用
   - 测试公式确保计算正确

4. **TotalPeriods 参数**
   - 行操作后，必须更新 TotalPeriods 参数
   - 使用 `this.p.m_worksheet.Range(this.p.addr("TotalPeriods")).Value2` 更新

5. **本金比例法**
   - 最后一期使用 SUM 公式调整，确保总和为 100%
   - 重新分配时使用 `100 / 总期数` 计算平均比例

6. **缓存机制**
   - 使用 `p.addr()` 获取单元格地址时，会自动使用缓存
   - 首次访问后，后续访问会直接返回缓存值

## 调试技巧

### 启用详细日志

```javascript
// 在关键位置添加 console.log
console.log(`[${this.MODULE_NAME}] 总期数: ${totalPeriods}`);
console.log(`[${this.MODULE_NAME}] 还款方式: ${repaymentMethod}`);
console.log(`[${this.MODULE_NAME}] 公式数组:`, arrFormula);
```

### 验证公式生成

```javascript
// 生成公式后，检查数组内容
const arr = this.formulaGenerator.generateFormulas("等额本息（后付）");
if (arr && arr[1] && arr[1][3]) {
    console.log(`第1期租金公式: ${arr[1][3]}`);
}
```

### 检查工作表数据

```javascript
// 检查特定单元格的值
const cellValue = this.p.m_worksheet.Range("A1").Value2;
console.log(`单元格A1的值: ${cellValue}`);

// 检查单元格公式
const cellFormula = this.p.m_worksheet.Range("A1").Formula;
console.log(`单元格A1的公式: ${cellFormula}`);
```

## 性能优化建议

1. **批量写入**
   - 使用 `Range.Value2` 批量写入数据，而非逐个单元格写入
   - 数据数组转换为二维数组后一次性写入

2. **利用缓存**
   - 使用 `p.addr()` 而非直接调用转换方法
   - 地址缓存会显著提升性能

3. **减少重新计算**
   - 批量操作完成后，再触发工作表重新计算
   - 避免在循环中频繁调用 `Calculate()`

## 错误处理

### 常见错误

| 错误信息 | 原因 | 解决方法 |
|----------|------|----------|
| "不支持的还款方式" | 还款方式名称拼写错误 | 检查 `_initRepaymentMethods` 中的配置 |
| "公式模板生成失败" | 公式生成方法返回 null | 检查公式生成方法的逻辑 |
| "列索引越界" | 列索引超出数组范围 | 检查列定义和数组长度 |
| "数据数组无效" | 数组为空或格式错误 | 检查数据数组的初始化和格式 |

### 错误处理模式

```javascript
try {
    // 执行操作
    const result = this.someMethod();
    return result;
} catch (error) {
    console.error(`[${this.MODULE_NAME}] 操作失败: ${error.message}`);
    // 返回默认值或抛出更具体的错误
    return null;
}
```

## 测试文档

### 测试概述

本章节提供了完整的测试用例和测试步骤，用于验证租金测算系统 V2 的所有功能。

### 测试环境准备

#### 前置条件
1. ✅ 确保已加载 `mShared_constants_v2.js`
2. ✅ 确保已加载 `mParameterManager_v2.js`
3. ✅ 确保存在名为 "1租金测算表V1" 的工作表
4. ✅ 确保参数区域已正确初始化

#### 测试数据准备

```javascript
// 基础测试参数
p.SetParameterValue("Principal", 100000000);        // 租赁成本：1亿元
p.SetParameterValue("InterestRate", 0.03);           // 年利率：3%
p.SetParameterValue("TotalPeriods", 12);             // 总期数：12期
p.SetParameterValue("PaymentInterval", 1);          // 支付间隔：1个月
p.SetParameterValue("PaymentsPerYear", 12);          // 每年支付次数：12次
p.SetParameterValue("LeaseStartDate", "2026-01-01"); // 起租日：2026-01-01
```

---

### 单元测试

#### 测试 1: 配置管理器初始化

**目的**: 验证 `clsRentalConfig` 类的正确初始化

**测试代码**:
```javascript
function test_ConfigInitialization() {
    try {
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        
        // 验证配置管理器已创建
        if (!r.config) {
            throw new Error("配置管理器未创建");
        }
        
        // 验证列定义
        const cols = r.config.columnDefinitions;
        if (cols.PERIOD !== 'A') {
            throw new Error("列定义错误");
        }
        
        // 验证还款方式配置
        const methods = r.config.repaymentMethods;
        if (!methods["等额本息（后付）"]) {
            throw new Error("还款方式配置缺失");
        }
        
        console.log("✅ 测试1通过: 配置管理器初始化正常");
        return true;
    } catch (error) {
        console.error("❌ 测试1失败:", error.message);
        return false;
    }
}
```

**预期结果**: 所有验证通过，无错误抛出

---

#### 测试 2: 公式生成器初始化

**目的**: 验证 `clsFormulaGenerator` 类的正确初始化

**测试代码**:
```javascript
function test_FormulaGeneratorInitialization() {
    try {
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        
        // 验证公式生成器已创建
        if (!r.formulaGenerator) {
            throw new Error("公式生成器未创建");
        }
        
        // 验证公式生成方法存在
        const methods = [
            "generateEqualPaymentFormulas",
            "generateEqualPaymentAdvanceFormulas",
            "generateEqualPrincipalDailyInterestFormulas",
            "generateEqualPrincipalPeriodicInterestFormulas",
            "generatePrincipalRatioPeriodicInterestFormulas",
            "generatePrincipalRatioDailyInterestFormulas"
        ];
        
        for (const method of methods) {
            if (typeof r.formulaGenerator[method] !== 'function') {
                throw new Error(`公式生成方法缺失: ${method}`);
            }
        }
        
        console.log("✅ 测试2通过: 公式生成器初始化正常");
        return true;
    } catch (error) {
        console.error("❌ 测试2失败:", error.message);
        return false;
    }
}
```

**预期结果**: 所有验证通过，无错误抛出

---

#### 测试 3: 样式管理器初始化

**目的**: 验证 `clsStyleManager` 类的正确初始化

**测试代码**:
```javascript
function test_StyleManagerInitialization() {
    try {
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        
        // 验证样式管理器已创建
        if (!r.styleManager) {
            throw new Error("样式管理器未创建");
        }
        
        // 验证样式管理器方法存在
        const methods = [
            "createTableHeaders",
            "applyDataFormat",
            "addBorder",
            "setBackColor"
        ];
        
        for (const method of methods) {
            if (typeof r.styleManager[method] !== 'function') {
                throw new Error(`样式管理方法缺失: ${method}`);
            }
        }
        
        console.log("✅ 测试3通过: 样式管理器初始化正常");
        return true;
    } catch (error) {
        console.error("❌ 测试3失败:", error.message);
        return false;
    }
}
```

**预期结果**: 所有验证通过，无错误抛出

---

### 集成测试

#### 测试 4: 完整的租金测算流程

**目的**: 验证从初始化到生成租金测算表的完整流程

**测试代码**:
```javascript
function test_CompleteRentalCalculationFlow() {
    try {
        // 1. 设置测试参数
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        
        // 2. 初始化租金测算模块
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        
        // 3. 创建表头
        const headerResult = r.创建租金测算表表头(1, 10);
        if (!headerResult) {
            throw new Error("表头创建失败");
        }
        
        // 4. 生成数据
        const dataResult = r.createDataRange();
        if (!dataResult) {
            throw new Error("数据生成失败");
        }
        
        // 5. 验证数据行数
        const actualRows = r.arrData.length;
        const expectedRows = 12;
        if (actualRows !== expectedRows) {
            throw new Error(`数据行数错误: 期望${expectedRows}行, 实际${actualRows}行`);
        }
        
        // 6. 验证数据列数
        const actualCols = r.arrData[0].length;
        const expectedCols = 13;
        if (actualCols !== expectedCols) {
            throw new Error(`数据列数错误: 期望${expectedCols}列, 实际${actualCols}列`);
        }
        
        console.log("✅ 测试4通过: 完整租金测算流程正常");
        return true;
    } catch (error) {
        console.error("❌ 测试4失败:", error.message);
        return false;
    }
}
```

**预期结果**: 所有验证通过，租金测算表生成成功

---

### 功能测试

#### 测试 5: 等额本息（后付）还款方式

**目的**: 验证等额本息（后付）还款方式的正确性

**测试代码**:
```javascript
function test_EqualPaymentMethod() {
    try {
        // 设置参数
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        r.createDataRange();
        
        // 验证第1期
        const period1 = r.arrData[0];
        if (period1[0] !== "1") {
            throw new Error("第1期期次错误");
        }
        
        // 验证租金公式存在
        const rentFormula = period1[2];
        if (!rentFormula || typeof rentFormula !== 'string') {
            throw new Error("第1期租金公式错误");
        }
        if (!rentFormula.includes("PMT")) {
            throw new Error("第1期租金公式未包含PMT函数");
        }
        
        // 验证本金公式存在
        const principalFormula = period1[3];
        if (!principalFormula || typeof principalFormula !== 'string') {
            throw new Error("第1期本金公式错误");
        }
        if (!principalFormula.includes("PPMT")) {
            throw new Error("第1期本金公式未包含PPMT函数");
        }
        
        console.log("✅ 测试5通过: 等额本息（后付）还款方式正常");
        return true;
    } catch (error) {
        console.error("❌ 测试5失败:", error.message);
        return false;
    }
}
```

**预期结果**: 公式正确，包含PMT和PPMT函数

---

#### 测试 6: 等额本金（按天计息）还款方式

**目的**: 验证等额本金（按天计息）还款方式的正确性

**测试代码**:
```javascript
function test_EqualPrincipalDailyInterestMethod() {
    try {
        // 设置参数
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本金（按天计息）");
        
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        r.createDataRange();
        
        // 验证本金公式（等额本金）
        const period1 = r.arrData[0];
        const principalFormula = period1[3];
        if (!principalFormula.includes("/")) {
            throw new Error("等额本金公式错误");
        }
        
        // 验证利息公式（按天计息）
        const interestFormula = period1[4];
        if (!interestFormula.includes("360")) {
            throw new Error("按天计息公式错误");
        }
        
        console.log("✅ 测试6通过: 等额本金（按天计息）还款方式正常");
        return true;
    } catch (error) {
        console.error("❌ 测试6失败:", error.message);
        return false;
    }
}
```

**预期结果**: 公式正确，包含360天计息逻辑

---

#### 测试 7: 本金比例（按期计息）还款方式

**目的**: 验证本金比例（按期计息）还款方式的正确性

**测试代码**:
```javascript
function test_PrincipalRatioMethod() {
    try {
        // 设置参数
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "本金比例（按期计息）");
        
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        r.createDataRange();
        
        // 验证本金比例列存在
        const period1 = r.arrData[0];
        if (!period1[11] || period1[11].length === 0) {
            throw new Error("本金比例列错误");
        }
        
        // 验证本金比例公式
        const ratioFormula = period1[11];
        if (!ratioFormula.includes("round") && !ratioFormula.includes("ROUND")) {
            throw new Error("本金比例公式错误");
        }
        
        // 验证最后一期的调整公式
        const lastPeriod = r.arrData[11];
        const lastRatioFormula = lastPeriod[11];
        if (!lastRatioFormula.includes("SUM")) {
            throw new Error("最后一期本金比例调整公式错误");
        }
        
        console.log("✅ 测试7通过: 本金比例（按期计息）还款方式正常");
        return true;
    } catch (error) {
        console.error("❌ 测试7失败:", error.message);
        return false;
    }
}
```

**预期结果**: 公式正确，最后一期有SUM调整

---

### 行操作测试

#### 测试 8: 插入行功能

**目的**: 验证插入行功能的正确性

**测试代码**:
```javascript
function test_InsertRows() {
    try {
        // 准备测试数据
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        r.createDataRange();
        
        // 获取当前数据
        let arrData = r.arrData;
        const oldTotalPeriods = arrData.length;
        
        // 创建行管理器
        const rm = new clsRentalTableRowManager();
        rm.Initialize("1租金测算表V1");
        
        // 在第5行插入2行
        const result = rm.insertRows(
            arrData,
            4,  // 位置（从0开始）
            2,  // 插入2行
            null,
            "等额本息（后付）",
            {
                autoUpdateTotalPeriods: true,
                syncToWorksheet: false  // 不同步到工作表
            }
        );
        
        // 验证插入成功
        if (!result.success) {
            throw new Error("插入行失败");
        }
        
        // 验证新总期数
        const newTotalPeriods = result.newTotalPeriods;
        if (newTotalPeriods !== oldTotalPeriods + 2) {
            throw new Error(`总期数错误: 期望${oldTotalPeriods + 2}, 实际${newTotalPeriods}`);
        }
        
        // 验证数据行数
        if (result.arrData.length !== newTotalPeriods) {
            throw new Error("数据行数错误");
        }
        
        console.log("✅ 测试8通过: 插入行功能正常");
        return true;
    } catch (error) {
        console.error("❌ 测试8失败:", error.message);
        return false;
    }
}
```

**预期结果**: 插入成功，总期数增加2行

---

#### 测试 9: 删除行功能

**目的**: 验证删除行功能的正确性

**测试代码**:
```javascript
function test_DeleteRows() {
    try {
        // 准备测试数据
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        r.createDataRange();
        
        // 获取当前数据
        let arrData = r.arrData;
        const oldTotalPeriods = arrData.length;
        
        // 创建行管理器
        const rm = new clsRentalTableRowManager();
        rm.Initialize("1租金测算表V1");
        
        // 删除第5-6行（2行）
        const result = rm.deleteRows(
            arrData,
            4,  // 位置（从0开始）
            2,  // 删除2行
            null,
            "等额本息（后付）",
            {
                autoUpdateTotalPeriods: true,
                syncToWorksheet: false  // 不同步到工作表
            }
        );
        
        // 验证删除成功
        if (!result.success) {
            throw new Error("删除行失败");
        }
        
        // 验证新总期数
        const newTotalPeriods = result.newTotalPeriods;
        if (newTotalPeriods !== oldTotalPeriods - 2) {
            throw new Error(`总期数错误: 期望${oldTotalPeriods - 2}, 实际${newTotalPeriods}`);
        }
        
        // 验证数据行数
        if (result.arrData.length !== newTotalPeriods) {
            throw new Error("数据行数错误");
        }
        
        console.log("✅ 测试9通过: 删除行功能正常");
        return true;
    } catch (error) {
        console.error("❌ 测试9失败:", error.message);
        return false;
    }
}
```

**预期结果**: 删除成功，总期数减少2行

---

### 边界测试

#### 测试 10: 最小期数测试

**目的**: 验证最小期数（1期）的处理

**测试代码**:
```javascript
function test_MinimumPeriods() {
    try {
        // 设置最小期数
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 1);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        
        // 生成数据
        const result = r.createDataRange();
        if (!result) {
            throw new Error("最小期数数据生成失败");
        }
        
        // 验证数据行数
        if (r.arrData.length !== 1) {
            throw new Error("最小期数数据行数错误");
        }
        
        console.log("✅ 测试10通过: 最小期数处理正常");
        return true;
    } catch (error) {
        console.error("❌ 测试10失败:", error.message);
        return false;
    }
}
```

**预期结果**: 1期数据处理正常

---

#### 测试 11: 最大期数测试

**目的**: 验证较大期数（100期）的处理

**测试代码**:
```javascript
function test_MaximumPeriods() {
    try {
        // 设置较大期数
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 100);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        
        // 生成数据
        const result = r.createDataRange();
        if (!result) {
            throw new Error("最大期数数据生成失败");
        }
        
        // 验证数据行数
        if (r.arrData.length !== 100) {
            throw new Error("最大期数数据行数错误");
        }
        
        console.log("✅ 测试11通过: 最大期数处理正常");
        return true;
    } catch (error) {
        console.error("❌ 测试11失败:", error.message);
        return false;
    }
}
```

**预期结果**: 100期数据处理正常

---

#### 测试 12: 无效输入测试

**目的**: 验证无效输入的错误处理

**测试代码**:
```javascript
function test_InvalidInput() {
    try {
        const rm = new clsRentalTableRowManager();
        rm.Initialize("1租金测算表V1");
        
        // 测试1: 空数组
        let validation = rm.validateInput(null, 0, 1, 'insert');
        if (validation.valid) {
            throw new Error("空数组验证失败");
        }
        
        // 测试2: 负数位置
        validation = rm.validateInput([[]], -1, 1, 'insert');
        if (validation.valid) {
            throw new Error("负数位置验证失败");
        }
        
        // 测试3: 零行数
        validation = rm.validateInput([[]], 0, 0, 'insert');
        if (validation.valid) {
            throw new Error("零行数验证失败");
        }
        
        // 测试4: 超出范围
        validation = rm.validateInput([[]], 10, 1, 'insert');
        if (validation.valid) {
            throw new Error("超出范围验证失败");
        }
        
        console.log("✅ 测试12通过: 无效输入处理正常");
        return true;
    } catch (error) {
        console.error("❌ 测试12失败:", error.message);
        return false;
    }
}
```

**预期结果**: 所有无效输入都被正确拒绝

---

### 性能测试

#### 测试 13: 大数据量测试

**目的**: 验证大数据量（500期）的处理性能

**测试代码**:
```javascript
function test_LargeDataPerformance() {
    try {
        const startTime = Date.now();
        
        // 设置大数据量
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 500);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        r.createDataRange();
        
        const endTime = Date.now();
        const duration = endTime - startTime;
        
        // 验证处理时间（应小于10秒）
        if (duration > 10000) {
            console.warn(`⚠️  警告: 大数据量处理时间较长: ${duration}ms`);
        } else {
            console.log(`✅ 性能良好: 处理时间 ${duration}ms`);
        }
        
        // 验证数据行数
        if (r.arrData.length !== 500) {
            throw new Error("大数据量数据行数错误");
        }
        
        console.log("✅ 测试13通过: 大数据量处理正常");
        return true;
    } catch (error) {
        console.error("❌ 测试13失败:", error.message);
        return false;
    }
}
```

**预期结果**: 500期数据处理正常，处理时间合理

---

### 综合测试

#### 测试 14: 完整测试套件

**目的**: 运行所有测试，验证系统完整性

**测试代码**:
```javascript
function runAllTests() {
    console.log("========================================");
    console.log("开始运行完整测试套件");
    console.log("========================================\n");
    
    const tests = [
        test_ConfigInitialization,
        test_FormulaGeneratorInitialization,
        test_StyleManagerInitialization,
        test_CompleteRentalCalculationFlow,
        test_EqualPaymentMethod,
        test_EqualPrincipalDailyInterestMethod,
        test_PrincipalRatioMethod,
        test_InsertRows,
        test_DeleteRows,
        test_MinimumPeriods,
        test_MaximumPeriods,
        test_InvalidInput,
        test_LargeDataPerformance
    ];
    
    let passed = 0;
    let failed = 0;
    
    for (let i = 0; i < tests.length; i++) {
        const test = tests[i];
        console.log(`\n--- 运行测试 ${i + 1}/${tests.length}: ${test.name} ---`);
        
        try {
            const result = test();
            if (result) {
                passed++;
            } else {
                failed++;
            }
        } catch (error) {
            console.error(`❌ 测试异常: ${error.message}`);
            failed++;
        }
    }
    
    console.log("\n========================================");
    console.log("测试套件运行完成");
    console.log("========================================");
    console.log(`总测试数: ${tests.length}`);
    console.log(`通过: ${passed}`);
    console.log(`失败: ${failed}`);
    console.log(`通过率: ${((passed / tests.length) * 100).toFixed(2)}%`);
    console.log("========================================\n");
    
    return {
        total: tests.length,
        passed: passed,
        failed: failed,
        passRate: (passed / tests.length) * 100
    };
}
```

**运行方式**:
```javascript
// 运行所有测试
const testResults = runAllTests();

// 检查测试结果
if (testResults.failed === 0) {
    console.log("🎉 所有测试通过！");
} else {
    console.log("⚠️  存在失败的测试，请检查日志");
}
```

---

### 测试清单

#### 必须通过的测试（P0）

| 测试编号 | 测试名称 | 优先级 | 状态 |
|----------|----------|--------|------|
| 测试1 | 配置管理器初始化 | P0 | ⬜ |
| 测试2 | 公式生成器初始化 | P0 | ⬜ |
| 测试3 | 样式管理器初始化 | P0 | ⬜ |
| 测试4 | 完整租金测算流程 | P0 | ⬜ |
| 测试5 | 等额本息（后付） | P0 | ⬜ |
| 测试6 | 等额本金（按天计息） | P0 | ⬜ |
| 测试7 | 本金比例（按期计息） | P0 | ⬜ |
| 测试8 | 插入行功能 | P0 | ⬜ |
| 测试9 | 删除行功能 | P0 | ⬜ |
| 测试10 | 最小期数测试 | P0 | ⬜ |
| 测试11 | 最大期数测试 | P0 | ⬜ |
| 测试12 | 无效输入测试 | P0 | ⬜ |

#### 建议通过的测试（P1）

| 测试编号 | 测试名称 | 优先级 | 状态 |
|----------|----------|--------|------|
| 测试13 | 大数据量测试 | P1 | ⬜ |
| 测试14 | 完整测试套件 | P1 | ⬜ |

---

### 测试报告模板

```javascript
/**
 * 测试报告
 * 
 * 测试日期：YYYY-MM-DD
 * 测试人员：[姓名]
 * 系统版本：V2
 * 
 * 测试结果汇总：
 * - 总测试数：14
 * - 通过：14
 * - 失败：0
 * - 通过率：100%
 * 
 * 详细测试结果：
 * [列出每个测试的详细结果]
 * 
 * 发现的问题：
 * [列出发现的问题和解决方法]
 * 
 * 建议：
 * [列出改进建议]
 */
```

---

### 持续集成建议

#### 自动化测试脚本

```javascript
/**
 * 自动化测试脚本
 * 每次修改代码后运行
 */
function runAutomatedTests() {
    console.log("开始自动化测试...");
    
    const results = runAllTests();
    
    // 生成测试报告
    const report = {
        timestamp: new Date().toISOString(),
        version: "V2",
        results: results
    };
    
    // 保存测试报告
    console.log("测试报告:", JSON.stringify(report, null, 2));
    
    // 如果有失败的测试，抛出错误
    if (results.failed > 0) {
        throw new Error(`${results.failed}个测试失败`);
    }
    
    console.log("✅ 所有自动化测试通过");
    return report;
}
```

---

### 测试最佳实践

1. **测试隔离**
   - 每个测试应该独立运行
   - 不依赖于其他测试的执行顺序
   - 测试后清理数据

2. **测试覆盖率**
   - 单元测试：覆盖所有公共方法
   - 集成测试：覆盖主要业务流程
   - 边界测试：覆盖边界条件

3. **测试可重复性**
   - 使用固定的测试数据
   - 避免随机性
   - 确保测试结果可预测

4. **测试文档化**
   - 为每个测试添加说明
   - 记录预期结果
   - 记录已知问题

5. **持续测试**
   - 每次修改后运行测试
   - 定期运行完整测试套件
   - 维护测试报告

---

## 版本信息

- **版本**: V2
- **基于**: mParameterManager_v2.js
- **作者**: 徐晓冬
- **设计原则**: 配置驱动、单一职责、消除重复、缓存优先、完善注释
- **最后更新**: 2026-01-28
- **测试版本**: 1.0
