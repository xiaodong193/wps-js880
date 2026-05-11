# m调息.js 重构说明文档

> **版本**: V2.20260130  
> **更新日期**: 2026-01-30  
> **作者**: 徐晓冬

---

## 一、重构概述

### 1.1 重构目标

1. **支持所有六种还款方式的调息功能**
2. **消除代码重复**，提高代码复用性
3. **规范命名**，符合WPS JSA编码规范
4. **完善JSDoc注释**，提高代码可读性
5. **加强错误处理**，提高系统稳定性
6. **添加安全机制**（操作确认、备份功能）
7. **实现操作撤销/重做功能**（支持多步历史记录）

### 1.2 主要变更

| 变更项 | 变更前 | 变更后 |
|--------|--------|--------|
| 类名 | `cls调息` | `clsInterestRateAdjustment` |
| 还款方式支持 | 仅支持等额本金法和本金比例法 | 支持全部六种还款方式 |
| 代码结构 | 重复代码多，难以维护 | 模块化设计，消除重复 |
| 命名规范 | 中英文混用 | 统一英文驼峰命名 |
| 注释 | 部分缺失 | 完整JSDoc注释 |
| 版本管理 | 无版本信息 | 完整的版本和日期信息 |

---

## 二、支持的还款方式

### 2.1 六种还款方式完整支持

```javascript
// 1. 等额本息（后付）
"等额本息（后付）"

// 2. 等额本息（先付）
"等额本息（先付）"

// 3. 等额本金（按天计息）
"等额本金（按天计息）"

// 4. 等额本金（按期计息）
"等额本金（按期计息）"

// 5. 本金比例（按期计息）
"本金比例（按期计息）"

// 6. 本金比例（按天计息）
"本金比例（按天计息）"
```

### 2.2 使用示例

```javascript
// 创建实例
const calc = new clsInterestRateAdjustment();
calc.Initialize("1租金测算表V1");

// 启用调息功能
calc.initializeAdjustment({
    isEnabled: true,
    adjustmentType: ADJUSTMENT_TYPES.FIXED
});

// 添加调息节点
calc.addAdjustmentPeriod(6, 0.035);  // 第6期开始利率3.5%
calc.addAdjustmentPeriod(12, 0.04);  // 第12期开始利率4%

// 生成调息公式数组（支持任意还款方式）
const formulaArray = calc.generateAdjustmentFormulaArray("等额本息（后付）");
```

---

## 三、核心类设计

### 3.1 类结构

```
clsInterestRateAdjustment (调息功能类)
├── 继承: clsRentalCalculation (租金测算基类)
├── 属性
│   ├── m_adjustmentPeriods: Array     // 调息节点数组
│   ├── m_adjustmentConfig: Object     // 调息配置
│   ├── VERSION: String                // 版本号
│   └── MODIFY_DATE: String            // 修改日期
└── 方法
    ├── 调息配置管理
    │   ├── initializeAdjustment()      // 初始化调息配置
    │   ├── addAdjustmentPeriod()       // 添加调息节点
    │   ├── batchAddAdjustments()       // 批量添加
    │   ├── clearAdjustments()          // 清除所有节点
    │   └── getApplicableRate()         // 获取适用利率
    ├── 公式生成（六种方式）
    │   ├── generateAdjustmentFormulaArray()       // 通用入口
    │   ├── generateEqualPaymentPostFormula()      // 等额本息后付
    │   ├── generateEqualPaymentAdvanceFormula()   // 等额本息先付
    │   ├── generateEqualPrincipalDailyFormula()   // 等额本金按天
    │   ├── generateEqualPrincipalPeriodicFormula()// 等额本金按期
    │   ├── generatePrincipalRatioPeriodicFormula()// 本金比例按期
    │   └── generatePrincipalRatioDailyFormula()   // 本金比例按天
    ├── 数据生成
    │   ├── createDataRange()           // 重写：生成数据区域
    │   ├── processAdjustmentColumn()   // 处理调息列
    │   └── highlightAdjustmentArea()   // 标黄调息区域
    ├── 工具方法
    │   ├── createFormulaArrayTemplate()// 创建公式数组模板
    │   ├── fillLastRow()               // 填充末期行
    │   ├── copyWorksheet()             // 复制工作表
    │   ├── generateNewSheetName()      // 生成新表名
    │   └── confirmBackup()             // 确认备份
    └── 便捷函数
        ├── generateAdjustmentRentalTable()        // 生成调息表
        └── quickGenerateEqualPaymentAdjustment()  // 快速生成
```

### 3.2 代码复用设计

```javascript
// 1. 统一入口方法
generateAdjustmentFormulaArray(repaymentMethod) {
    // 根据还款方式分发到对应的具体实现
    switch (repaymentMethod) {
        case "等额本息（后付）":
            return this.generateEqualPaymentPostFormula();
        // ... 其他方式
    }
}

// 2. 公共模板方法
createFormulaArrayTemplate() {
    // 创建5x13的二维数组模板
}

// 3. 公共填充方法
fillLastRow(arr, principalCell, leaseStartDateCell, ...) {
    // 统一填充末期行
}
```

---

## 四、命名规范改进

### 4.1 类命名

| 类型 | 旧命名 | 新命名 | 说明 |
|------|--------|--------|------|
| 类名 | `cls调息` | `clsInterestRateAdjustment` | PascalCase，英文命名 |
| 属性 | `adjustmentPeriods` | `m_adjustmentPeriods` | m_前缀表示成员变量 |
| 属性 | `adjustmentConfig` | `m_adjustmentConfig` | m_前缀表示成员变量 |
| 方法 | `InitializeAdjustment` | `initializeAdjustment` | camelCase命名 |
| 方法 | `AddAdjustmentPeriod` | `addAdjustmentPeriod` | camelCase命名 |

### 4.2 常量定义

```javascript
// 调整类型常量
const ADJUSTMENT_TYPES = {
    FIXED: "固定调整",
    FLOATING: "浮动调整",
    CUSTOM: "自定义"
};

// 调整依据常量
const ADJUSTMENT_BASIS = {
    BENCHMARK: "基准利率",
    LPR: "LPR",
    FIXED: "固定值"
};
```

---

## 五、安全机制

### 5.1 操作确认

```javascript
confirmBackup(sheetName) {
    // 显示确认对话框
    const message = `即将为工作表"${sheetName}"生成调息表...`;
    return confirm(message);
}
```

### 5.2 备份机制

```javascript
generateAdjustmentTable(adjustmentArray, sourceSheetName) {
    // 1. 先确认备份
    if (!this.confirmBackup(sourceSheetName)) {
        return false;
    }
    
    // 2. 复制工作表（原表不受影响）
    const newSheet = this.copyWorksheet(sourceSheetName);
    
    // 3. 在新表上操作
    // ...
}
```

---

## 六、操作撤销/重做功能

### 6.1 功能概述

系统实现了完整的**撤销(Undo)/重做(Redo)**功能，支持：
- 最多10步操作历史记录（可配置）
- 撤销后可以继续重做
- 新操作自动清空重做栈
- 操作前自动备份数据

### 6.2 核心方法

#### 撤销/重做控制方法

| 方法 | 说明 |
|------|------|
| `backupCurrentState(type, desc)` | 备份当前状态 |
| `undo()` | 撤销上一步操作 |
| `redo()` | 重做上一步撤销 |
| `canUndo()` | 检查是否可以撤销 |
| `canRedo()` | 检查是否可以重做 |
| `getUndoHistory()` | 获取撤销历史 |
| `clearHistory()` | 清空所有历史 |
| `setMaxHistorySize(size)` | 设置最大历史记录数 |

#### 支持撤销的操作方法

| 操作方法 | 说明 | 自动应用 |
|----------|------|----------|
| `addAdjustmentPeriodWithUndo(period, rate, autoApply)` | 添加调息节点 | 可选 |
| `batchAddAdjustmentsWithUndo(array, autoApply)` | 批量添加调息节点 | 可选 |
| `removeAdjustmentPeriodWithUndo(period, autoApply)` | 删除单个调息节点 | 可选 |
| `updateAdjustmentRateWithUndo(period, rate, autoApply)` | 修改调息利率 | 可选 |
| `clearAdjustmentsWithUndo(autoApply)` | 清除所有调息节点 | 可选 |
| `initializeAdjustmentWithUndo(config, autoApply)` | 初始化调息配置 | 可选 |
| `generateAdjustmentTableWithUndo(array, sheet)` | 生成调息表 | 是 |
| `applyAdjustmentsWithUndo(array)` | 应用调息节点 | 是 |

### 6.3 使用示例

#### 基础示例

```javascript
// 创建实例
const calc = new clsInterestRateAdjustment();
calc.Initialize("1租金测算表V1");

// 应用调息节点（自动备份）
calc.applyAdjustmentsWithUndo([
    { period: 6, newRate: 0.035 },
    { period: 12, newRate: 0.04 }
]);

// 撤销操作
const undoResult = calc.undo();
if (undoResult.success) {
    console.log("已撤销：", undoResult.message);
}

// 重做操作
const redoResult = calc.redo();
if (redoResult.success) {
    console.log("已重做：", redoResult.message);
}

// 检查状态
console.log("可撤销：", calc.canUndo());
console.log("可重做：", calc.canRedo());

// 查看历史
const history = calc.getUndoHistory();
history.forEach(function(item) {
    console.log(`${item.index + 1}. ${item.description}`);
});
```

#### 添加调息节点（带撤销）

```javascript
// 添加单个调息节点，自动应用（重新生成数据）
calc.addAdjustmentPeriodWithUndo(6, 0.035, true);

// 撤销添加操作
calc.undo();  // 恢复添加前的状态

// 重做
calc.redo();  // 重新添加调息节点
```

#### 批量添加调息节点（带撤销）

```javascript
// 批量添加多个调息节点
calc.batchAddAdjustmentsWithUndo([
    { period: 3, newRate: 0.032 },
    { period: 6, newRate: 0.035 },
    { period: 9, newRate: 0.038 }
], true);  // true = 自动应用

// 撤销批量添加
calc.undo();
```

#### 删除调息节点（带撤销）

```javascript
// 删除第6期的调息节点
calc.removeAdjustmentPeriodWithUndo(6, true);

// 撤销删除
calc.undo();  // 恢复被删除的调息节点
```

#### 修改调息利率（带撤销）

```javascript
// 将第6期的利率从3.5%改为4%
calc.updateAdjustmentRateWithUndo(6, 0.04, true);

// 撤销修改
calc.undo();  // 恢复为3.5%
```

#### 清除所有调息节点（带撤销）

```javascript
// 清除所有调息节点
calc.clearAdjustmentsWithUndo(true);

// 撤销清除
calc.undo();  // 恢复所有调息节点
```

### 6.4 撤销功能实现原理

```javascript
// 操作历史记录结构
this.m_operationHistory = [
    {
        type: "generate_table",           // 操作类型
        description: "生成调息表",         // 操作描述
        timestamp: Date,                  // 操作时间
        data: [[...], [...]],             // 备份的数据
        adjustmentPeriods: [...],         // 备份的调息节点
        adjustmentConfig: {...},          // 备份的配置
        range: { startRow: 10, endRow: 21, ... }  // 数据范围
    },
    // ... 更多历史记录
];

// 重做栈结构
this.m_redoStack = [
    // 结构同上，存储撤销前的状态
];
```

### 6.5 便捷函数

```javascript
// 撤销上一步操作
undoAdjustment();

// 重做上一步撤销
redoAdjustment();

// 检查是否可以撤销
canUndoAdjustment();

// 生成带撤销支持的调息表
generateAdjustmentTableWithUndoSupport(adjustments, sheetName);

// 显示撤销历史
showUndoHistory();
```

---

## 七、JSDoc注释规范

### 6.1 类注释

```javascript
/**
 * ============== 租金测算系统 - 调息功能扩展类 ==============
 * 作者：徐晓冬
 * 版本：V2.20260130
 * 描述：继承自 clsRentalCalculation，支持所有还款方式的利率调整功能
 * 
 * 核心改进：
 * - 支持六种还款方式的调息功能
 * - 统一的调息利率管理和应用
 * - 代码复用，消除重复
 * - 完善的错误处理和日志记录
 * - 符合WPS JSA环境规范
 * ====================================================
 */
```

### 6.2 方法注释

```javascript
/**
 * 生成带调息功能的公式数组（通用方法，支持所有还款方式）
 * @param {string} repaymentMethod - 还款方式
 * @returns {Array} 公式数组
 * @throws {Error} 不支持的还款方式时抛出错误
 * @example
 * const arr = calc.generateAdjustmentFormulaArray("等额本息（后付）");
 */
```

### 6.3 参数和返回值

```javascript
/**
 * 添加调息节点
 * @param {number} period - 调息期次（从第几期开始调整）
 * @param {number} newRate - 新利率（年化利率，例如 0.05 表示 5%）
 * @returns {boolean} 是否成功
 */
```

---

## 八、WPS JSA兼容性

### 7.1 遵循的规范

- ✅ 使用 `let`/`const`，不使用 `var`
- ✅ 使用标准JavaScript对象：`Array`, `Object`, `String`, `Date`, `Math`
- ✅ 使用WPS对象模型：`Application`, `Worksheets`, `Range`
- ✅ 方法调用带括号：`Worksheets(1)`, `Range("A1")`
- ✅ 不使用浏览器对象：`window`, `document`
- ✅ 不使用Node.js模块：`require`, `module.exports`
- ✅ 不使用ES2020+语法：`?.`, `??`, `&&=`

### 7.2 兼容性检查清单

- [x] 未使用 `?.` 或 `??` 或 `&&=` 等ES2020+运算符
- [x] 所有方法调用都带有括号 `()`
- [x] 未使用 `import`/`export` 语法
- [x] 未引用 `window`, `document` 等浏览器对象
- [x] 未使用 `require()` 或Node.js内置模块
- [x] 使用了 `let`/`const` 而非 `var`
- [x] 代码可在WPS宏编辑器中正常编译运行

---

## 九、使用示例

### 8.1 基本使用流程

```javascript
// 1. 创建实例
const calc = new clsInterestRateAdjustment();

// 2. 初始化
calc.Initialize("1租金测算表V1");

// 3. 启用调息功能
calc.initializeAdjustment({
    isEnabled: true,
    adjustmentType: ADJUSTMENT_TYPES.FIXED
});

// 4. 添加调息节点
calc.addAdjustmentPeriod(6, 0.035);   // 第6期开始，利率3.5%
calc.addAdjustmentPeriod(12, 0.04);   // 第12期开始，利率4%

// 5. 生成调息表
calc.generateAdjustmentTable([
    { period: 6, newRate: 0.035 },
    { period: 12, newRate: 0.04 }
], "等额本息（后付）", "1租金测算表V1");
```

### 8.2 支持的还款方式使用

```javascript
// 等额本息（后付）
const arr1 = calc.generateAdjustmentFormulaArray("等额本息（后付）");

// 等额本息（先付）
const arr2 = calc.generateAdjustmentFormulaArray("等额本息（先付）");

// 等额本金（按天计息）
const arr3 = calc.generateAdjustmentFormulaArray("等额本金（按天计息）");

// 等额本金（按期计息）
const arr4 = calc.generateAdjustmentFormulaArray("等额本金（按期计息）");

// 本金比例（按期计息）
const arr5 = calc.generateAdjustmentFormulaArray("本金比例（按期计息）");

// 本金比例（按天计息）
const arr6 = calc.generateAdjustmentFormulaArray("本金比例（按天计息）");
```

---

## 十、版本历史

| 版本 | 日期 | 变更内容 |
|------|------|----------|
| V1.0 | 2026-01-28 | 初始版本，支持等额本金法和本金比例法 |
| V2.20260130 | 2026-01-30 | 重构版本，支持全部六种还款方式，规范命名，完善注释 |
| V2.1.20260130 | 2026-01-30 | 新增操作撤销/重做功能，支持多步历史记录（8个操作方法支持撤销/重做） |

---

## 十一、依赖关系

```
m调息.js
├── 继承: mRentalCalculation_v2.js
├── 依赖: mParameterManager_v2.js
├── 依赖: mShared_constants_v2.js
└── 环境: WPS Office JSA
```

---

**维护者**: 徐晓冬  
**最后更新**: 2026-01-30
