# mRentalCalculation_v2.js 撤销功能集成说明

> **版本**: V2.20260130  
> **更新日期**: 2026-01-30  
> **状态**: ✅ 已完成

---

## 一、集成概述

已将统一撤销管理器（`clsUndoManager`）集成到 `mRentalCalculation_v2.js` 中，所有会修改工作表数据的操作现在都支持撤销/重做功能。

---

## 二、核心变更

### 2.1 构造函数更新

```javascript
// 旧构造函数
constructor() {
    // ...
}

// 新构造函数（支持依赖注入）
constructor(parameterManager, undoManager) {
    // ...
    this.p = parameterManager || null;
    this.m_undoManager = undoManager || (typeof g_undoManager !== "undefined" ? g_undoManager : null);
    this.m_undoEnabled = this.m_undoManager !== null;
}
```

### 2.2 新增属性

| 属性名 | 类型 | 说明 |
|--------|------|------|
| `m_undoManager` | clsUndoManager | 撤销管理器实例 |
| `m_undoEnabled` | boolean | 是否启用撤销功能 |

---

## 三、新增方法

### 3.1 撤销管理方法

```javascript
// 设置撤销管理器
setUndoManager(undoManager)

// 获取撤销管理器
getUndoManager()

// 检查是否启用撤销
isUndoEnabled()

// 启用/禁用撤销
setUndoEnabled(enabled)
```

### 3.2 核心操作方法

```javascript
/**
 * 执行可撤销的操作
 * @param {string} type - 操作类型
 * @param {string} description - 操作描述
 * @param {Function} doFn - 执行函数
 * @param {Function} undoFn - 撤销函数
 * @param {Object} metadata - 元数据
 * @returns {boolean} 是否成功
 */
executeUndoable(type, description, doFn, undoFn, metadata)
```

### 3.3 数据备份/恢复方法

```javascript
// 备份当前工作表数据
backupWorksheetData() → Object|null

// 恢复工作表数据
restoreWorksheetData(backup) → boolean
```

### 3.4 撤销/重做控制方法

```javascript
undo()           // 撤销上一步操作
redo()           // 重做上一步撤销
canUndo()        // 检查是否可以撤销
canRedo()        // 检查是否可以重做
getUndoHistory() // 获取撤销历史
clearUndoHistory() // 清空撤销历史
```

---

## 四、支持撤销的操作方法

### 4.1 创建数据区域（带撤销）

```javascript
/**
 * 创建数据区域（带撤销支持）
 * @returns {boolean} 是否成功
 */
createDataRangeWithUndo()

// 使用示例
const calc = new clsRentalCalculation();
calc.Initialize("1租金测算表V1");
calc.createDataRangeWithUndo();  // 可撤销的操作

// 撤销
calc.undo();
```

### 4.2 清除数据（带撤销）

```javascript
/**
 * 清除原有表中数据（带撤销支持）
 * @returns {boolean} 是否成功
 */
清除原有表中数据WithUndo()

// 使用示例
calc.清除原有表中数据WithUndo();
calc.undo();  // 恢复被清除的数据
```

### 4.3 生成每期利率表（带撤销）

```javascript
/**
 * 使用每期适用利率生成测算表（带撤销支持）
 * @returns {boolean} 是否成功
 */
使用每期适用利率生成测算表WithUndo()

// 使用示例
calc.使用每期适用利率生成测算表WithUndo();
calc.undo();  // 恢复到普通利率表
```

### 4.4 修改自定义支付日（带撤销）

```javascript
/**
 * 改变自定义支付日（带撤销支持）
 * @param {number} period - 期次
 * @param {number} value - 值
 * @returns {boolean} 是否成功
 */
改变自定义支付日WithUndo(period, value)

// 使用示例
calc.改变自定义支付日WithUndo(5, 3);  // 第5期改为3个月间隔
calc.undo();  // 恢复原来的间隔
```

---

## 五、便捷函数

### 5.1 创建实例（带撤销）

```javascript
/**
 * 创建租金测算实例（带撤销支持）
 * @param {string} sheetName - 工作表名称
 * @param {clsUndoManager} undoManager - 撤销管理器（可选）
 * @returns {clsRentalCalculation} 租金测算实例
 */
function createRentalCalculationWithUndo(sheetName, undoManager)

// 使用示例
const calc = createRentalCalculationWithUndo("1租金测算表V1");
```

### 5.2 生成表格（带撤销）

```javascript
/**
 * 生成租金测算表（带撤销支持）
 * @param {string} sheetName - 工作表名称
 * @returns {boolean} 是否成功
 */
function generateRentalTableWithUndo(sheetName)

// 使用示例
generateRentalTableWithUndo("1租金测算表V1");
// ...
undoAdjustment();  // 撤销
```

### 5.3 其他便捷函数

```javascript
// 清除租金测算表（带撤销）
clearRentalTableWithUndo(sheetName)

// 使用每期适用利率生成测算表（带撤销）
generatePeriodRateTableWithUndo(sheetName)

// 修改自定义支付日（带撤销）
changeCustomPaymentDayWithUndo(period, value, sheetName)

// 显示撤销历史
showRentalUndoHistory()
```

---

## 六、使用示例

### 6.1 基础使用

```javascript
// 加载模块后使用
function demo() {
    // 1. 创建实例（自动使用全局撤销管理器）
    const calc = createRentalCalculationWithUndo("1租金测算表V1");
    
    // 2. 生成表格（可撤销）
    calc.createDataRangeWithUndo();
    
    // 3. 修改支付日（可撤销）
    calc.改变自定义支付日WithUndo(3, 6);  // 第3期改为6个月
    
    // 4. 查看撤销历史
    showRentalUndoHistory();
    
    // 5. 撤销操作
    calc.undo();  // 撤销修改支付日
    calc.undo();  // 撤销生成表格
    
    // 6. 重做操作
    calc.redo();  // 重做生成表格
}
```

### 6.2 使用独立撤销管理器

```javascript
// 创建独立管理器
const undoManager = createUndoManager({
    maxHistory: 50,
    enableGrouping: true
});

// 创建实例并指定管理器
const calc = new clsRentalCalculation(null, undoManager);
calc.Initialize("1租金测算表V1");

// 执行可撤销操作
calc.createDataRangeWithUndo();

// 使用管理器撤销
undoManager.undo();
```

### 6.3 批量操作（事务）

```javascript
const calc = createRentalCalculationWithUndo("1租金测算表V1");
const manager = calc.getUndoManager();

// 开始事务
manager.beginGroup("批量修改");

// 执行多个操作（都记录到同一组）
calc.改变自定义支付日WithUndo(3, 6);
calc.改变自定义支付日WithUndo(6, 3);
calc.改变自定义支付日WithUndo(9, 6);

// 结束事务
manager.endGroup();

// 一次性撤销所有操作
manager.undo();
```

---

## 七、加载顺序

```javascript
// 1. 必须先加载撤销管理器
mUndoManager_v2.js

// 2. 加载其他依赖模块
mShared_constants_v2.js
mParameterManager_v2.js

// 3. 加载租金测算模块（已集成撤销功能）
mRentalCalculation_v2.js
```

---

## 八、向后兼容

原有方法保持不变，新的撤销功能通过以下方式提供：

1. **新增带`WithUndo`后缀的方法** - 如 `createDataRangeWithUndo()`
2. **原方法仍然可用** - 如 `createDataRange()` 不受影响
3. **可选择性使用** - 不强制要求使用撤销功能

---

## 九、文件统计

| 文件 | 更新前行数 | 更新后行数 | 新增 |
|------|-----------|-----------|------|
| mRentalCalculation_v2.js | ~2100 | 2506 | +406行 |

新增内容：
- 构造函数参数支持
- 15个撤销相关方法
- 6个便捷函数
- 完整文档注释

---

**维护者**: 徐晓冬  
**最后更新**: 2026-01-30
