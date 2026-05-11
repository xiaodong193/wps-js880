# SuperPivot 完整功能测试套件 - 技术说明

## 📋 文档信息

- **模块名称**: SuperPivotTestSuite
- **版本**: 2.1.0
- **更新日期**: 2026-02-07
- **符合规范**: WPS JSA ES6-ES2019

---

## 🎯 设计目标

本测试套件旨在全面验证 JSA880 框架中 `z超级透视` 功能的正确性和稳定性，同时作为 WPS JSA 规范编写的示例代码。

---

## 📁 文件结构

```
SuperPivot_完整功能测试.js
├── 常量定义区 (UPPER_SNAKE_CASE)
├── 测试运行器类 (clsTestRunner)
├── 测试数据生成器
├── 测试用例组 (9个测试组)
├── 报告输出函数
├── 主运行函数
└── 模块导出区
```

---

## 🔧 核心组件详解

### 1. 测试运行器类 `clsTestRunner`

```javascript
class clsTestRunner {
    constructor() {
        this.m_results = [];           // 测试结果数组
        this.m_startTime = null;       // 测试开始时间
        this._isRunning = false;       // 私有状态标志
    }
}
```

**设计模式**: 单例模式（通过全局实例 `testRunner` 实现）

**主要方法**:
- `runTest(testName, testFunc)` - 执行单个测试并记录结果
- `assertEqual(actual, expected, message)` - 断言相等
- `assertTrue(condition, message)` - 断言为真
- `assertArrayLength(arr, expectedLength, message)` - 断言数组长度
- `printSummary()` - 打印测试汇总

**技术要点**:
- 使用 ES6 Class 语法定义
- 公共属性使用 `m_` 前缀
- 私有属性使用 `_` 前缀
- 错误处理使用 try-catch 机制

---

### 2. 测试数据生成器

三个标准数据集：

| 函数 | 描述 | 列结构 |
|------|------|--------|
| `createTestData()` | 标准销售数据 | 产品, 地区, 年份, 季度, 销售额, 数量 |
| `createSimpleData()` | 简单测试数据 | 产品, 地区, 销售额 |
| `createMultiLevelData()` | 多层结构数据 | 大区, 省份, 城市, 产品, 销售额 |

**设计思路**:
- 纯函数设计，每次调用返回新的数组
- 数据量适中，兼顾测试覆盖和执行速度
- 包含典型的数据类型：字符串、数字、混合类型

---

### 3. 测试用例分组

### 测试组 1: 基础功能测试 (`testBasicFunctions`)

验证核心功能的正确性：
- ✅ 基础透视 - 单行单列
- ✅ 无列字段（仅行字段）
- ✅ 无行字段（仅列字段）
- ✅ JSA880.透视 快捷方式

### 测试组 2: 多字段测试 (`testMultipleFields`)

验证多维度透视：
- ✅ 多行字段 - 产品和地区
- ✅ 多列字段 - 年份和季度
- ✅ 多行多列字段组合
- ✅ 三层行字段 - 大区省份城市

### 测试组 3: 排序功能测试 (`testSorting`)

验证排序符号功能：
- ✅ 行字段升序 (`+`)
- ✅ 行字段降序 (`-`)
- ✅ 多字段混合排序

### 测试组 4: 聚合函数测试 (`testAggregation`)

验证所有聚合类型：
- ✅ `count()` - 计数
- ✅ `sum("f3")` - 求和
- ✅ `average("f3")` - 平均值
- ✅ `max("f3")` - 最大值
- ✅ `min("f3")` - 最小值
- ✅ 多聚合函数组合

### 测试组 5: Options 选项测试 (`testOptions`)

验证配置选项：
- ✅ `cornerTitle` - 角落标题
- ✅ `layoutMode: outline` - 大纲布局
- ✅ `layoutMode: compact` - 紧凑布局
- ✅ `rowFieldIndent` - 行缩进

### 测试组 6: 小计和总计测试 (`testSubtotalsAndGrandTotals`)

验证汇总功能：
- ✅ 行小计
- ✅ 总计行
- ✅ 总计列
- ✅ 行列总计

### 测试组 7: 字段标题自定义测试 (`testFieldTitles`)

验证标题自定义：
- ✅ 行字段自定义标题
- ✅ 数据字段自定义标题
- ✅ 多数据字段自定义标题

### 测试组 8: 边界情况测试 (`testEdgeCases`)

验证异常处理：
- ✅ 空数据（仅表头）
- ✅ 单行数据
- ✅ 重复数据
- ✅ null/undefined 值
- ✅ 大数据量 (1000行)

### 测试组 9: 实战场景测试 (`testRealWorldScenarios`)

验证实际应用场景：
- ✅ 销售报表（产品×地区）
- ✅ 年度季度对比
- ✅ 区域层级分析

---

## 📊 报告输出

### 工作表结构

**工作表名称**: `SuperPivot测试报告`

| 区域 | 内容 |
|------|------|
| A1 | 报告标题 |
| A3:B5 | 基本信息（时间、版本、状态） |
| A7:B11 | 汇总信息（总数、通过、失败、通过率） |
| A13:D13 | 详细结果表头 |
| A14:D... | 详细测试结果列表 |

### 样式规范

- **通过**: 绿色字体 (`0x008000`)
- **失败**: 红色字体 (`0xFF0000`)
- **表头**: 浅蓝色背景 (`0xD9E1F2`) + 粗体
- **标题**: 16号字体 + 粗体

---

## 🚀 使用指南

### 基础用法

```javascript
// 运行所有测试
runAllTests();

// 运行测试 + 透视示例
runTestsWithExamples();

// 快速测试
runQuickTest();

// 性能测试
runPerformanceTest(10000);
```

### 扩展测试

添加新的测试用例：

```javascript
function testMyNewFeature() {
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 X: 我的新功能");
    Console.log("------------------------------------------------------------");
    
    testRunner.runTest("我的测试", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(
            data,
            ["f1+"],
            [],
            ["sum(\"f3\")"],
            1
        );
        
        testRunner.assertTrue(result.length > 0, "结果不应为空");
    });
}

// 在 runAllTests() 中添加调用
testMyNewFeature();
```

---

## 📝 命名规范对照表

| 类型 | 规范 | 示例 |
|------|------|------|
| 类名 | cls + PascalCase | `clsTestRunner` |
| 常量 | UPPER_SNAKE_CASE | `MODULE_VERSION` |
| 公共属性 | m_ + camelCase | `m_results` |
| 私有属性 | _ + camelCase | `_isRunning` |
| 函数/方法 | camelCase | `runAllTests()` |
| 变量 | camelCase | `testData` |

---

## ⚠️ 兼容性检查清单

- [x] **ES2020+ 运算符**: 未使用 `?.` `??` `&&=` `||=`
- [x] **WPS 方法调用**: 所有方法调用带括号 `()`
- [x] **模块系统**: 未使用 `import`/`export`
- [x] **浏览器对象**: 未使用 `window`/`document`
- [x] **Node.js API**: 未使用 `require()`/`fs`
- [x] **变量声明**: 使用 `let`/`const` 而非 `var`
- [x] **字符串引号**: 统一使用双引号 `"`

---

## 🔍 技术难点解析

### 难点 1: WPS 工作表操作

```javascript
// 正确获取工作表
const sheet = Worksheets(REPORT_SHEET_NAME);

// 正确清除内容
sheet.Cells.ClearContents();

// 正确使用 Cells
sheet.Cells(row, col).Value = "数据";
sheet.Cells(row, col).Font.Bold = true;
```

**注意**: WPS JSA 中 `Worksheets()` 是方法，必须带括号。

### 难点 2: 动态行列引用

```javascript
// 获取最后一行
const lastRow = sheet.Cells(sheet.Rows.Count, 1).End(XL_UP).Row;

// 列号转字母
const colLetter = JSA880.列号(columnIndex);
```

### 难点 3: 字符串转义

```javascript
// 在 WPS JSA 中使用 JSON.stringify 安全地序列化数据
Console.log("结果: " + JSON.stringify(result[0]));
```

---

## 📈 版本历史

| 版本 | 日期 | 变更 |
|------|------|------|
| 2.1.0 | 2026-02-07 | 符合 WPS JSA 命名规范，添加完整注释 |
| 2.0.0 | 2026-02-07 | 初始版本，功能完整 |

---

## 📞 技术支持

- **框架文档**: https://vbayyds.com/api/jsa880/
- **原作者**: 郑广学 (微信: EXCEL880B)
- **维护者**: JSA880 框架团队
