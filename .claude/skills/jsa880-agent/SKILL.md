---
name: jsa880-agent
description: |
  JSA880智能助手Agent，用于理解自然语言需求并自动生成WPS JSA代码。当用户询问WPS宏、JSA代码、表格数据处理、自动化办公等需求时使用此技能。
  This skill should be used when:
  - User asks to generate WPS JSA macro code
  - User wants to process Excel/spreadsheet data with automation
  - User asks about Array2D, JSA880 framework functions
  - User describes data processing needs in natural language (Chinese/English)
  - User needs help with WPS Office automation, VBA to JSA migration
---

# JSA880智能助手Agent

## 概述

JSA880智能助手能够理解用户的中文或英文自然语言描述，自动生成符合WPS JSA规范的代码。该Agent集成了JSA880框架的全部能力，包括Array2D类、JSA全局函数、以及其他工具函数。

## 核心功能

### 1. 自然语言理解

Agent能够解析以下类型的用户需求：

| 意图类型 | 关键词 | 示例 |
|---------|--------|------|
| 数组去重 | 去重、唯一值、重复 | "对A1:D100的数据按第1列去重" |
| 分组汇总 | 分组、汇总、统计 | "按部门分组统计人数" |
| 数据筛选 | 筛选、过滤、查找 | "筛选出金额大于1000的记录" |
| 排序操作 | 排序、升序、降序 | "按销售额降序排列" |
| 透视表 | 透视、交叉、分组统计 | "生成按产品和国家分组的销售透视表" |
| 文件操作 | 遍历、复制、移动 | "遍历D盘下的所有Excel文件" |
| 日期处理 | 日期、天数、间隔 | "计算两个日期之间的天数" |
| 单元格操作 | 单元格、区域、格式 | "将选中区域设置为货币格式" |

### 2. 代码生成能力

Agent生成的代码遵循以下原则：

- **使用var声明**：所有变量使用 `var` 声明（本地窗口调试友好）
- **避免禁用语法**：不使用 `?.`、`??`、`??=`、`BigInt` 等ES2020+语法
- **框架优先**：优先使用JSA880框架函数替代手写循环
- **数组驱动**：超过1000行数据必须使用数组操作
- **日期转换**：所有日期操作包含 `cdate()` 或 `DateUtils.fromExcelDate()` 转换

### 3. 安全约束

Agent严格遵守以下约束：

- 不使用 `import`/`export` 模块语法
- 不使用 `window`、`document`、`fetch` 等浏览器API
- 不使用 `require()` 或 Node.js 模块
- 比较操作使用 `===` 而非 `=`
- 方法调用必须带括号：`Worksheets(1)` 而非 `Worksheets 1`

## 意图与函数映射

**Array2D核心方法映射：**

| 用户描述 | 框架函数 | 参数格式 |
|---------|---------|---------|
| "按第1列去重" | `z去重('f1')` | 列选择器 |
| "筛选金额>1000" | `z筛选('f3 > 1000')` | Lambda字符串 |
| "按第1列分组求和" | `z分组('f1')` + 聚合 | 分组键+聚合函数 |
| "按A升序B降序" | `z多列排序('f1+,f2-')` | 列+方向 |
| "左连接两个表" | `z左连接(arr2, 'f1', 'f2')` | 左表键,右表键 |
| "超级透视" | `z超级透视(data, rows, cols, data)` | 四参数 |
| "分页每页10条" | `z按行数分页(10)` | 每页行数 |
| "跳过前5行" | `z跳过(5)` | 跳过行数 |
| "取前10行" | `z取前N个(10)` | 取行数 |

**JSA全局函数映射：**

| 用户描述 | 全局函数 | 用途 |
|---------|---------|------|
| "获取今天日期" | `JSA.z今天()` | 日期字符串 |
| "转换为Excel日期" | `cdate(jsDate)` | Date→OA数值 |
| "Excel日期转JS" | `DateUtils.fromExcelDate()` | OA数值→Date |
| "计算天数差" | `DateUtils.datedif(d1, d2, 'D')` | 日期间隔 |
| "生成序列" | `JSA.z生成数字序列(1, 10, 2)` | 等差数列 |
| "RMB大写" | `JSA.z人民币大写(1234)` | 数字→中文大写 |

## 代码模板

### 模板A：数组处理

```javascript
/**
 * {功能描述}
 * @date {当前日期}
 */
function {函数名}() {
    // 1. 读取数据到二维数组
    var arr = $.maxArray("{数据范围}");
    
    if (!arr || arr.length === 0) {
        console.log("数据为空");
        return;
    }
    
    // 2. 数据处理
    var result = {处理逻辑};
    
    // 3. 输出结果
    result.toRange("{输出范围}", true);
    
    console.log("处理完成，共" + result.length + "行");
}
```

### 模板B：文件操作

```javascript
/**
 * {功能描述}
 * @date {当前日期}
 */
function {函数名}() {
    // 选择文件夹
    var folderPath = IO.showFolderDialog();
    if (!folderPath) {
        MsgBox("未选择文件夹");
        return;
    }
    
    // 遍历文件
    var files = IO.getFiles(folderPath, true, false);
    var result = [];
    
    files.forEach(function(file) {
        {文件处理逻辑}
    });
    
    // 输出结果
    result.toRange("A1", true);
    MsgBox("处理完成，共" + files.length + "个文件");
}
```

### 模板C：数据透视

```javascript
/**
 * {功能描述}
 * @date {当前日期}
 */
function {函数名}() {
    var data = $.maxArray("A1");
    
    // 超级透视
    var pivot = Array2D.z超级透视(
        data,
        ['f1', '{行字段}'],      // 行字段
        ['f2', '{列字段}'],      // 列字段
        ['sum("f3")', '{数据字段}']  // 数据字段
    );
    
    pivot.toRange("H1", true);
    console.log("透视完成");
}
```

## f模式（列选择器）

f模式是JSA880框架的核心语法，用于简化列引用：

| 模式 | 含义 | 示例 |
|------|------|------|
| `f1`, `f2`, `f3` | 第1,2,3列 | `f1 > 100` |
| `f1+`, `f1-` | 升序/降序 | `f1+,f2-` |
| `f1,f2` | 多列选择 | 分组键组合 |
| `f1,f2-f4` | 连续列 | 列范围 |
| `"f1"` | 带引号指定列 | `sum("f3")` |

## Lambda表达式

Lambda字符串可替代回调函数：

```javascript
// Lambda字符串（推荐）
arr.z筛选('f2 > 100');
arr.z映射('x => x.map(v => v * 2)');

// 箭头函数
arr.z筛选(row => row[1] > 100);

// 普通函数
arr.z筛选(function(row, index) {
    return row[1] > 100 && index > 0;
});
```

## 错误诊断

当代码出现错误时，Agent能够诊断以下常见问题：

| 错误类型 | 症状 | 解决方案 |
|---------|------|---------|
| 日期转换缺失 | `d.getFullYear()` 报错 | 使用 `DateUtils.fromExcelDate()` |
| 变量未声明 | 隐式全局变量警告 | 添加 `var` 声明 |
| 比较操作错误 | `if (a = 1)` 导致赋值 | 使用 `if (1 === a)` |
| 数组引用共享 | 修改影响原数组 | 使用 `.z克隆()` 或 `arr.copy()` |
| 可选链禁用 | `?.` 语法报错 | 改用 `obj != null ? obj.prop : undefined` |

## 参考资源

当需要详细API信息时，读取以下文件：
- `references/api_reference.md` - 快速API查询
- `references/array2d_detailed.md` - Array2D类完整方法列表
- `references/agent_logic.js` - 代码生成引擎参考实现