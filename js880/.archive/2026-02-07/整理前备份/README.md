# JSA880 - WPS Office JavaScript API 开发框架

## 📋 项目简介

JSA880 是一个专为 WPS Office JavaScript API (JSA) 设计的快速开发框架，提供了类似 pandas DataFrame 的二维数组处理能力，以及丰富的办公自动化工具函数。

**核心特性**:
- 🚀 高性能二维数组处理 (Array2D 类)
- 📊 Excel 风格超级透视表 (superPivot)
- 📁 文件 IO 操作
- 📅 日期时间工具
- 🔧 Range/工作表快捷操作

## 🎯 适用环境

- **WPS Office**: Windows 版 WPS Office 2019+
- **JavaScript 引擎**: WPS JSA (基于 ChakraCore/V8)
- **语法规范**: ES5-ES2019 (严格避免 ES2020+ 语法)

## ⚠️ 兼容性说明

### 支持的语法
```javascript
// ✅ 推荐使用的语法
var x = 1;                      // 变量声明
function foo() { }              // 函数声明
var fn = function() { };        // 函数表达式
var arr = [1, 2, 3];            // 数组字面量
var obj = { a: 1 };             // 对象字面量
for (var i = 0; i < n; i++) { } // for循环
if (condition) { }              // 条件语句

// ✅ ES6+ 支持的语法（视WPS版本而定）
let y = 2;                      // 块级作用域变量
const Z = 3;                    // 常量
var fn = (x) => x * 2;          // 箭头函数
var [a, b] = [1, 2];            // 解构赋值
var arr = [...oldArr];          // 展开运算符
var str = `Hello ${name}`;      // 模板字符串
```

### 禁止使用的语法
```javascript
// ❌ ES2020+ 运算符（坚决不能使用）
obj?.property;                  // 可选链
value ?? default;               // 空值合并
obj &&= value;                  // 逻辑赋值

// ❌ 浏览器/Node.js 对象
window;                         // 浏览器全局对象
document;                       // DOM对象
localStorage;                   // 本地存储
fetch();                        // 网络请求
require('fs');                  // Node.js模块
module.exports;                 // CommonJS导出
```

## 📦 安装使用

### 方法一：直接导入 WPS

1. 打开 WPS 表格
2. 按 `Alt + F11` 打开宏编辑器
3. 选择「工具」→「导入文件」
4. 选择 `JSA880.js` 文件

### 方法二：代码片段

```javascript
// 在WPS宏编辑器中直接复制以下测试代码
function 测试透视表() {
    // 准备测试数据
    var data = [
        ['产品', '年份', '地区', '销售额'],
        ['A', '2023', '华东', 100],
        ['A', '2023', '华南', 200],
        ['A', '2024', '华东', 150],
        ['B', '2023', '华南', 300],
        ['B', '2024', '华东', 250]
    ];
    
    // 创建透视表
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品'],           // 行字段
        ['f2', '年份'],           // 列字段
        ['sum("f4")', '销售额']   // 数据字段
    );
    
    // 输出到工作表
    result.toRange("A1");
    
    Console.log("透视表已生成！");
}

测试透视表();
```

## 🎓 核心功能示例

### 1. Array2D - 二维数组处理

```javascript
// 创建 Array2D 对象
var arr = new Array2D([['A', 'B'], [1, 2], [3, 4]]);

// 选择列
var col = arr.z选择列(['f1', 'f2']);  // 或 'f1,f2'

// 跳过表头
var data = arr.z跳过(1);

// 筛选
var filtered = arr.z筛选(function(row) {
    return row[0] > 1;
});

// 排序
var sorted = arr.z排序(1, true);  // 按第2列升序

// 去重
var unique = arr.z去重('f1');

// 统计
var sum = arr.z求和('f2');
var avg = arr.z平均值('f2');

// 转置
var transposed = arr.z转置();

// 输出到单元格
arr.toRange("A1");
```

### 2. superPivot - 超级透视表

```javascript
// 基础用法
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],           // 行字段
    ['f2', '年份'],           // 列字段
    ['sum("f4")', '销售额']   // 数据字段
);

// 多层行/列字段
var result = Array2D.z超级透视(
    data,
    ['f1,f2', '大区,省份'],   // 2个行字段
    ['f3,f4', '年份,季度'],   // 2个列字段
    ['count(),sum("f5")', '订单数,销售额']
);

// 带排序符号 (+升序, -降序, #原始顺序)
var result = Array2D.z超级透视(
    data,
    ['f1+,f2-', '大区,省份'],  // 大区升序, 省份降序
    ['f3+', '年份'],
    ['sum("f5")', '销售额']
);

// v3.9.0 高级选项
var result = Array2D.z超级透视(
    data,
    ['f1,f2', '大区,省份'],
    ['f3', '年份'],
    ['sum("f4")', '销售额'],
    1,      // headerRows
    1,      // outputHeader
    '@^@',  // separator
    {
        cornerTitle: '销售分析表',
        layoutMode: 'outline',      // compact/outline/tabular
        rowFieldIndent: true,       // 启用层级缩进
        rowFieldIndentSize: 4,      // 缩进4空格
        rowSubtotals: {
            enabled: true,
            position: 'bottom',
            label: '小计'
        },
        colSubtotals: {
            enabled: true,
            position: 'right',
            label: '小计'
        },
        grandTotals: {
            row: true,
            column: true,
            label: '总计'
        },
        displayAs: {
            mode: 'percentOfGrandTotal',  // percentOfRowTotal/percentOfColTotal/value
            decimals: 2
        }
    }
);

// 输出到工作表
result.toRange("A1", true);  // true=自动合并单元格

// 获取元数据
var meta = result.getMeta();
Console.log(JSON.stringify(meta));
```

### 3. IO - 文件操作

```javascript
// 判断文件/文件夹
var isFile = IO.z是否文件("C:\\test.txt");
var isDir = IO.z是否文件夹("C:\\test");

// 读写文件
IO.z写文件("C:\\test.txt", "Hello World");
var content = IO.z读文件("C:\\test.txt");

// 遍历文件夹
var files = IO.z遍历文件夹("C:\\test", "*.txt");
```

### 4. RngUtils - Range 工具

```javascript
// Range 转数组
var arr = RngUtils.z转数组("A1:D10");

// 数组写入 Range
RngUtils.z从数组("A1", [[1, 2], [3, 4]]);

// 偏移
var newRng = RngUtils.z偏移("A1", 2, 3);  // 向下2行,向右3列

// 调整大小
var resized = RngUtils.z调整大小("A1:B10", 5, 5);
```

### 5. ShtUtils - 工作表工具

```javascript
// 创建/删除
ShtUtils.z创建("新工作表");
ShtUtils.z删除("旧工作表");

// 复制
ShtUtils.z复制("源表", "目标表");

// 隐藏/显示
ShtUtils.z隐藏("Sheet1");
ShtUtils.z显示("Sheet1");

// 清空
ShtUtils.z清空("Sheet1");

// 保护/解除保护
ShtUtils.z保护("Sheet1", "password");
ShtUtils.z解除保护("Sheet1", "password");
```

### 6. DateUtils - 日期工具

```javascript
// 格式化
var str = DateUtils.z格式化(new Date(), "yyyy-MM-dd");

// 日期计算
var nextDay = DateUtils.z添加天数(new Date(), 1);
var nextMonth = DateUtils.z添加月(new Date(), 1);

// 月初月末
var firstDay = DateUtils.z月初(new Date());
var lastDay = DateUtils.z月末(new Date());

// 季度/周
var quarter = DateUtils.z季度(new Date());
var week = DateUtils.z年周(new Date());

// 工作日
var workDays = DateUtils.z工作日计算(new Date(2024, 0, 1), new Date(2024, 0, 31));
```

### 7. $ - 快捷工具

```javascript
// Range 转 Array2D
var arr = $.maxArray("A1:D10");

// 遍历
$.each("A1:A10", function(value, index) {
    Console.log(index + ": " + value);
});

// 映射
var doubled = $.map("A1:A10", function(v) { return v * 2; });

// 筛选
var positives = $.filter("A1:A10", function(v) { return v > 0; });
```

## 📚 API 文档

详细 API 文档请访问：[https://vbayyds.com/api/jsa880/](https://vbayyds.com/api/jsa880/)

## 🔧 开发规范

### 命名规范

| 类型 | 命名规则 | 示例 |
|------|----------|------|
| 类名 | PascalCase | `Array2D`, `RngUtils` |
| 方法名 | camelCase | `z超级透视`, `z选择列` |
| 变量名 | camelCase | `rowCount`, `colIndex` |
| 常量名 | UPPER_SNAKE_CASE | `DEFAULT_SEPARATOR` |
| 私有属性 | `_` + camelCase | `_items`, `_header` |

### 方法调用规范

```javascript
// ✅ 正确 - 带参数的属性必须加括号
Worksheets(1).Range("A1").Value = 100;
Range("B2").Select();

// ❌ 错误 - 缺少括号
Worksheets 1;  // 语法错误
Range "A1";    // 语法错误
```

### 注释规范

```javascript
/**
 * 函数说明
 * @param {Type} paramName - 参数说明
 * @returns {Type} 返回值说明
 * @example
 * // 使用示例
 * var result = foo(1, 2);
 */
function foo(paramName) {
    return paramName * 2;
}
```

## 🐛 调试技巧

```javascript
// 使用 Console.log 输出调试信息
Console.log("调试信息");
Console.log("变量值: " + variable);
Console.log(JSON.stringify(object));

// 使用 try-catch 捕获错误
try {
    // 可能出错的代码
} catch (e) {
    Console.log("错误: " + e.message);
    Console.log("堆栈: " + e.stack);
}

// 性能计时
var start = new Date().getTime();
// ... 执行代码
var elapsed = new Date().getTime() - start;
Console.log("耗时: " + elapsed + "ms");
```

## 📝 版本历史

### v3.9.0 (2026-02-06)
- ✨ 新增 superPivot v3.9.0 功能
  - 行/列小计与总计
  - 百分比显示（占总计%/占行%/占列%）
  - 多种布局模式（compact/outline/tabular）
  - 层级缩进显示
  - 角标题支持
  - getMeta() 元数据方法

### v3.8.2 (2026-02-01)
- 🐛 修复多层列字段表头生成
- ⚡ 优化表头生成性能
- ⚡ 优化 toRange 输出性能

### v3.8.0 (2026-01-15)
- ✨ 支持多行多列透视表
- ✨ 动态计算表头行数
- 🐛 修复 _header 属性传递链

## 📄 许可证

本项目基于 MIT 许可证开源。

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request。

## 📮 联系方式

- 原作者: 郑广学 (EXCEL880)
- 维护者: 徐晓冬
- 官方网站: https://vbayyds.com

---

**免责声明**: 使用本框架产生的任何数据丢失或文件损坏，开发者不承担责任。建议在使用前备份重要数据。
