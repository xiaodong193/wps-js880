# API 概述

JSA880 框架提供两大核心模块：**JSA 全局函数** 和 **Array2D 类**。

## 模块结构

| 模块 | 说明 | 函数/方法数量 |
|------|------|--------------|
| **JSA** | 全局工具函数库 | 33个 |
| **Array2D** | 二维数组处理类 | 140+个 |
| **RngUtils** | Range 区域工具类 | 20+个 |
| **ShtUtils** | 工作表工具类 | 15+个 |
| **DateUtils** | 日期工具类 | 10+个 |

## 快速使用

### JSA 全局函数

```javascript
// 转置数组
var arr = JSA.z转置([[1,2,3],[4,5,6]]);

// 查找
var result = JSA.z查找索引('关键字', arr, 1, false, '未找到');

// 随机打乱
var shuffled = JSA.z随机打乱([1,2,3,4,5]);
```

### Array2D 类

```javascript
// 创建 Array2D 实例
var arr = new Array2D(data);

// 筛选
var filtered = arr.z筛选('f2 > 100');

// 排序
var sorted = arr.z多列排序('f1+,f2-');

// 超级透视
var pivot = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")']);
```

## 函数命名规则

| 规则 | 示例 | 说明 |
|------|------|------|
| 中文名以 `z` 开头 | `z筛选`, `z排序` | 便于识别框架函数 |
| 英文别名 | `filter`, `sort` | 便于英文编程习惯 |
| 静态方法 | `Array2D.z超级透视()` | 通过类名调用 |
| 实例方法 | `arr.z筛选()` | 通过实例调用 |

## 下一步

- [快速开始指南](/guide/getting-started) - 5分钟上手
- [JSA 全局函数](/api/global-functions) - 工具函数参考
- [Array2D 类](/api/array2d-class) - 数组操作参考