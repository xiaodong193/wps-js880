# Lambda 表达式

JSA880 框架支持 Lambda 表达式，可以简化回调函数的编写。

## 基本语法

### 箭头函数形式

```javascript
// 单参数
x => x * 2
row => row[0] > 100
x => x.filter('y => y > 0')

// 多参数
(x, y) => x + y
(row, index) => row[0] + index
(a, b, c) => a + b + c
```

### f模式（列选择器）

```javascript
// f1, f2, f3 ... 代表第1, 2, 3 列
'f1 > 100'           // 第1列大于100
'f1 === "中国" && f2 > 5000'  // 复合条件
'f3 * f4'            // 列运算
```

## 使用场景

### z筛选

```javascript
// Lambda 字符串
arr.z筛选('f2 > 100');

// 箭头函数
arr.z筛选(row => row[1] > 100);

// 复杂条件
arr.z筛选('f1 === "手机" && (f3 > 5000 || f4 < 10)');
```

### z映射

```javascript
// 简单变换
arr.z映射('x => x.map(v => v * 2)');

// 提取列
arr.z映射('x => x[0]');

// 条件变换
arr.z映射('x => x.map(v => v > 0 ? v : 0)');
```

### z排序

```javascript
// 按列排序
arr.z排序('x => x[0]');

// 降序
arr.z排序('x => -x[0]');
```

### z分组

```javascript
arr.z分组('x => x[0]');  // 按第1列分组
arr.z分组('x => x[0] + "-" + x[1]');  // 组合键
```

### z查找

```javascript
// 查找单个
arr.z查找单个('x => x[0] === "目标"');

// 查找所有下标
arr.z查找所有下标('x => x[1] > 100');
```

## f模式详解

| 模式 | 含义 | 示例 |
|------|------|------|
| `f1`, `f2`, `f3` | 第1, 2, 3列 | `f1 > 100` |
| `f1+`, `f1-` | 升序/降序 | `f1+,f2-` |
| `f1,f2` | 多列 | `f1,f2` |
| `f1,f2,f3` | 列组合 | 分组键组合 |

## jsaLambda 动态执行

```javascript
// 动态创建函数
var fn = JSA.jsaLambda('x => x * 2', 5);  // 10

// 多参数
var add = JSA.jsaLambda('(x, y) => x + y', 3, 4);  // 7

// 在数组方法中使用
var result = [1, 2, 3].map(JSA.jsaLambda('x => x * 2'));
// 结果: [2, 4, 6]
```

## 与普通函数的对比

```javascript
// 普通函数
arr.z筛选(function(row, index) {
    return row[1] > 100 && index > 0;
});

// Lambda 字符串（推荐）
arr.z筛选('f2 > 100');

// 箭头函数
arr.z筛选((row, index) => row[1] > 100 && index > 0);
```