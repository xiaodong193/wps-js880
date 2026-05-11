# 快速开始

5 分钟上手 JSA880 框架

## 环境要求

- WPS Office 2021 或更高版本
- 启用宏功能

## 安装

### 方式一：复制代码

1. 打开 WPS 宏编辑器（`Alt + F11`）
2. 新建模块，将 `JSA880.js` 代码复制进去
3. 保存即可使用

### 方式二：引用文件

```javascript
// 在 WPS 宏编辑器中引用 JSA880.js
var fso = new ActiveXObject("Scripting.FileSystemObject");
var file = fso.OpenTextFile("C:\\path\\to\\JSA880.js");
var code = file.ReadAll();
file.Close();

// 使用 eval 执行代码（加载框架）
eval(code);
```

## 第一个示例

### 1. 创建 Array2D 实例

```javascript
// 源数据
var data = [
    ['产品', '国家', '销售额', '数量'],
    ['手机', '中国', 10000, 50],
    ['电脑', '美国', 15000, 30],
    ['手机', '美国', 8000, 40],
    ['电脑', '中国', 12000, 25]
];

// 创建 Array2D 实例
var arr = new Array2D(data);
```

### 2. 筛选数据

```javascript
// 筛选销售额大于 9000 的行
var filtered = arr.z筛选('f3 > 9000');

// 输出结果
filtered.z结果().toRange('A10');
```

**结果：**
| 产品 | 国家 | 销售额 | 数量 |
|------|------|--------|------|
| 手机 | 中国 | 10000 | 50 |
| 电脑 | 美国 | 15000 | 30 |
| 电脑 | 中国 | 12000 | 25 |

### 3. 排序

```javascript
// 按销售额降序排序
var sorted = arr.z多列排序('f3-');

// 输出结果
sorted.z结果().toRange('A20');
```

### 4. 超级透视

```javascript
// 按产品和国家进行透视汇总
var pivot = Array2D.z超级透视(
    data,
    ['f1', '产品'],        // 行字段
    ['f2', '国家'],        // 列字段
    ['sum("f3")', '销售额'] // 数据字段
);

// 输出结果
pivot.toRange('A30');
```

**透视结果：**
| 产品 | 中国 | 美国 |
|------|------|------|
| 手机 | 10000 | 8000 |
| 电脑 | 12000 | 15000 |

## 常用操作

### 链式调用

```javascript
// 筛选 → 排序 → 输出
arr.z筛选('f3 > 5000')
   .z多列排序('f1+,f3-')
   .toRange('A1');
```

### 多条件筛选

```javascript
// AND 条件
arr.z筛选('f1 === "手机" && f3 > 5000');

// OR 条件（链式调用自动 OR）
arr.z筛选('f1 === "手机"')
   .z筛选('f2 === "中国"');  // 相当于 AND
```

### Lambda 表达式

```javascript
// 字符串形式的 Lambda
arr.z筛选('x => x[2] > 1000');
arr.z映射('x => x.map(v => v * 2)');

// 参数化 Lambda
var threshold = 5000;
arr.z筛选(function(x) {
    return x[2] > threshold;
});
```

## 下一步

- [Lambda 表达式详解](/guide/lambda) - 掌握 Lambda 语法
- [链式调用](/guide/chaining) - 链式操作详解
- [superPivot 透视表](/guide/super-pivot) - 超级透视表用法
- [API 参考](/api/) - 完整 API 文档