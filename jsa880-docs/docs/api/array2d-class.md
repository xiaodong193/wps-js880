# Array2D 类

> 二维数组处理工具类，支持链式调用

## 概述

Array2D 是 JSA880 框架的核心类，专门用于处理二维数组数据。支持链式调用、Lambda 表达式、多种聚合操作。

## 创建实例

```javascript
// 构造函数方式
var arr = new Array2D(data);

// 静态方法方式
var arr = Array2D(data);

// 从 Range 创建
var arr = Array2D.fromRange('A1:D10');
```

## 基础操作

### val / z值

获取或设置数组值

```javascript
// 获取值
var data = arr.val();

// 设置值（返回新实例）
var newArr = arr.val(newData);
```

---

### z是否为空 / isEmpty

检查数组是否为空

```javascript
var empty = arr.z是否为空();  // true 或 false
```

---

### z数量 / count

获取元素数量

```javascript
var count = arr.z数量();
```

---

### z克隆 / copy

深拷贝数组

```javascript
var cloned = arr.z克隆();
```

---

## 统计计算

### z求和 / sum

对指定列求和

```javascript
var total = arr.z求和();           // 所有元素求和
var total = arr.z求和('f1');        // 指定列求和
var total = arr.z求和(r => r[0] > 5); // 条件求和
```

---

### z平均值 / average

计算平均值

```javascript
var avg = arr.z平均值('f2');
```

---

### z最大值 / max

获取最大值

```javascript
var max = arr.z最大值('f3');
```

---

### z最小值 / min

获取最小值

```javascript
var min = arr.z最小值('f3');
```

---

### z中位数 / median

计算中位数

```javascript
var mid = arr.z中位数('f1');
```

---

## 条件筛选

### z筛选 / filter

按条件筛选行

```javascript
// Lambda 字符串
var filtered = arr.z筛选('f2 > 100');

// 回调函数
var filtered = arr.z筛选(function(row) {
    return row[1] > 100;
});

// 链式筛选
arr.z筛选('f1 === "中国"')
   .z筛选('f3 > 1000');
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `predicate` | String/Function | 条件（Lambda 表达式或函数） |
| `skipHeader` | Boolean | 是否跳过表头（默认自动检测） |

**返回:** `Array2D` - 筛选后的新实例

---

### z跳过 / skip

跳过前 N 行

```javascript
var rest = arr.z跳过(5);  // 跳过前5行
```

---

### z取前N个 / take

取前 N 行

```javascript
var top10 = arr.z取前N个(10);
```

---

### z跳过前面连续满足 / skipWhile

跳过前面连续满足条件的元素

```javascript
var result = arr.z跳过前面连续满足('f1 > 0');
```

---

### z取前面连续满足 / takeWhile

取前面连续满足条件的元素

```javascript
var result = arr.z取前面连续满足('f1 > 0');
```

---

## 排序

### z多列排序 / sortByCols

按多个列排序

```javascript
// 升序/降序
var sorted = arr.z多列排序('f1+,f2-');

// f模式
var sorted = arr.z多列排序('f1+,f2-', 1);  // 1行表头

// 自定义排序
var sorted = arr.z多列排序('f1,f2', 1, ['中国', '美国', '日本']);
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `sortParams` | String | 排序参数（如 'f1+,f2-'） |
| `headerRows` | Number | 表头行数（默认0） |
| `customOrder` | Array | 自定义排序列表 |

**返回:** `Array2D` - 排序后的新实例

---

### z自定义排序 / sortByList

按指定列表排序

```javascript
var sorted = arr.z自定义排序('f1', ['中国', '美国', '日本'], 1);
```

---

### z升序排序 / sortAsc

升序排序

```javascript
var sorted = arr.z升序排序();
```

---

### z降序排序 / sortDesc

降序排序

```javascript
var sorted = arr.z降序排序();
```

---

## 分组透视

### z超级透视 / superPivot

超级透视表 - 核心功能

```javascript
// 基础透视
var pivot = Array2D.z超级透视(
    data,
    ['f1', '产品'],      // 行字段
    ['f2', '国家'],      // 列字段
    ['sum("f3")', '销售额']  // 数据字段
);

// 多层行列字段
var pivot = Array2D.z超级透视(
    data,
    ['f1,f5', '产品类别,地区'],
    ['f3,f4', '年份,季度'],
    ['sum("f7"),count()', '销售额,订单数']
);

// 带选项
var pivot = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 1, '@^@', {
    cornerTitle: '销售报表',
    rowSubtotals: { enabled: true, label: '小计' },
    grandTotals: { row: true, column: true, label: '总计' }
});
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `arr` | Array | 源数据二维数组 |
| `rowFields` | Array/String | 行字段配置 |
| `colFields` | Array/String | 列字段配置 |
| `dataFields` | Array/String | 数据字段配置 |
| `headerRows` | Number | 源数据表头行数（默认1） |
| `outputHeader` | Number/String | 输出表头配置（默认1） |
| `separator` | String | 分隔符（默认'@^@'） |
| `options` | Object | 高级选项 |

**返回:** `Array` 或 `Map` - 透视结果

---

### z分组 / groupBy

按指定键分组

```javascript
var groups = arr.z分组('f1');
var groups = arr.z分组(r => r[0]);
```

---

## 行列操作

### z选择列 / selectCols

选择指定列

```javascript
var cols = arr.z选择列([0, 2, 4]);
var cols = arr.z选择列(['姓名', '年龄']);
var cols = arr.z选择列('f1,f3', ['新标题1', '新标题2']);
```

---

### z选择行 / selectRows

选择指定行

```javascript
var rows = arr.z选择行([0, 2, 5]);
```

---

### z批量删除行 / deleteRows

删除指定行

```javascript
var result = arr.z批量删除行([1, 3, 5]);
```

---

### z批量删除列 / deleteCols

删除指定列

```javascript
var result = arr.z批量删除列([1, 3, 5]);
```

---

### z批量插入行 / insertRows

批量插入行

```javascript
var result = arr.z批量插入行(0, '插入内容', 3);  // 从第0行开始，插入3行
```

---

### z批量插入列 / insertCols

批量插入列

```javascript
var result = arr.z批量插入列(2, '新列', 1);  // 在第2列位置插入1列
```

---

### z插入行号 / insertRowNum

插入行号列

```javascript
var result = arr.z插入行号(1);  // 从1开始的行号
```

---

## 矩阵操作

### z转置 / transpose

转置数组

```javascript
var transposed = arr.z转置();
```

---

### z扁平化 / flat

扁平化数组

```javascript
var flat = arr.z扁平化();
```

---

### z反转 / reverse

反转数组

```javascript
var reversed = arr.z反转();
```

---

### z重复N次 / repeat

重复数组 N 次

```javascript
var repeated = arr.z重复N次(3);
```

---

## 输出

### toRange / z写入单元格

输出到单元格

```javascript
arr.toRange('A1');
arr.z写入单元格('D5', true);
```

---

### z转JSON / toJson

转换为 JSON 字符串

```javascript
var json = arr.z转JSON(true);  // 格式化输出
```

---

### z输出HTML / toHtml

输出 HTML 表格

```javascript
var html = arr.z输出HTML({
    headers: true,
    className: 'table table-striped'
});
```

---

## 遍历

### z遍历执行 / forEach

遍历执行回调

```javascript
arr.z遍历执行(function(row, index) {
    Console.log(index, row);
});
```

---

### z映射 / map

映射转换

```javascript
var mapped = arr.z映射(function(row) {
    return [row[0], row[1] * 2];
});
```

---

### z归约 / reduce

归约操作

```javascript
var sum = arr.z归约(function(acc, row) {
    return acc + row[0];
}, 0);
```

---

## 查找

### z查找单个 / find

查找满足条件的第一个元素

```javascript
var found = arr.z查找单个(function(row) {
    return row[0] === '目标';
});
```

---

### z查找所有下标 / findAllIndex

查找所有满足条件的下标

```javascript
var indexes = arr.z查找所有下标(function(row) {
    return row[1] > 100;
});
```

---

### z值位置 / indexOf

查找值首次出现的位置

```javascript
var pos = arr.z值位置('目标值');
```

---

## 连接操作

### z上下连接 / concat

上下连接数组

```javascript
var combined = arr1.z上下连接(arr2);
```

---

### z左连接 / leftjoin

左外连接

```javascript
var result = arr1.z左连接(arr2, 'f1', 'f1');
```

---

### z左右全连接 / fulljoin

全外连接

```javascript
var result = arr1.z左右全连接(arr2, 'f1', 'f1');
```

---

### z左右连接 / zip

左右拼接（配对合并）

```javascript
var result = arr1.z左右连接(arr2);
```

---

## 集合操作

### z去重 / distinct

数组去重

```javascript
var unique = arr.z去重();
var unique = arr.z去重('f1');  // 按某列去重
```

---

### z排除 / except

差集操作

```javascript
var result = arr1.z排除(arr2);
```

---

### z取交集 / intersect

交集操作

```javascript
var result = arr1.z取交集(arr2);
```

---

### z去重并集 / union

并集操作

```javascript
var result = arr1.z去重并集(arr2);
```

---

## 分页

### z按行数分页 / pageByRows

按行数分页

```javascript
var pages = arr.z按行数分页(10);  // 每页10行
```

---

### z按页数分页 / pageByCount

按页数分页

```javascript
var pages = arr.z按页数分页(5);  // 分成5页
```

---

### z间隔取数 / nth

按间隔取数

```javascript
var result = arr.z间隔取数(2);   // 每2个取1个
var result = arr.z间隔取数(3, 1); // 从第2个开始，每3个取1个
```

---

## 工具

### z版本 / version

获取框架版本

```javascript
var ver = arr.z版本();  // "3.9.3"
```

---

### z错误值 / isError

创建错误值对象

```javascript
var err = arr.z错误值('#VALUE!');
```

---

### z结果 / res

获取最终数组结果

```javascript
var result = arr.z筛选('f1 > 0').z结果();
```

---

### z随机打乱 / shuffle

随机打乱顺序

```javascript
var shuffled = arr.z随机打乱();
```

---

### z随机一项 / random

随机获取 N 项

```javascript
var random = arr.z随机一项(3);
```