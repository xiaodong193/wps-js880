# Array2D 维护指南

> **Array2D - 郑广学JSA880二维数组工具库**
> 
> 版本: 3.5.0 (2026年1月28日维护版)
> 
> 原作者: 郑广学 (EXCEL880)  
> 维护者: Claude Code  
> API文档: https://vbayyds.com/api/jsa880/Array2D.html
> 
> 别名: z二维数组

---

## 📑 目录

1. [功能实现对比总览](#一功能实现对比总览)
2. [基础操作](#二基础操作)
3. [统计计算](#三统计计算)
4. [行列操作](#四行列操作)
5. [筛选排序](#五筛选排序)
6. [连接操作](#六连接操作)
7. [分组透视](#七分组透视)
8. [批量操作](#八批量操作)
9. [分页切片](#九分页切片)
10. [输入输出](#十输入输出)
11. [Lambda表达式](#十一lambda表达式)
12. [链式调用与静态方法](#十二链式调用与静态方法)
13. [增强筛选功能](#十三增强筛选功能)
14. [智能类型识别](#十四智能类型识别)

---

## 架构设计深度解析

### 寄生组合式继承

Array2D 采用**寄生组合式继承**（Parasitic Combination Inheritance）实现，这是 JavaScript 中实现类式继承的最佳模式：

```javascript
// 问题：如何让自定义对象既拥有数组特性，又能添加自定义方法？

// 方案对比：
// 1. 原型继承：Array2D.prototype = new Array() 
//    问题：会执行Array构造函数，可能有问题
//
// 2. 对象扩展：直接修改Array.prototype 
//    问题：污染原生对象，极不推荐
//
// 3. 包装模式：内部持有一个数组 
//    问题：需要代理所有数组方法，复杂且性能差
//
// 4. 寄生组合式继承（本方案）
//    优点：只继承原型，不执行父类构造函数

// 核心实现代码：
Array2D.prototype = Object.create(Array.prototype);
Array2D.prototype.constructor = Array2D;

function Array2D(data) {
    // 将数据附加到实例，使其成为真正的数组
    Array.prototype.push.apply(this, items);
}
```

**设计优点**：
1. Array2D 实例是真正的数组（`[] instanceof Array === true`）
2. 可使用所有原生数组方法（`slice`, `concat`, `forEach` 等）
3. `JSON.stringify` 自动正确处理
4. `for...of` 循环可用

### 链式调用实现

```javascript
// 不可变性原则：每个方法返回新实例，不修改原数据

// 核心方法：_new
Array2D.prototype._new = function(data) {
    var instance = [];
    Array.prototype.push.apply(instance, data);
    Object.setPrototypeOf(instance, Array2D.prototype);
    return instance;
};

// 使用示例：
Array2D.prototype.skip = function(count) {
    return this._new(this._items.slice(count));
};
```

### Lambda 表达式解析

```javascript
// 支持多种形式：
'f1>0'              // 列选择器
'$0*2'              // 索引选择器
'row=>row[0]>0'     // 箭头函数

// 解析流程：
// 1. 检查缓存
// 2. 检测箭头函数语法
// 3. 检测 $ 索引语法
// 4. 检测 f 列选择器语法
// 5. 其他情况作为表达式
```

---

## 一、功能实现对比总览

### 1.1 已实现功能（✅）

| 类别 | 数量 | 说明 |
|------|------|------|
| 基础操作 | 15+ | val, copy, count, fill, reverse, transpose 等 |
| 统计计算 | 8 | sum, average, max, min, median 等 |
| 行列操作 | 20+ | skip, take, getRow, addRow, deleteRow 等 |
| 筛选排序 | 15+ | filter, sort, sortByCols, distinct 等 |
| 连接操作 | 10+ | leftjoin, fulljoin, zip, crossjoin 等 |
| 分组透视 | 8 | groupBy, pivotBy, z超级透视 等 |
| 批量操作 | 10 | deleteCols, insertCols, selectCols 等 |
| 分页切片 | 8 | pageByRows, pageByCount, slice, nth 等 |
| 输入输出 | 5 | toRange, toJson, toString 等 |

### 1.2 功能对照表

| API文档功能 | 实现状态 | 实例方法 | 静态方法 | 备注 |
|------------|---------|---------|---------|------|
| version | ✅ | - | Array2D.version() | 版本号 |
| z克隆/copy | ✅ | z克隆() | Array2D.copy(arr) | 深拷贝 |
| z解析函数表达式/parseLambda | ✅ | - | Array2D.parseLambda() | Lambda解析 |
| z平均值/average | ✅ | z平均值(col) | Array2D.average(arr, col) | 平均值 |
| z上下连接/concat | ✅ | z上下连接(brr) | Array2D.concat(arr, brr) | 上下连接 |
| z复制到指定位置/copyWithin | ✅ | z复制到指定位置() | Array2D.copyWithin() | 复制到位置 |
| z数量/count | ✅ | z数量() | Array2D.count(arr) | 行数统计 |
| z笛卡尔积/crossjoin | ✅ | z笛卡尔积(brr) | Array2D.crossjoin(arr, brr) | 笛卡尔积 |
| z批量删除列/deleteCols | ✅ | z批量删除列(cols) | Array2D.deleteCols(arr, cols) | 删除列 |
| z批量删除行/deleteRows | ✅ | z批量删除行(rows) | Array2D.deleteRows(arr, rows) | 删除行 |
| z排除/except | ✅ | z排除(brr) | Array2D.except(arr, brr) | 排除 |
| z填充/fill | ✅ | z填充(v,r,c) | Array2D.fill() | 填充 |
| z筛选/filter | ✅ | z筛选(fn) | Array2D.filter(arr, fn) | 筛选 |
| z第一个/first | ✅ | z第一个() | Array2D.first(arr) | 第一个元素 |
| z扁平化/flat | ✅ | z扁平化() | Array2D.flat(arr) | 降维 |
| z分组/groupBy | ✅ | z分组(key) | Array2D.groupBy(arr, key) | 分组 |
| z分组汇总/groupInto | ✅ | z分组汇总() | Array2D.groupInto() | 分组汇总 |
| z取交集/intersect | ✅ | z取交集(brr) | Array2D.intersect(arr, brr) | 交集 |
| z最后一个/last | ✅ | z最后一个() | Array2D.last(arr) | 最后一个元素 |
| z映射/map | ✅ | z映射(fn) | Array2D.map(arr, fn) | 映射 |
| z最大值/max | ✅ | z最大值(col) | Array2D.max(arr, col) | 最大值 |
| z中位数/median | ✅ | z中位数(col) | Array2D.median(arr, col) | 中位数 |
| z最小值/min | ✅ | z最小值(col) | Array2D.min(arr, col) | 最小值 |
| z左右全连接/fulljoin | ✅ | z左右全连接() | Array2D.fulljoin() | 全连接 |
| z左连接/leftjoin | ✅ | z左连接() | Array2D.leftjoin() | 左连接 |
| z一对多连接/leftFulljoin | ✅ | z一对多连接() | Array2D.leftFulljoin() | 一对多 |
| z多列排序/sortByCols | ✅ | z多列排序(cols) | Array2D.sortByCols() | 多列排序 |
| z跳过/skip | ✅ | z跳过(n) | Array2D.skip(arr, n) | 跳过前n行 |
| z取前N个/take | ✅ | z取前N个(n) | Array2D.take(arr, n) | 取前n行 |
| z求和/sum | ✅ | z求和(col) | Array2D.sum(arr, col) | 求和 |
| z超级透视/superPivot | ✅ | z超级透视() | Array2D.z超级透视() | 超级透视表 |
| z文本连接/textjoin | ✅ | z文本连接() | Array2D.textjoin() | 文本连接 |
| z去重并集/union | ✅ | z去重并集() | Array2D.union() | 并集去重 |
| z左右连接/zip | ✅ | z左右连接() | Array2D.zip() | 左右连接 |
| z转置/transpose | ✅ | z转置() | Array2D.transpose() | 矩阵转置 |
| z写入单元格/toRange | ✅ | z写入单元格(rng) | Array2D.toRange() | 输出到单元格 |

---

## 二、基础操作

### 2.1 创建 Array2D 对象

```javascript
// 方式1：构造函数
var arr = new Array2D([[1,2,3], [4,5,6]]);

// 方式2：工厂函数（推荐）
var arr = Array2D([[1,2,3], [4,5,6]]);

// 方式3：从 Range 转换
var arr = $.safeArray("A1:D10");

// 方式4：使用 asArray 或 asArray2D
var arr = asArray([[1,2], [3,4]]);
var arr = asArray2D([[1,2], [3,4]]);
```

### 2.2 基础方法

#### val() - 获取/设置值
```javascript
// 获取值
var data = arr.val();

// 设置值
arr.val([[1,2], [3,4]]);
```

#### z克隆() / copy() - 深拷贝
```javascript
// 实例方法
var copy = arr.z克隆();
var copy = arr.copy();

// 静态方法
var copy = Array2D.z克隆(arr);
var copy = Array2D.copy(arr);
```

#### z数量() / count() - 统计数量
```javascript
// 实例方法
var count = arr.z数量();
var count = arr.count();

// 静态方法
var count = Array2D.z数量(arr);
var count = Array2D.count(arr);
```

#### z是否为空() / isEmpty() - 判断是否为空
```javascript
// 实例方法
if (arr.z是否为空()) {
    console.log("数组为空");
}
```

---

## 三、统计计算

### 3.1 统计方法

所有统计方法支持指定列（可选）：

```javascript
// 不传参数 - 统计所有数值
arr.z求和();
arr.z平均值();
arr.z最大值();
arr.z最小值();
arr.z中位数();

// 传列选择器 - 统计指定列
arr.z求和('f1');           // 第1列求和
arr.z平均值('f2');         // 第2列平均
arr.z最大值('f3');         // 第3列最大
arr.z最小值('f4');         // 第4列最小
arr.z中位数('f5');         // 第5列中位数
```

### 3.2 静态方法

```javascript
var arr = [[1,2,3], [4,5,6], [7,8,9]];

Array2D.sum(arr);           // 45
Array2D.sum(arr, 'f1');     // 12 (第1列: 1+4+7)
Array2D.average(arr, 'f2'); // 5 (第2列: (2+5+8)/3)
Array2D.max(arr, 'f3');     // 9
Array2D.min(arr, 'f1');     // 1
Array2D.median(arr, 'f2');  // 5
```

### 3.3 获取首尾元素

```javascript
arr.z第一个();      // 第一个元素
arr.z最后一个();    // 最后一个元素
```

---

## 四、行列操作

### 4.1 跳过和取数

```javascript
// 跳过前N行
arr.z跳过(3);
arr.skip(3);
arr.z跳过前N个(3);      // 别名
arr.z跳过前几个(3);     // 别名

// 取前N行
arr.z取前N个(10);
arr.take(10);
arr.z取前几个(10);      // 别名

// 静态方法
Array2D.skip(arr, 3);
Array2D.take(arr, 10);
```

### 4.2 获取行列

```javascript
// 获取行
arr.z获取行(0);         // 获取第1行
arr.z首行();            // 获取第一行
arr.z末行();            // 获取最后一行

// 获取列
arr.z获取列(0);         // 获取第1列
arr.z首列();            // 获取第一列
arr.z末列();            // 获取最后一列
arr.z提取列(1);         // 提取指定列（返回数组）
```

### 4.3 添加删除行列

```javascript
// 添加行
arr.z添加行([7,8,9]);

// 添加列
arr.z添加列(['a','b','c']);

// 删除行
arr.z删除行(0);         // 删除第1行
arr.z批量删除行([0,2,4]); // 批量删除

// 删除列
arr.z删除列(0);         // 删除第1列
arr.z批量删除列([0,2]); // 批量删除
arr.z批量删除列('f3');  // 使用f模式
```

### 4.4 矩阵操作

```javascript
// 转置
arr.z转置();
arr.transpose();
Array2D.transpose(arr);

// 反转
arr.z反转();
arr.reverse();

// 矩阵信息
arr.z矩阵信息();        // 返回 "行数x列数"
```

---

## 五、筛选排序

### 5.1 筛选

```javascript
// 使用 Lambda 表达式
arr.z筛选('f1>5');
arr.z筛选('f2=="中国"');
arr.z筛选('f6==2023');
arr.z筛选('f1>0 && f2<100');

// 使用函数
arr.z筛选(function(row) {
    return row[0] > 5;
});

// 静态方法
Array2D.filter(arr, 'f1>5');
```

### 5.2 排序

```javascript
// 单列排序
arr.z升序排序();
arr.z降序排序();
arr.z按规则升序('f1');
arr.z按规则降序('f1');

// 多列排序（推荐）
arr.z多列排序('f1+,f2-');       // 第1列升序，第2列降序
arr.z多列排序('f1+,f2-', 1);    // 跳过1行表头

// 自定义排序
arr.z自定义排序('f3', '美国,德国,中国');

// 静态方法
Array2D.sort(arr);
Array2D.sortDesc(arr);
Array2D.sortByCols(arr, 'f1+,f2-', 1);
```

### 5.3 去重

```javascript
// 整行去重
arr.z去重();
arr.distinct();

// 按指定列去重
arr.z去重('f1');
arr.distinct('f1');
```

---

## 六、连接操作

### 6.1 上下连接

```javascript
// 实例方法
arr.z上下连接(brr);
arr.concat(brr);

// 静态方法
Array2D.z上下连接(arr, brr);
Array2D.concat(arr, brr);
```

### 6.2 左右连接

```javascript
// 实例方法
arr.z左右连接(brr);
arr.zip(brr);

// 静态方法
Array2D.zip(arr, brr);
```

### 6.3 SQL风格连接

```javascript
// 左连接
arr.z左连接(brr, 'f1', 'f1');           // 指定左右表连接键
arr.z左连接(brr, 'f1', 'f1', 'f1,f2,f4'); // 指定结果列

// 一对多连接
arr.z一对多连接(brr, 'f1', 'f1');

// 左右全连接（FULL OUTER JOIN）
arr.z左右全连接(brr, 'f1', 'f1');

// 静态方法
Array2D.leftjoin(arr, brr, 'f1', 'f1');
Array2D.leftFulljoin(arr, brr, 'f1', 'f1');
Array2D.fulljoin(arr, brr, 'f1', 'f1');
```

### 6.4 集合操作

```javascript
// 排除
arr.z排除(brr);
Array2D.except(arr, brr);

// 交集
arr.z取交集(brr);
Array2D.intersect(arr, brr);

// 并集（去重）
arr.z去重并集(brr);
Array2D.union(arr, brr);
```

### 6.5 笛卡尔积

```javascript
// 实例方法
arr.z笛卡尔积(brr);

// 静态方法
Array2D.crossjoin(arr, brr);
```

---

## 七、分组透视

### 7.1 分组

```javascript
// 基础分组
arr.z分组('f2');
arr.groupBy('f2');

// 分组汇总
Array2D.groupInto(arr, 'f2', 'g=>g.sum("f3")');

// 分组到Map
Array2D.groupIntoMap(arr, 'f2');
```

### 7.2 超级透视表

```javascript
// 基本用法
var result = Array2D.z超级透视(
    data,                           // 数据源
    ['产品+,国家-'],                 // 行字段（+升序，-降序）
    ['月份+'],                       // 列字段
    ['count(),sum("销量"),average("金额")']  // 数据字段
);

// 带标题的完整用法
var result = Array2D.z超级透视(
    data,
    ['f1,f5,f6', '期数,年,月'],      // [字段配置, 标题]
    ['f2', '国家'],
    [[g=>g.count(),g=>g.sum("f3")], '计数,求和'],
    2,                              // 表头行数
    'map'                           // 返回Map格式
);

// 输出到单元格
result.toRange("A1");

// 或者返回普通数组
var arr = result.res();
```

#### 聚合函数

| 函数 | 说明 | 示例 |
|------|------|------|
| count() | 计数 | count() |
| sum(col) | 求和 | sum("f3") |
| average(col) | 平均值 | average("f4") |
| max(col) | 最大值 | max("f5") |
| min(col) | 最小值 | min("f5") |
| textjoin(col, sep) | 文本连接 | textjoin("f2", ",") |

---

## 八、批量操作

### 8.1 批量删除

```javascript
// 批量删除列
arr.z批量删除列([0, 2, 4]);     // 删除第1,3,5列
arr.z批量删除列('f3');          // 删除第3列
arr.z批量删除列('f2,f4,f6');    // 删除多列
arr.z批量删除列('f3-f7');       // 删除第3到7列

// 批量删除行
arr.z批量删除行([0, 2, 4]);
arr.z批量删除行('f2-f4');       // 删除第2到4行
```

### 8.2 批量插入

```javascript
// 批量插入列
arr.z批量插入列(1, 'x', 2);     // 在第2列位置插入2列，值为'x'

// 批量插入行
arr.z批量插入行([2, 4], 'x', 1); // 在第3和第5行前插入1行，值为'x'
```

### 8.3 选择行列

```javascript
// 选择列
arr.z选择列([0, 2, 4]);         // 选择第1,3,5列
arr.z选择列('f1,f3,f5');        // 使用f模式
arr.z选择列(['产品', '价格']);   // 按表头名选择

// 选择行
arr.z选择行([0, 2, 4, 6]);      // 选择指定行
```

### 8.4 插入行号

```javascript
// 在第1列插入行号
arr.z插入行号(1);               // 从1开始
arr.z插入行号(0);               // 从0开始

// 静态方法
Array2D.insertRowNum(arr, 1, '序号');
```

---

## 九、分页切片

### 9.1 分页

```javascript
// 按页数分页
arr.z按页数分页(5);             // 分成5页

// 按行数分页
arr.z按行数分页(10);            // 每页10行

// 静态方法
Array2D.pageByCount(arr, 5);
Array2D.pageByRows(arr, 10, 2); // 第2页
```

### 9.2 切片

```javascript
// 行切片
arr.z行切片(1, 5);              // 取第2到6行

// 行切片删除行
arr.z行切片删除行(1, 2);        // 从第2行开始删除2行
arr.z行切片删除行(1, 0, ['新行1'], ['新行2']); // 在第2行插入
```

### 9.3 其他切片

```javascript
// 间隔取数
arr.z间隔取数(3);               // 每3行取1行
arr.z间隔取数(3, 1);            // 从第2行开始，每3行取1行

// nth 取数
arr.z间隔取数(2, 0);            // 取偶数行
```

---

## 十、输入输出

### 10.1 输出到单元格

```javascript
// 实例方法
arr.z写入单元格("A1");
arr.toRange("A1");

// 静态方法
Array2D.toRange(arr, "A1");

// 普通数组也支持 toRange
arr.res().toRange("A1");
```

### 10.2 文本输出

```javascript
// 连接成字符串
arr.z连接(',');                 // 用逗号连接
arr.join(',');

// 文本连接（按列）
arr.z文本连接('f1', ',');       // 第1列用逗号连接
arr.textjoin('f2', '+');
```

### 10.3 JSON输出

```javascript
// 格式化输出
arr.z转JSON();
arr.toJson();

// 紧凑格式
arr.z转JSON(false);
arr.toJson(false);
```

### 10.4 转换为矩阵

```javascript
// 转为矩阵格式
arr.z转矩阵(3, 4, 'r');         // 3行4列，行优先
arr.z转矩阵(3, 4, 'c');         // 3行4列，列优先
```

---

## 十一、Lambda表达式

### 11.1 基本语法

```javascript
// 列选择器（f模式）
'f1'           // 第1列
'f2'           // 第2列
'f1+f2'        // 第1列+第2列
'f1*f2'        // 第1列*第2列

// 条件表达式
'f1>5'         // 大于
'f1>=5'        // 大于等于
'f1<5'         // 小于
'f1<=5'        // 小于等于
'f1==5'        // 等于
'f1!="文本"'   // 不等于

// 多条件
'f1>5 && f2<100'
'f1>0 || f2=="中国"'
'f6==2023 && f3=="中国"'
```

### 11.2 在排序中使用

```javascript
// 单列排序
arr.z多列排序('f1+');           // 第1列升序
arr.z多列排序('f1-');           // 第1列降序

// 多列排序
arr.z多列排序('f1+,f2-');       // 第1列升序，第2列降序
arr.z多列排序('f3+,f5-,f2+');   // 多列组合
```

### 11.3 在筛选中使用

```javascript
// 简单筛选
arr.z筛选('f1>5');
arr.z筛选('f2=="中国"');

// 复杂筛选
arr.z筛选('f1>0 && f2<100 && f3=="北京"');
```

---

## 十二、链式调用与静态方法

### 12.1 链式调用

```javascript
// 流畅的链式调用
$.maxArray("A1:H1")
    .skip(1)                    // 跳过表头
    .filter('f2>0')             // 筛选有效数据
    .sortByCols('f1+,f2-')      // 排序
    .take(100)                  // 取前100行
    .toRange("K1");             // 输出

// 分步调试
var arr = $.maxArray("A1:H1");
var filtered = asArray2D(arr).filter('f2>0');
var sorted = filtered.sortByCols('f1+');
var top100 = sorted.take(100);
top100.toRange("K1");
```

### 12.2 静态方法链

```javascript
// 静态方法返回普通数组，可以继续链式调用
var arr = [[1,2], [3,4], [5,6]];
var rs = Array2D.skip(arr, 1);
rs = Array2D.filter(rs, 'f1>0');
rs = Array2D.take(rs, 10);
rs.toRange("A1");               // 普通数组也有 toRange 方法
```

### 12.3 混合使用

```javascript
// 实例方法和静态方法可以混用
var arr = $.maxArray("A1:H100");

// 先用实例方法处理
var processed = asArray2D(arr)
    .filter('f1>0')
    .sortByCols('f2+')
    .res();                     // 获取普通数组

// 再用静态方法
var final = Array2D.take(processed, 10);
final.toRange("K1");
```

---

## 十三、增强筛选功能

### 13.1 对象参数形式

除了传统的 Lambda 表达式，`z筛选` 方法现在支持对象参数形式，提供更灵活的筛选能力：

#### 基础筛选

```javascript
// 简单条件
arr.z筛选({
    column: 'f1',
    operator: '>',
    value: 0
});

// 等价于
arr.z筛选('f1>0');
```

#### 支持的运算符

| 运算符 | 说明 | 别名 |
|--------|------|------|
| `>` | 大于 | `gt` |
| `>=` | 大于等于 | `gte` |
| `<` | 小于 | `lt` |
| `<=` | 小于等于 | `lte` |
| `==` | 等于 | `eq` |
| `!=` | 不等于 | `neq` |
| `in` | 在列表中 | - |
| `nin` | 不在列表中 | `notin` |
| `contains` | 包含子串 | - |
| `startswith` | 以...开头 | - |
| `endswith` | 以...结尾 | - |
| `match` | 正则匹配 | `regex` |
| `between` | 在范围内 | - |
| `isnull` | 为空 | `empty` |
| `notnull` | 不为空 | `notempty` |
| `function` | 自定义函数 | `func` |

#### 复合条件

```javascript
// AND 条件
arr.z筛选({
    column: 'f1',
    operator: '>',
    value: 0,
    and: [
        {column: 'f2', operator: '<', value: 100},
        {column: 'f3', operator: '==', value: '中国'}
    ]
});

// OR 条件
arr.z筛选({
    column: 'f1',
    operator: '>',
    value: 100,
    or: [
        {column: 'f2', operator: '==', value: '中国'},
        {column: 'f2', operator: '==', value: '美国'}
    ]
});

// 复杂组合
arr.z筛选({
    column: 'f1',
    operator: 'between',
    value: [1, 100],
    and: [
        {column: 'f2', operator: 'in', value: ['北京', '上海', '广州']},
        {column: 'f3', operator: 'notnull'}
    ],
    or: [
        {column: 'f4', operator: '>', value: 0}
    ]
});
```

#### 高级筛选示例

```javascript
// 在列表中
arr.z筛选({
    column: 'f1',
    operator: 'in',
    value: ['产品A', '产品B', '产品C']
});

// 包含子串
arr.z筛选({
    column: 'f2',
    operator: 'contains',
    value: '有限公司'
});

// 正则匹配
arr.z筛选({
    column: 'f3',
    operator: 'match',
    value: /^\d{4}-\d{2}-\d{2}$/  // 匹配日期格式
});

// 范围筛选
arr.z筛选({
    column: 'f4',
    operator: 'between',
    value: [100, 1000]
});

// 自定义函数
arr.z筛选({
    column: 'f1',
    operator: 'func',
    value: function(cellValue, row, index) {
        return cellValue > 0 && row[1] < 100;
    }
});
```

### 13.2 链式筛选（QueryBuilder）

使用 `where` 方法（或 `z筛选链`）可以构建流畅的链式筛选：

#### 基础用法

```javascript
// 简单链式筛选
arr.where('f1').gt(0)
   .and('f2').eq('中国')
   .execute();

// 等价于
arr.z筛选('f1>0 && f2=="中国"');
```

#### 所有比较方法

```javascript
arr.where('f1')
   .gt(0)           // 大于
   .gte(0)          // 大于等于
   .lt(100)         // 小于
   .lte(100)        // 小于等于
   .eq('中国')       // 等于
   .neq('美国')      // 不等于
   .in(['北京', '上海'])      // 在列表中
   .nin(['广州', '深圳'])     // 不在列表中
   .contains('公司')          // 包含
   .between(1, 100)           // 在范围内
   .match(/^\d+$/)            // 正则匹配
   .isNull()                  // 为空
   .isNotNull()               // 不为空
   .execute();
```

#### 逻辑组合

```javascript
// AND 组合
arr.where('f1').gt(0)
   .and('f2').lt(100)
   .and('f3').eq('中国')
   .execute();

// OR 组合
arr.where('f1').gt(100)
   .or('f2').eq('VIP')
   .execute();

// 复杂组合
arr.where('f1').between(1, 100)
   .and('f2').in(['北京', '上海'])
   .or('f3').gt(1000)
   .and('f4').isNotNull()
   .execute();
```

#### 切换列

```javascript
// 使用 column 方法切换列
arr.where('f1').gt(0)
   .column('f2').lt(100)      // 切换到 f2 列
   .column('f3').eq('中国')    // 切换到 f3 列
   .execute();

// 使用 and/or 参数切换列
arr.where('f1').gt(0)
   .and('f2').lt(100)         // 切换到 f2 列
   .or('f3').eq('中国')        // 切换到 f3 列
   .execute();
```

#### 中文方法名

```javascript
// 支持中文方法名
arr.z筛选链('f1').大于(0)
   .且('f2').等于('中国')
   .或('f3').小于(100)
   .执行();
```

### 13.3 静态方法

```javascript
// 使用 Array2D.where 创建查询
Array2D.where(arr, 'f1')
    .gt(0)
    .and('f2').eq('中国')
    .execute();
```

---

## 十四、智能类型识别

### 14.1 类型检测

Array2D 现在可以自动识别列的数据类型：

```javascript
// 检测单列类型
var typeInfo = arr.z检测类型('f1');
console.log(typeInfo);
// { type: 'number', format: 'number' }
// { type: 'date', format: 'yyyy-MM-dd' }
// { type: 'string', format: 'text' }
// { type: 'boolean', format: 'boolean' }

// 静态方法
var typeInfo = Array2D.detectType(arr, 0);  // 检测第1列
var typeInfo = Array2D.z检测类型(arr, 'f2'); // 检测第2列
```

#### 支持的类型

| 类型 | 说明 | 示例 |
|------|------|------|
| `number` | 数字 | 123, 45.67, "1,234" |
| `date` | 日期 | 2024-01-15, 2024/01/15, 2024年01月15日 |
| `string` | 字符串 | 文本内容 |
| `boolean` | 布尔值 | true/false, 是/否, YES/NO |

#### 日期格式识别

```javascript
// 自动识别日期格式
'2024-01-15'        -> { type: 'date', format: 'yyyy-MM-dd' }
'2024/01/15'        -> { type: 'date', format: 'yyyy/MM/dd' }
'2024.01.15'        -> { type: 'date', format: 'yyyy.MM.dd' }
'2024年01月15日'    -> { type: 'date', format: 'yyyy年MM月dd日' }
'20240115'          -> { type: 'date', format: 'yyyyMMdd' }
'01-15'             -> { type: 'date', format: 'MM-dd' }
```

### 14.2 智能排序

`z智能排序` 方法会自动识别数据类型并进行适当排序：

```javascript
// 数字列 - 按数值排序
arr.z智能排序('f1');           // 数字升序
arr.z智能排序('f1', '+');      // 数字升序（显式）
arr.z智能排序('f1', '-');      // 数字降序

// 日期列 - 按日期排序
arr.z智能排序('f2');           // 日期升序
arr.z智能排序('f2', '-');      // 日期降序（最新在前）

// 字符串列 - 按拼音排序（中文）或字母排序
arr.z智能排序('f3');           // A-Z 升序
arr.z智能排序('f3', '-');      // Z-A 降序

// 跳过表头
arr.z智能排序('f1', '+', 1);   // 跳过第1行表头

// 静态方法
Array2D.smartSort(arr, 'f1', '+', 1);
Array2D.z智能排序(arr, 'f1', '-');
```

#### 智能排序特性

```javascript
// 自动处理千分位数字
[['1,000'], ['500'], ['2,000']] -> 按数值 500, 1000, 2000 排序

// 自动解析日期格式
[['2024-01-15'], ['2024-01-01'], ['2024-12-31']] -> 按日期排序

// 中文按拼音排序
[['张三'], ['李四'], ['王五']] -> 李、王、张（拼音顺序）

// 空值处理
// 升序：空值在前
// 降序：空值在后
```

### 14.3 智能分组

`z智能分组` 方法根据数据类型自动选择合适的分组策略：

```javascript
// 基本分组
arr.z智能分组('f1');           // 自动按类型分组

// 日期分组 - 按年/月/日/周/季度
arr.z智能分组('f1', 'year');       // 按年分组：2023年、2024年...
arr.z智能分组('f1', 'month');      // 按月分组：2024年1月、2024年2月...
arr.z智能分组('f1', 'day');        // 按日分组：2024-01-15...
arr.z智能分组('f1', 'week');       // 按周分组：2024年第1周...
arr.z智能分组('f1', 'quarter');    // 按季度分组：2024年Q1...

// 数字分组 - 按范围
arr.z智能分组('f1', 'decade');     // 按十位数分组：0-9, 10-19...
arr.z智能分组('f1', 'hundred');    // 按百位数分组：0-99, 100-199...
arr.z智能分组('f1', 'thousand');   // 按千位数分组：0-999, 1000-1999...

// 静态方法
Array2D.smartGroup(arr, 'f1', 'month');
Array2D.z智能分组(arr, 'f1', 'year');
```

#### 智能分组示例

```javascript
// 销售数据按年月分组
var sales = [
    ['2024-01-15', 1000],
    ['2024-01-20', 2000],
    ['2024-02-10', 1500],
    ['2024-03-05', 3000]
];

var groups = Array2D.z智能分组(sales, 'f1', 'month');
// 结果：
// '2024年1月' -> [['2024-01-15', 1000], ['2024-01-20', 2000]]
// '2024年2月' -> [['2024-02-10', 1500]]
// '2024年3月' -> [['2024-03-05', 3000]]

// 年龄数据按十位数分组
var ages = [
    [25], [32], [28], [45], [56], [38], [22]
];

var groups = Array2D.z智能分组(ages, 'f1', 'decade');
// 结果：
// '20-29' -> [[25], [28], [22]]
// '30-39' -> [[32], [38]]
// '40-49' -> [[45]]
// '50-59' -> [[56]]
```

### 14.4 实际应用示例

```javascript
// 示例1：智能处理销售数据
function 分析销售数据() {
    var data = $.maxArray("A1:D1000");  // 日期, 产品, 地区, 金额
    
    // 按年月分组统计
    var groups = data.z智能分组('f1', 'month');
    
    var result = [];
    result.push(['月份', '销售额']);
    
    groups.forEach(function(rows, month) {
        var sum = rows.reduce(function(acc, row) {
            return acc + (parseFloat(row[3]) || 0);
        }, 0);
        result.push([month, sum]);
    });
    
    Array2D.toRange(result, "F1");
}

// 示例2：智能排序混合类型数据
function 智能排序数据() {
    var arr = Array2D([
        ['张三', '2024-01-15', 1500],
        ['李四', '2024-01-01', 2300],
        ['王五', '2024-02-10', 800]
    ]);
    
    // 按日期智能排序
    var sorted = arr.z智能排序('f2');
    sorted.toRange("A5");
    
    // 按金额智能排序（降序）
    var sorted2 = arr.z智能排序('f3', '-');
    sorted2.toRange("E5");
}

// 示例3：复杂筛选 + 智能分组
function 筛选并分组() {
    var data = $.maxArray("A1:E1000");
    
    // 先筛选出大额订单
    var filtered = data.where('f5').gt(10000).execute();
    
    // 再按年月智能分组
    var groups = filtered.z智能分组('f1', 'month');
    
    // 统计每组数量
    groups.forEach(function(rows, month) {
        console.log(month + ': ' + rows.length + ' 笔大额订单');
    });
}
```

---

## 附录：版本更新记录

### v3.4.0 (2024-01-28)

- ✅ 新增 **增强筛选功能**
  - `z筛选` 支持对象参数形式（复杂条件、复合逻辑）
  - 新增 `where` / `z筛选链` 链式筛选构建器
  - 支持 15+ 种运算符（in, between, contains, match 等）
  - 中英文方法名支持
  
- ✅ 新增 **智能类型识别**
  - `z检测类型` 自动识别数字/日期/字符串/布尔值
  - `z智能排序` 根据类型自动选择排序策略
  - `z智能分组` 支持日期（年月日周季度）和数字（十百千位）分组
  
- ✅ 所有新功能同时支持实例方法和静态方法

### v3.3.0 (2024-01-28)

- ✅ 新增 `$.maxArray()` 方法
- ✅ 新增 `z跳过前N个`、`z跳过前几个`、`z取前几个` 别名
- ✅ 完善超级透视表 `z超级透视` 功能
- ✅ 优化链式调用体验
- ✅ 为普通数组添加 `toRange()` 方法

### 原始版本 (Array2D 1.8.3)

- 基础二维数组操作
- 统计计算函数
- 筛选排序功能
- 分组透视表
- 连接操作

---

## 附录：调试报告

详见 `JSA880_调试报告.md`，包含：
- 代码结构分析
- 架构设计深度解析
- 问题修复记录
- 性能优化建议
- 测试建议

---

**文档版本**: 1.1  
**最后更新**: 2026-01-28  
**维护者**: Claude Code
