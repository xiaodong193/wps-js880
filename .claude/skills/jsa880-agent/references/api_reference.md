# JSA880框架 API参考

## Array2D类

### 构造函数

```javascript
// 构造函数方式
var arr = new Array2D(data);

// 静态方法方式
var arr = Array2D(data);

// 从Range创建
var arr = Array2D.fromRange('A1:D10');
```

---

## 基础方法

### val / z值

获取或设置数组值

```javascript
var data = arr.val();           // 获取值
var newArr = arr.val(newData);   // 设置值，返回新实例
```

### z是否为空 / isEmpty

```javascript
var empty = arr.z是否为空();  // true 或 false
```

### z数量 / count

```javascript
var count = arr.z数量();
```

### z克隆 / copy

```javascript
var cloned = arr.z克隆();
```

---

## 条件筛选

### z筛选 / filter

```javascript
// Lambda字符串
arr.z筛选('f2 > 100');

// 箭头函数
arr.z筛选(row => row[1] > 100);

// 回调函数
arr.z筛选(function(row, index) {
    return row[1] > 100;
});

// 链式筛选（AND条件）
arr.z筛选('f1 === "中国"').z筛选('f3 > 1000');
```

### z跳过 / skip

```javascript
var rest = arr.z跳过(5);  // 跳过前5行
```

### z取前N个 / take

```javascript
var top10 = arr.z取前N个(10);
```

---

## 排序

### z多列排序 / sortByCols

```javascript
// 升序/降序
var sorted = arr.z多列排序('f1+,f2-');

// 表头模式
var sorted = arr.z多列排序('f1+,f2-', 1);

// 自定义排序
var sorted = arr.z多列排序('f1', 1, ['中国', '美国', '日本']);
```

### z升序排序 / sortAsc

```javascript
var sorted = arr.z升序排序();
```

### z降序排序 / sortDesc

```javascript
var sorted = arr.z降序排序();
```

---

## 分组透视

### z超级透视 / superPivot

```javascript
// 基础透视
var pivot = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")']);

// 多层行列
var pivot = Array2D.z超级透视(
    data,
    ['f1,f5', '产品类别,地区'],
    ['f3,f4', '年份,季度'],
    ['sum("f7"),count()', '销售额,订单数']
);

// 带高级选项
var pivot = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 1, '@^@', {
    cornerTitle: '销售报表',
    rowSubtotals: { enabled: true, label: '小计' },
    grandTotals: { row: true, column: true, label: '总计' }
});
```

### z分组 / groupBy

```javascript
var groups = arr.z分组('f1');
var groups = arr.z分组(r => r[0]);
```

---

## 行列操作

### z选择列 / selectCols

```javascript
var cols = arr.z选择列([0, 2, 4]);
var cols = arr.z选择列(['姓名', '年龄']);
var cols = arr.z选择列('f1,f3', ['新标题1', '新标题2']);
```

### z批量删除行 / deleteRows

```javascript
var result = arr.z批量删除行([1, 3, 5]);
```

### z批量删除列 / deleteCols

```javascript
var result = arr.z批量删除列([1, 3, 5]);
```

### z批量插入行 / insertRows

```javascript
var result = arr.z批量插入行(0, '插入内容', 3);
```

### z批量插入列 / insertCols

```javascript
var result = arr.z批量插入列(2, '新列', 1);
```

### z插入行号 / insertRowNum

```javascript
var result = arr.z插入行号(1);
```

---

## 连接操作

### z上下连接 / concat

```javascript
var combined = arr1.z上下连接(arr2);
```

### z左连接 / leftjoin

```javascript
var result = arr1.z左连接(arr2, 'f1', 'f1');
```

### z左右全连接 / fulljoin

```javascript
var result = arr1.z左右全连接(arr2, 'f1', 'f1');
```

### z左右连接 / zip

```javascript
var result = arr1.z左右连接(arr2);
```

---

## 集合操作

### z去重 / distinct

```javascript
var unique = arr.z去重();
var unique = arr.z去重('f1');
```

### z排除 / except

```javascript
var result = arr1.z排除(arr2);
```

### z取交集 / intersect

```javascript
var result = arr1.z取交集(arr2);
```

### z去重并集 / union

```javascript
var result = arr1.z去重并集(arr2);
```

---

## 输出

### toRange / z写入单元格

```javascript
arr.toRange('A1');
arr.z写入单元格('D5', true);
```

### z转JSON / toJson

```javascript
var json = arr.z转JSON(true);
```

### z输出HTML / toHtml

```javascript
var html = arr.z输出HTML({
    headers: true,
    className: 'table table-striped'
});
```

---

## 统计计算

### z求和 / sum

```javascript
var total = arr.z求和();
var total = arr.z求和('f1');
var total = arr.z求和(r => r[0] > 5);
```

### z平均值 / average

```javascript
var avg = arr.z平均值('f2');
```

### z最大值 / max

```javascript
var max = arr.z最大值('f3');
```

### z最小值 / min

```javascript
var min = arr.z最小值('f3');
```

### z中位数 / median

```javascript
var mid = arr.z中位数('f1');
```

---

## 分页

### z按行数分页 / pageByRows

```javascript
var pages = arr.z按行数分页(10);
```

### z按页数分页 / pageByCount

```javascript
var pages = arr.z按页数分页(5);
```

### z间隔取数 / nth

```javascript
var result = arr.z间隔取数(2);
var result = arr.z间隔取数(3, 1);
```

---

## 矩阵操作

### z转置 / transpose

```javascript
var transposed = arr.z转置();
```

### z扁平化 / flat

```javascript
var flat = arr.z扁平化();
```

### z反转 / reverse

```javascript
var reversed = arr.z反转();
```

### z重复N次 / repeat

```javascript
var repeated = arr.z重复N次(3);
```

---

## 遍历

### z遍历执行 / forEach

```javascript
arr.z遍历执行(function(row, index) {
    Console.log(index, row);
});
```

### z映射 / map

```javascript
var mapped = arr.z映射(function(row) {
    return [row[0], row[1] * 2];
});
```

### z归约 / reduce

```javascript
var sum = arr.z归约(function(acc, row) {
    return acc + row[0];
}, 0);
```

---

## 查找

### z查找单个 / find

```javascript
var found = arr.z查找单个(function(row) {
    return row[0] === '目标';
});
```

### z查找所有下标 / findAllIndex

```javascript
var indexes = arr.z查找所有下标(function(row) {
    return row[1] > 100;
});
```

### z值位置 / indexOf

```javascript
var pos = arr.z值位置('目标值');
```

---

## 工具

### z版本 / version

```javascript
var ver = arr.z版本();
```

### z错误值 / isError

```javascript
var err = arr.z错误值('#VALUE!');
```

### z结果 / res

```javascript
var result = arr.z筛选('f1 > 0').z结果();
```

### z随机打乱 / shuffle

```javascript
var shuffled = arr.z随机打乱();
```

### z随机一项 / random

```javascript
var random = arr.z随机一项(3);
```

---

# JSA全局函数

## 数组操作

### z转置 / transpose

```javascript
var arr = JSA.z转置([[1,2,3],[4,5,6]]);
// 结果: [[1,4],[2,5],[3,6]]
```

### z选择列 / selectCols

```javascript
var cols = JSA.z选择列(data, [0, 2, 4]);
var cols = JSA.z选择列(data, ['姓名', '年龄']);
```

---

## 数学计算

### z求和 / sum

```javascript
var total = JSA.z求和(1, 2, 3, 4, 5);  // 15
var total = JSA.z求和(...[1, 2, 3]);   // 6
```

### z最大值 / max

```javascript
var max = JSA.z最大值(3, 1, 4, 1, 5);  // 5
```

### z最小值 / min

```javascript
var min = JSA.z最小值(3, 1, 4, 1, 5);  // 1
```

### z平均值 / average

```javascript
var avg = JSA.z平均值(1, 2, 3, 4, 5);  // 3
```

### z生成数字序列 / getNumberArray

```javascript
var nums = JSA.z生成数字序列(1, 10, 2);  // [1, 3, 5, 7, 9]
```

### z取整数 / cint

```javascript
var int = JSA.z取整数(3.14);  // 3
var int = JSA.z取整数(-3.14); // -3
```

### z取小数 / getDecimal

```javascript
var dec = JSA.z取小数(3.14);  // 0.14
```

### z表达式求值 / eval880

```javascript
var result = JSA.z表达式求值('2+3*4');  // 14
```

---

## 日期时间

### z今天 / today

```javascript
var today = JSA.z今天();  // "2026-05-03"
```

### z转日期数值 / cdate

```javascript
var num = cdate(new Date());  // Excel日期序列号
```

### z日期间隔 / datedif

```javascript
var diff = DateUtils.datedif('2024-01-01', '2024-12-31', 'Y');  // 年
var diff = DateUtils.datedif('2024-01-01', '2024-06-15', 'M');  // 月
var diff = DateUtils.datedif('2024-01-01', '2024-01-15', 'D');  // 日
```

---

## 字符串处理

### z转数值 / val

```javascript
var num = JSA.z转数值('123');     // 123
var num = JSA.z转数值('123abc');  // 123
var num = JSA.z转数值('abc');     // 0
```

### z转文本 / cstr

```javascript
var str = JSA.z转文本(123);      // "123"
var str = JSA.z转文本(true);     // "TRUE"
```

### z替换 / replace

```javascript
var str = JSA.z替换('hello world', 'world', 'JSA');  // "hello JSA"
```

### z截取字符 / mid

```javascript
var str = JSA.z截取字符('hello', 2, 3);  // "llo"
```

### z模糊匹配 / like

```javascript
var match = JSA.z模糊匹配('hello', 'h*o');  // true
var match = JSA.z模糊匹配('hello', 'h?llo');  // true
```

---

## 查找函数

### z查找索引 / match

```javascript
var result = JSA.z查找索引('关键字', arr, 1, false, '未找到');
```

### z左侧查找 / vlookup

```javascript
var result = JSA.z左侧查找('关键字', arr, 2, false);
```

### z增强查找 / xlookup

```javascript
var result = JSA.z增强查找('关键字', 查找数组, 结果数组, '未找到');
```

---

## 工具函数

### z人民币大写 / rmbdx

```javascript
var str = JSA.z人民币大写(1234.56);  // "壹仟贰佰叁拾肆元伍角陆分"
```

### z随机整数 / rndInt

```javascript
var rand = JSA.z随机整数(1, 100);  // 1-100
```

### z随机整数数组 / rndIntArray

```javascript
var arr = JSA.z随机整数数组(1, 100, 10);  // 10个随机整数
```

### z随机小数 / rndFloat

```javascript
var rand = JSA.z随机小数(0, 1, 2);  // 0-1之间，保留2位
```

### z随机打乱 / shuffle

```javascript
var shuffled = JSA.z随机打乱([1,2,3,4,5]);
```

### z延时 / delay

```javascript
await JSA.z延时(1000);  // 1000毫秒
```

---

## Lambda / 函数式

### jsaLambda

```javascript
var result = JSA.jsaLambda('x => x * 2', 5);     // 10
var result = JSA.jsaLambda('(x, y) => x + y', 3, 4);  // 7
```

### z解析函数表达式 / parseLambda

```javascript
var fn = JSA.z解析函数表达式('x => x * 2');
fn(5);  // 10
```

---

# 快捷对象

## $ (Range快捷访问)

```javascript
$("A1")          // Range("A1")
$(1, 1)          // Cells(1, 1)
$.maxRange("A1") // 从A1扩展到最大行
$.maxArray("A1") // 一步到位取数组
```

## agg (聚合函数)

```javascript
agg.sum(arr, 'f1')           // 求和
agg.count(arr)              // 计数
agg.average(arr, 'f2')       // 平均值
agg.max(arr, 'f3')           // 最大值
agg.min(arr, 'f3')           // 最小值
agg.textjoin(arr, 'f4', '+') // 文本连接
agg.median(arr, 'f2')        // 中位数
```

---

# 快捷语法

| 写法 | 等价于 | 说明 |
|------|--------|------|
| $("A1") | Range("A1") | 单元格快捷访问 |
| $(1, 1) | Cells(1, 1) | 行列号访问 |
| $.maxRange("A1") | 从A1扩展到最大行Range | 自动按有效数据扩展 |
| $.maxArray("A1") | $.maxRange("A1").safeArray() | 一步到位取数组 |
| rng.safeArray() | 安全获取二维数组 | 单个单元格也返回二维数组 |
| arr.toRange("A1") | 数组输出到单元格 | 自动扩展区域，可清空下方 |
| arr.copy() | 深拷贝数组 | 避免引用共享导致数据污染 |
| asArray(rng) | Range集合转数组 | 可用.filter().unionAll() |
| arr.unionAll() | 数组中的Range联合为一个区域 | 配合asArray实现批量操作 |
| cdate(jsDate) | JS Date → OA数值 | 写入Excel日期用 |
| DateUtils.fromExcelDate(value) | OA数值 → JS Date | 读取Excel日期用 |

---

# IO文件操作

### IO.getFiles

```javascript
// 遍历文件
var files = IO.getFiles(路径, 是否递归, 是否包含隐藏);

// 示例
var files = IO.getFiles("C:\\Data", true, false);
```

### IO.copyFile / IO.moveFile

```javascript
IO.copyFile(原路径, 新路径);
IO.moveFile(文件路径, 目标路径);
```

### IO.mkDir2

```javascript
IO.mkDir2(文件夹路径);  // 已存在不报错
```

### IO.showFolderDialog

```javascript
var folderPath = IO.showFolderDialog();  // 返回选择的路径
```

---

# DateUtils日期处理

```javascript
DateUtils.fromExcelDate(value)       // OA数值 → JS Date
DateUtils.format(date, "yyyy-MM")    // Date → 格式化字符串
DateUtils.datedif(d1, d2, "M")       // 日期间隔（年/月/日）
DateUtils.z只留日期(date)            // 去掉时间部分
DateUtils.z只留时间(date)            // 去掉日期部分
JSA.now                              // 当前日期时间字符串
```

---

# RngUtils单元格工具

```javascript
RngUtils.MergeCells(区域, "r")       // 按行合并相同
RngUtils.MergeCells(区域, "c")       // 按列合并相同
RngUtils.MergeCells(区域, "rm")      // 按上下级关系合并
RngUtils.UnMergeCells(区域)          // 取消合并并填充空行
RngUtils.hitRange(目标单元格, 监测区域)  // 判断命中区域
```

---

# 禁用语法清单

| 禁止项 | ES版本 | 替代方案 |
|--------|--------|----------|
| `?.` 可选链 | ES2020 | `obj != null ? obj.prop : undefined` |
| `??` 空值合并 | ES2020 | `(val != null) ? val : defaultVal` |
| `&&=` `\|\|=` `??=` | ES2021 | 展开为完整 if 条件赋值 |
| `BigInt` | ES2020 | 使用 Number，大数用字符串存储 |
| `String.replaceAll()` | ES2021 | `str.replace(/pattern/g, newStr)` |
| `import` / `export` | 模块系统 | 所有代码在同一文件中，全局对象共享 |
| `window` `document` `navigator` `localStorage` `fetch` | 浏览器API | WPS对象模型 |
| `require()` `fs` `path` 等 | Node.js 模块 | JSA880 IO 函数库 |

---

# 常见错误诊断

| 错误类型 | 错误写法 | 正确写法 |
|---------|----------|----------|
| 日期读取 | `var d = Range("A1").Value2; d.getFullYear()` | `var d = DateUtils.fromExcelDate(Range("A1").Value2)` |
| 日期写入 | `Range("A1").Value2 = new Date()` | `Range("A1").Value2 = cdate(new Date())` |
| 日期比较 | `d1 === d2` (两个Date对象) | `d1.getTime() === d2.getTime()` |
| 单等号判断 | `if (a = 1)` | `if (1 === a)` 值写前面 |
| 区间判断 | `if (1 < x < 5)` | `if (x > 1 && x < 5)` |
| 数组引用共享 | `var brr = arr` | `var brr = arr.copy()` |
| Date对象引用 | `var d2 = d1` | `var d2 = new Date(d1)` |