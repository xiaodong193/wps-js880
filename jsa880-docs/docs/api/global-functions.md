# JSA 全局函数

> JSA 全局工具函数库，增强 JSA 能力，全局可用

## 函数列表

### 数组操作

#### z转置 / transpose

将数组转置（行列互换）

```javascript
var arr = JSA.z转置([[1,2,3],[4,5,6]]);
// 结果: [[1,4],[2,5],[3,6]]
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `arr` | Array | 要转置的二维数组 |

**返回:** `Array` - 转置后的数组

---

#### z选择列 / selectCols

按列索引或表头名称选择列

```javascript
var cols = JSA.z选择列(data, [0, 2, 4]);
var cols = JSA.z选择列(data, ['姓名', '年龄']);
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `arr` | Array | 二维数组 |
| `colIndexes` | Array | 列索引或表头名称数组 |
| `newHeaders` | Array | 新表头（可选） |

**返回:** `Array` - 选择的列组成的新数组

---

### 数学计算

#### z求和 / sum

对多个数值求和

```javascript
var total = JSA.z求和(1, 2, 3, 4, 5);  // 15
var total = JSA.z求和(...[1, 2, 3]);   // 6
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `...args` | Number | 要求和的数值 |

**返回:** `Number` - 求和结果

---

#### z最大值 / max

获取最大值

```javascript
var max = JSA.z最大值(3, 1, 4, 1, 5);  // 5
```

---

#### z最小值 / min

获取最小值

```javascript
var min = JSA.z最小值(3, 1, 4, 1, 5);  // 1
```

---

#### z平均值 / average

计算平均值

```javascript
var avg = JSA.z平均值(1, 2, 3, 4, 5);  // 3
```

---

#### z生成数字序列 / getNumberArray

生成数字序列数组

```javascript
var nums = JSA.z生成数字序列(1, 10, 2);  // [1, 3, 5, 7, 9]
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `start` | Number | 起始值 |
| `end` | Number | 结束值 |
| `step` | Number | 步长（默认1） |

**返回:** `Array` - 数字序列数组

---

#### z取整数 / cint

取数值的整数部分

```javascript
var int = JSA.z取整数(3.14);  // 3
var int = JSA.z取整数(-3.14); // -3
```

---

#### z取小数 / getDecimal

取数值的小数部分

```javascript
var dec = JSA.z取小数(3.14);  // 0.14
```

---

#### z表达式求值 / eval880

计算数学表达式的值

```javascript
var result = JSA.z表达式求值('2+3*4');  // 14
```

---

### 日期时间

#### z今天 / today

获取当前日期字符串

```javascript
var today = JSA.z今天();  // "2026-05-03"
```

**返回:** `String` - 格式为 YYYY-MM-DD 的日期字符串

---

#### z转日期数值 / cdate

将日期转换为 Excel 日期数值

```javascript
var num = JSA.z转日期数值(new Date());  // Excel日期序列号
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `d` | Date | 日期对象 |

**返回:** `Number` - Excel 日期数值

---

#### z日期间隔 / datedif

计算两个日期之间的间隔

```javascript
var diff = JSA.z日期间隔('2024-01-01', '2024-12-31', 'Y');  // 年
var diff = JSA.z日期间隔('2024-01-01', '2024-06-15', 'M');  // 月
var diff = JSA.z日期间隔('2024-01-01', '2024-01-15', 'D');  // 日
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `d1` | String | 开始日期 |
| `d2` | String | 结束日期 |
| `format` | String | 返回格式: 'Y'年, 'M'月, 'D'日, 'YM'年月, 'YD'年月日 |

**返回:** `String` 或 `Number` - 间隔值

---

### 字符串处理

#### z转数值 / val

将字符串转换为数值

```javascript
var num = JSA.z转数值('123');     // 123
var num = JSA.z转数值('123abc');  // 123
var num = JSA.z转数值('abc');     // 0
```

---

#### z转文本 / cstr

将值转换为文本

```javascript
var str = JSA.z转文本(123);      // "123"
var str = JSA.z转文本(true);     // "TRUE"
```

---

#### z替换 / replace

替换字符串中的指定内容

```javascript
var str = JSA.z替换('hello world', 'world', 'JSA');  // "hello JSA"
```

---

#### z截取字符 / mid

截取字符串指定位置的字符

```javascript
var str = JSA.z截取字符('hello', 2, 3);  // "llo" (从第2个字符开始，取3个)
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `str` | String | 原字符串 |
| `start` | Number | 起始位置（从1开始） |
| `len` | Number | 长度 |

**返回:** `String` - 截取的字符串

---

#### z左取字符 / left

从左边截取字符

```javascript
var str = JSA.z左取字符('hello', 2);  // "he"
```

---

#### z右取字符 / right

从右边截取字符

```javascript
var str = JSA.z右取字符('hello', 2);  // "lo"
```

---

#### z模糊匹配 / like

使用通配符进行模糊匹配

```javascript
var match = JSA.z模糊匹配('hello', 'h*o');  // true
var match = JSA.z模糊匹配('hello', 'h?llo');  // true
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `str` | String | 要匹配的字符串 |
| `pattern` | String | 模式（* 任意字符，? 单个字符） |

**返回:** `Boolean` - 是否匹配

---

### 查找函数

#### z查找索引 / match

增强版 VLOOKUP 查找

```javascript
var result = JSA.z查找索引('关键字', arr, 1, false, '未找到');
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `关键字` | * | 查找关键字 |
| `数组` | Array | 二维数组 |
| `结果列` | Number | 结果列索引（从0开始） |
| `模式` | Boolean | 是否模糊匹配 |
| `错误值` | * | 找不到时返回的值 |

**返回:** `*` - 查找结果

---

#### z左侧查找 / vlookup

VLOOKUP 风格查找

```javascript
var result = JSA.z左侧查找('关键字', arr, 2, false);
```

---

#### z增强查找 / xlookup

XLOOKUP 风格查找

```javascript
var result = JSA.z增强查找('关键字', 查找数组, 结果数组, '未找到');
```

---

### WPS/Excel操作

#### z写入单元格 / toRange

将数组写入单元格区域

```javascript
JSA.z写入单元格(data, 'A1');
JSA.z写入单元格(data, Range('A1'), true);
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `arr` | Array | 二维数组 |
| `rng` | Range | 单元格区域 |
| `clearDown` | Boolean | 是否清空下方数据（默认false） |

**返回:** `Range` - 写入的区域

---

#### z转公式数组 / toExcelArray

将数组转换为公式字符串

```javascript
var formula = JSA.z转公式数组([[1,2],[3,4]]);
```

---

### 工具函数

#### z人民币大写 / rmbdx

数字转人民币大写

```javascript
var str = JSA.z人民币大写(1234.56);  // "壹仟贰佰叁拾肆元伍角陆分"
```

---

#### z随机整数 / rndInt

生成指定范围内的随机整数

```javascript
var rand = JSA.z随机整数(1, 100);  // 1-100 之间的随机整数
```

---

#### z随机整数数组 / rndIntArray

生成随机整数数组

```javascript
var arr = JSA.z随机整数数组(1, 100, 10);  // 10个 1-100 之间的随机整数
```

---

#### z随机小数 / rndFloat

生成随机小数

```javascript
var rand = JSA.z随机小数(0, 1, 2);  // 0-1 之间，保留2位小数
```

---

#### z随机小数数组 / rndFloatArray

生成随机小数数组

```javascript
var arr = JSA.z随机小数数组(0, 1, 5, 3);  // 5个 0-1 之间的小数，保留3位
```

---

#### z随机打乱 / shuffle

随机打乱数组顺序

```javascript
var shuffled = JSA.z随机打乱([1,2,3,4,5]);
```

---

#### z延时 / delay

延时执行

```javascript
await JSA.z延时(1000);  // 延时 1000 毫秒
```

---

#### z统一路径分隔符 / normalPath

统一路径分隔符

```javascript
var path = JSA.z统一路径分隔符('C:\\Users\\Name');  // "C:/Users/Name"
```

---

### Lambda / 函数式

#### jsaLambda

以字符串形式创建并执行函数

```javascript
var result = JSA.jsaLambda('x => x * 2', 5);     // 10
var result = JSA.jsaLambda('(x, y) => x + y', 3, 4);  // 7
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `fn` | String | Lambda 表达式字符串 |
| `...args` | * | 传递给函数的参数 |

**返回:** `*` - 函数执行结果

---

#### z解析函数表达式 / parseLambda

解析 Lambda 表达式为函数

```javascript
var fn = JSA.z解析函数表达式('x => x * 2');
fn(5);  // 10
```