# superPivot（z超级透视）API 参考文档

> JSA880 框架 v4.0.43 | 最后更新：2026-06-07

---

## 1. 函数签名

### 静态方法

```js
Array2D.z超级透视(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options)
Array2D.superPivot(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options)
$.superPivot(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options)
$.z超级透视(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options)
```

### 实例方法（用于链式调用）

```js
arr.z超级透视(rowFields, colFields, dataFields, headerRows, outputHeader, separator, options)
arr.superPivot(rowFields, colFields, dataFields, headerRows, outputHeader, separator, options)
```

### 简化调用（options 对象语法，v4.0.10+）

当第 5 参数为对象时，自动识别为新式调用：

```js
Array2D.z超级透视(arr, rowFields, colFields, dataFields, options)
```

---

## 2. 参数说明

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|---|---|---|---|---|
| `arr` | `Array \| Array2D \| Range` | ✅ | — | 二维源数据。支持普通数组、Array2D 实例、WPS Range 对象 |
| `rowFields` | `Array \| String` | ✅ | — | 行字段配置，见 [字段指定方式](#7-字段指定方式) |
| `colFields` | `Array \| String` | ✅ | — | 列字段配置 |
| `dataFields` | `Array \| String` | ✅ | — | 数据字段配置（含聚合函数） |
| `headerRows` | `Number \| Object` | ❌ | `1` | 标题行数。当传入对象时识别为 options（新式调用） |
| `outputHeader` | `Number \| String` | ❌ | `1` | 输出控制：`1`=含表头，`0`=无表头，`-1`=含表头不含行标题列，`'map'`=返回 Map |
| `separator` | `String` | ❌ | `'@^@'` | 多级字段键连接符 |
| `options` | `Object` | ❌ | `{}` | 高级选项，详见下文 |

---

## 3. options 对象详解

### 3.1 基础配置

| 属性 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `headerRows` | `Number` | `1` | 源数据标题行数（新式调用时有效） |
| `outputHeader` | `Number \| String` | `1` | 同参数 `outputHeader` |
| `separator` | `String` | `'@^@'` | 同参数 `separator` |
| `rowColSeparator` | `String` | `'\|\|\|'` | 行键与列键之间的分隔符 |
| `nullValue` | `any` | `0` | 缺失数据时的占位值 |
| `headerRowCount` | `Number` | 自动 | 输出表头行数。单列字段默认 2 行；多列字段默认 `colFields层级数 + 1` |

### 3.2 角标题与缩进

| 属性 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `cornerTitle` | `String` | 自动 | 左上角角标题文本。未指定时自动从 colConfig.titles 生成 |
| `rowFieldIndent` | `Boolean` | `true` | 是否启用层级缩进 |
| `rowFieldIndentSize` | `Number` | `4` | 缩进空格数 |
| `layoutMode` | `String` | `'outline'` | 布局模式：`'outline'` / `'compact'` / `'tabular'` |

### 3.3 小计与总计

#### subtotals（小计）

```js
options.subtotals = {
  enabled: false,   // 是否启用小计
  row: false,       // 行小计
  col: false,       // 列小计
  label: '小计'     // 小计标签
}
```

兼容旧版写法：
```js
options.rowSubtotals = { enabled: true }  // 等价于 subtotals.row = true
options.colSubtotals = { enabled: true }  // 等价于 subtotals.col = true
```

#### grandTotal（总计）

```js
options.grandTotal = {
  row: false,       // 行总计
  col: false,       // 列总计
  label: '总计'     // 总计标签
}
```

兼容旧版写法：
```js
options.grandTotals = { row: true, column: true }
```

### 3.4 筛选器（v3.9.4+）

在透视表输出前筛选特定行/列值。

#### filterRows — 筛选行键

支持两种写法：

**对象筛选**（按 fN 字段精确匹配）：
```js
filterRows: { f1: ['北京', '上海'] }   // 只保留第 1 列值为 '北京' 或 '上海' 的行
```

**函数筛选**（自定义逻辑）：
```js
filterRows: function(keyParts, key) {
  return keyParts[0] !== '深圳'        // 排除深圳
}
```

#### filterCols — 筛选列键

语法同 `filterRows`。

```js
filterCols: { f1: [2023, 2024] }     // 只保留年份为 2023 或 2024 的列
```

### 3.5 百分比显示（displayAs）

```js
options.displayAs = {
  mode: 'value',             // 'value' | 'percentOfGrandTotal' | 'percentOfRowTotal' | 'percentOfColTotal'
  decimals: 2                // 小数位数
}
```

---

## 4. 返回值说明

`superPivot` 返回一个 **Array2D 实例**（`wrappedResult`），除了可当普通二维数组使用外，还附加以下方法：

| 方法 | 返回值 | 说明 |
|---|---|---|
| `.val()` | `Array` | 返回对齐后的二维数组（空 cell 以 `''` 填充，WPS spill 不会显 0） |
| `.res()` | `Array` | `.val()` 的别名，返回原始对齐数组 |
| `.getMeta()` | `Object` | 返回透视表元数据（字段名、标题、总计值、选项等） |
| `.getMerges()` | `Object` | 返回表头合并信息 |
| `.toRange(rng, applyMerges?)` | `Range` | 将结果写入 WPS 单元格区域，自动应用表头合并 |
| `.getRange(rng, applyMerges?)` | `Range` | `.toRange()` 的别名 |
| `.applyMerges(rng)` | `Array` | 手动应用表头合并，返回已执行合并的列表 |
| `._header` | `Array` | 原始表头（如果可获取） |
| `._mergeInfo` | `Object` | 合并单元格信息 |

**元数据示例**（`.getMeta()` 返回值）：

```js
{
  version: '3.9.4',
  rowFields: ['f3', 'f2'],              // 行字段列表
  rowTitles: ['国家', '产品'],           // 行字段标题
  colFields: ['f6'],                    // 列字段列表
  colTitles: ['年'],                    // 列字段标题
  dataFields: ['count', 'sum', 'textjoin'],  // 聚合函数名
  dataTitles: ['计数', '求和', '多项合并'],  // 数据字段标题
  rowCount: 10,                         // 数据行数
  colCount: 4,                          // 列数
  headerRowCount: 2,                    // 表头行数
  grandTotal: [40, 95, '...'],          // 总计值
  options: { ... }                      // 所有 options 快照
}
```

---

## 5. 聚合函数列表

聚合函数以字符串形式在 `dataFields` 中指定，格式为 `函数名(参数)`。

| 函数 | 语法 | 说明 |
|---|---|---|
| `count` | `count()` | 计数（无参数） |
| `sum` | `sum("f4")` | 求和，参数为字段名。支持计算列表达式（v4.2.3+），如 `sum("f4*2")` |
| `average` | `average("f4")` | 平均值 |
| `max` | `max("f4")` | 最大值 |
| `min` | `min("f4")` | 最小值 |
| `textjoin` | `textjoin("f2", ",")` | 文本合并。第 1 参数为字段名，第 2 参数为分隔符 |
| `col` | `col("f5")` | 直接列引用（用于计算字段，返回 `null` 占位） |

### 语法细节

- 多个聚合函数用逗号连接：`"count(),sum(\"f4\"),average(\"f4\")"`
- `textjoin` 的第二个参数支持带引号的字符串：`"textjoin(\"f2\", \",\")"`
- 字段参数中的引号：在 WPS 公式中使用反引号（自动转为双引号），在 JSA 代码中直接使用双引号
- **计算列表达式**（v4.2.3+）：`sum("f4*2")` 表示对 f4 列值乘以 2 后求和；`sum("f3+f4")` 表示对两列之和求和

### 回调模式（Lambda 函数）

除字符串语法外，还支持传入回调函数数组：

```js
dataFields: [[
  g => g.count(),
  g => g.sum("f4"),
  g => g.average("f4")
], '计数,求和,平均']
```

回调函数接收的 `g` 对象提供以下方法：

| 方法 | 说明 |
|---|---|
| `g.count()` | 分组行数 |
| `g.sum(col)` | 求和（支持 `"f4"` 或 `"f4*2"` 计算列） |
| `g.average(col)` | 平均值 |
| `g.max(col)` | 最大值 |
| `g.min(col)` | 最小值 |
| `g.textjoin(col, sep)` | 文本合并 |

---

## 6. 排序语法

### 6.1 字段排序符号

在 `rowFields` 和 `colFields` 中，每个字段名后可以跟排序符号：

| 符号 | 含义 | 示例 |
|---|---|---|
| `+` | 升序（默认，可省略） | `f1+` 或 `f1` |
| `-` | 降序 | `f1-` |
| `#` | 保持原始出现顺序 | `f1#` |

```js
// 行字段：国家升序，产品降序
rowFields: ['f3+,f2-']

// 列字段：年份按原始出现顺序
colFields: ['f6#']

// 字符串形式同样支持
rowFields: 'f3+,f2-'
```

### 6.2 多列排序

对于数组数据本身的排序，可使用链式方法 `z多列排序` / `sortByCols`：

```js
arr.sortByCols('f3+,f4-', 1)   // 第3列升序，第4列降序，跳过1行表头
arr.z多列排序('f1+,f2-', 0, '中国,美国,英国')  // 支持自定义序列
```

---

## 7. 字段指定方式

### 7.1 fN 列选择器

用 `f1`、`f2`、`f3`… 引用第 1、2、3… 列（1-based 索引）。

```js
rowFields: ['f3,f2']     // 第3列和第2列作为行字段
colFields: 'f6'          // 第6列作为列字段
```

### 7.2 字符串指定（带排序和标题）

```js
// 字段 + 排序
rowFields: 'f1+,f2-'     // f1升序, f2降序

// 字段 + 自定义标题
rowFields: ['f3,f2', '国家,产品']    // [字段列表, 标题列表]
colFields: ['f6', '年份']            // [字段列表, 标题]

// 字段 + 排序 + 自定义标题（v3.8.3+）
rowFields: ['f3+,f2-', '国家,产品']
```

### 7.3 dataFields 的多种写法

```js
// 方式1：纯字符串
dataFields: 'count(),sum("f4")'

// 方式2：字符串数组
dataFields: ['count(),sum("f4"),average("f4")']

// 方式3：字符串 + 自定义标题
dataFields: ['count(),sum("f4")', '计数,求和']

// 方式4：回调函数数组 + 标题
dataFields: [[
  g => g.count(),
  g => g.sum("f4"),
  g => g.average("f4")
], '计数,求和,平均']
```

---

## 8. 完整使用示例

### 示例 1：基础透视（单行字段 + 单列字段）

```js
// 数据: JSA_Arr (ID, 产品, 国家, 数量, 价格, 年, 月, 日)
// 需求: 按国家（行）× 年（列）透视，统计数量总和

var result = Array2D.z超级透视(
  JSA_Arr,           // 源数据
  ['f3'],            // 行字段: 第3列(国家)
  ['f6'],            // 列字段: 第6列(年)
  ['sum("f4")'],     // 数据: 第4列(数量)求和
  1                  // 1 行标题
);
```

输出表头结构（单列字段，2 行表头）：

| (角标题) | 年 | 2021 | 2022 | 2023 | 2024 |
|---|---|---|---|---|---|
| 国家 | 求和 | 求和 | 求和 | 求和 | 求和 |
| 中国 | 19 | 13 | 19 | 11 | ... |
| 德国 | 18 | ... | ... | ... | ... |

### 示例 2：多行字段 + 多聚合

```js
// 需求: 按国家、产品（行）× 年（列），统计数量、均价

var result = Array2D.z超级透视(
  JSA_Arr,
  ['f3,f2', '国家,产品'],                // 行字段 + 自定义标题
  ['f6', '年份'],                         // 列字段
  ['sum("f4"),average("f4")', '总数量,均价'],  // 两聚合 + 标题
  1
);
```

### 示例 3：回调模式 + Map 返回

```js
// 需求: 用 Lambda 自定义聚合逻辑，返回 Map 供程序查询

var result = Array2D.z超级透视(
  JSA_Arr,
  ['f3', '国家'],
  ['f6', '年份'],
  [[
    g => g.count(),
    g => g.sum("f4"),
    g => g.max("f4")
  ], '订单数,总数量,最大数量'],
  1,
  'map'                // 返回 Map 而非 2D 数组
);

// 遍历结果
result.forEach(function(v, k) {
  console.log(k, v.agg, v.group);
});
```

### 示例 4：降序排序 + 无表头输出

```js
// 需求: 国家降序，产品升序，只输出数据行

var result = Array2D.z超级透视(
  JSA_Arr,
  ['f3-,f2+'],       // 国家降序, 产品升序
  ['f6'],
  ['count()'],
  1,
  0                   // 不输出表头
);
```

### 示例 5：筛选 + 小计 + 总计

```js
// 需求: 只保留中国和德国的数据，开启行总计和列总计

var result = Array2D.z超级透视(
  JSA_Arr,
  ['f3', '国家'],
  ['f6', '年份'],
  ['sum("f4")', '数量'],
  {
    headerRows: 1,
    filterRows: { f1: ['中国', '德国'] },   // 只保留这两国
    subtotals: { enabled: true, row: true },  // 行小计
    grandTotal: { row: true, col: true, label: '合计' },  // 行列总计
    nullValue: 0                              // 空值填 0
  }
);
```

### 示例 6：多列字段（两级表头）

```js
// 需求: 行=国家，列=年+月（两级列字段），统计数量

var result = Array2D.z超级透视(
  JSA_Arr,
  ['f3', '国家'],
  ['f6,f7', '年,月'],     // 列字段: 年 + 月（两级）
  ['sum("f4")', '数量'],
  1
);
```

多列字段时会生成多行表头（`numColFieldLevels + 1` 行），上层为年份、下层为月份。

### 示例 7：与 k() 函数配合（WPS 单元格公式）

```
// WPS 单元格公式 (无 __KJ_ARGS__)
=k("$$.superPivot", A1:H40, "f3,f2", "f6", "count(),sum("f4")")

// WPS 单元格公式 (含 __KJ_ARGS__，解决丢参问题)
=k("__KJ_ARGS__={`rowFields`:`f3,f2`,`colFields`:`f6`}  (...args)=>$$.superPivot(...args)", A1:H40, "count(),sum(`f4`)")
```

### 示例 8：链式调用 — 透视后筛选

```
// WPS 公式：透视后只保留 Product1 的行
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')", A1:H40, "f3,f2", "f6", "count()")
```

### 示例 9：百分比显示

```js
// 需求: 显示每项占总计的百分比

var result = Array2D.z超级透视(
  JSA_Arr,
  ['f3', '国家'],
  ['f6', '年份'],
  ['sum("f4")', '数量'],
  {
    headerRows: 1,
    grandTotal: { row: true, col: true },
    displayAs: {
      mode: 'percentOfGrandTotal',   // 占总计百分比
      decimals: 1                     // 保留 1 位小数
    }
  }
);
```

---

## 9. 与 k() 函数配合使用

### 9.1 基本用法

在 WPS 单元格中，`k()` 函数用于调用 JSA880 框架方法：

```
=k("$$.superPivot", 源数据区域, rowFields, colFields, dataFields)
```

`$$` 在 `k()` 的 Lambda 上下文中指向 `JSA` 命名空间，等价于 `$.superPivot`。

### 9.2 __KJ_ARGS__ 语法（v4.0.23+）

**背景**：WPS 公式引擎在传多个字符串参数时可能静默丢弃个别参数，导致 `colFields` 等字段丢失。

**解决方案**：将字段参数通过 JSON 标记嵌入到第一个字符串参数中。

```
// 语法
=k("__KJ_ARGS__={`rowFields`:`f3,f2`,`colFields`:`f6`,`dataFields`:`count()`}  (...args)=>$$.superPivot(...args)", A1:H40)
```

**规则说明**：

| 键名 | 对应参数 | 说明 |
|---|---|---|
| `rowFields` | 第 2 参数 | 行字段配置字符串 |
| `colFields` | 第 3 参数 | 列字段配置字符串 |
| `dataFields` | 第 4 参数 | 数据字段配置字符串 |
| `headerRows` | 第 5 参数 | 标题行数 |

- 反引号 `` ` `` 在公式中代替双引号（WPS 不支持公式内反引号时会自动转换）
- `__KJ_ARGS__` 标记之后的 Lambda 函数按正常逻辑解析
- 无 `__KJ_ARGS__` 标记的公式完全不受影响（向后兼容）

### 9.3 链式调用语法

```
// 基本链式
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='某值')", Range, ...)

// 多步链式
=k("(...args)=>$$.superPivot(...args).map(x=>[x.f1, x.f2*2])", Range, ...)
```

---

## 10. 链式调用

`superPivot` 返回 Array2D 实例，可直接链式调用以下方法：

### 10.1 filter / z筛选

```js
// 保留表头行 (i==0) 和 f2 列等于 'Product1' 的行
result.filter((x, i) => i === 0 || x.f2 === 'Product1')

// 字符串条件（需通过 _filterByObject）
result.filter('f2>0')
```

支持通过 `x.f1`、`x.f2`… 访问器访问列值（数组元素同时具有数字索引和 fN 字符串属性）。

### 10.2 sortByCols / z多列排序

```js
// 按第 1 列升序，第 2 列降序，跳过 1 行表头
result.sortByCols('f1+,f2-', 1)

// 带自定义序列
result.sortByCols('f3', 1, '中国,美国,德国,英国')
```

### 10.3 map / z映射

```js
// 对每行进行映射转换
result.map((x, i) => [x.f1, x.f2, x.f3 * 2])

// 字符串 Lambda
result.map('x=>[x.f1, x.f2]')
```

### 10.4 其他链式方法

| 方法 | 说明 |
|---|---|
| `.sort(fn)` | 原生数组排序 |
| `.sortBy(fn)` | 按规则升序 |
| `.sortByDesc(fn)` | 按规则降序 |
| `.sortRow(colIdx, asc)` | 按指定列排序行 |
| `.sortCol(rowIdx, asc)` | 按指定行排序列 |
| `.distinct(col?)` | 去重 |
| `.val()` | 返回纯数组（终止链式） |
| `.res()` | `.val()` 别名 |

---

## 11. smartGroup 智能分组

`smartGroup` 是独立于 `superPivot` 的分组辅助函数，用于按日期或数值维度对数据进行预分组。

### 静态方法

```js
Array2D.smartGroup(arr, col, groupBy)
Array2D.z智能分组(arr, col, groupBy)
```

### 实例方法

```js
arr.z智能分组(col, groupBy)
```

### 参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `arr` | `Array` | 二维源数据 |
| `col` | `Number \| String` | 列索引（数字 0-based）或列选择器（`'f2'` 1-based） |
| `groupBy` | `String` | 分组方式 |

### 日期分组方式

| 值 | 说明 | 输出键示例 |
|---|---|---|
| `'year'` / `'年'` | 按年 | `"2023年"` |
| `'quarter'` / `'季度'` | 按季度 | `"2023年Q2"` |
| `'month'` / `'月'` | 按月 | `"2023年6月"` |
| `'week'` / `'周'` | 按周 | `"2023年第24周"` |
| `'day'` / `'日'` | 按日 | `"2023-6-16"` |

### 数值分组方式

| 值 | 说明 | 输出键示例 |
|---|---|---|
| `'decade'` / `'十位数'` | 十位分组 | `"10-19"` |
| `'hundred'` / `'百位数'` | 百位分组 | `"100-199"` |
| `'thousand'` / `'千位数'` | 千位分组 | `"1000-1999"` |

### 返回值

返回 `Map` 对象，键为分组标签，值为该组的行数组。

### 使用示例

```js
// 按年份分组
var groups = Array2D.smartGroup(JSA_Arr, 'f6', 'year');
// groups: Map { "2021年" => [...], "2022年" => [...], ... }

// 按季度分组
var groups = arr.z智能分组('f6', 'quarter');

// 按十位数分组
var groups = Array2D.smartGroup(arr, 'f4', 'decade');
```

---

## 12. 已知限制和边界情况

### 12.1 数据相关

| 限制 | 说明 |
|---|---|
| **空行自动过滤** | 所有单元格均为 `null`/`undefined`/`''` 的行会被自动跳过 |
| **空字段键过滤** | 行/列字段中所有值均为空的组合会被跳过 |
| **无行字段时** | 若有数据但无有效行字段键（v4.2.5+），自动归入单一空键行 |
| **Range 对象** | 自动检测 WPS Range 并转换为 `Value2`；macOS WPS 的 function-type Range 也受支持（v4.0.43） |
| **Host Array** | WPS 的 host array（`slice`/`filter` 不可用）自动通过 JSON 往返转换为真数组 |

### 12.2 表头相关

| 限制 | 说明 |
|---|---|
| **多级表头对齐** | 各行长短不一时，`.val()` 自动对齐到最长行，空位用 `''` 填充（避免 WPS spill 显 0） |
| **单列字段表头** | 默认 2 行表头（可通过 `options.headerRowCount` 覆盖） |
| **多列字段表头** | 默认 `numColFieldLevels + 1` 行 |
| **合并单元格** | `.toRange()` 自动应用表头合并；`.applyMerges()` 可手动控制 |

### 12.3 功能相关

| 限制 | 说明 |
|---|---|
| **不支持多 dataFields 组合的独立聚合** | 所有 dataFields 共享同一套行/列分组逻辑 |
| **小计功能受限** | 行小计在 `enabled` 后仅插入占位，完整层级小计需外层处理 |
| **smartGroup 并非内置** | 日期分组需在调用 `superPivot` **前**用 `smartGroup` 预处理，pivot 自身不处理日期 |
| **`#` 排序符号** | 在多级字段中，非首级字段的 `#` 符号用于比较原始索引保持顺序 |
| **显示模式** | `displayAs` 的百分比模式需要先开启对应方向的 `grandTotal` |

### 12.4 WPS 公式相关

| 限制 | 说明 |
|---|---|
| **字符串参数丢失** | WPS 公式传多个字符串参数时可能静默丢弃，需用 `__KJ_ARGS__` 标记解决 |
| **反引号处理** | WPS 公式中的反引号 `` ` `` 自动转为双引号 `"` |
| **链式调用中的类型** | `.filter()`/`.map()` 在链式中通过 fN 代理访问列，需确保数据行转换为 Array2D 或带 fN 代理 |

---

## 13. 更新历史摘要

| 版本 | 日期 | 关键变更 |
|---|---|---|
| v3.7.5 | — | 自动表头检测；支持 Array2D 对象 `_header` 传递 |
| v3.7.9 | — | 表头布局方案 3；数据行与表头对齐 |
| v3.8.0 | — | 多行多列表头支持 |
| v3.8.3 | — | 排序符号 + 自定义标题同时支持 |
| v3.8.6 | — | 用户自定义 dataFields 标题优先 |
| v3.8.8 | — | `outputHeader = -1` 隐藏行标题列 |
| v3.9.0 | — | 小计、总计、百分比显示、缩进 |
| v3.9.1 | — | 无列字段场景支持 |
| v3.9.4 | — | filterRows / filterCols 内置筛选器 |
| v3.9.5 | — | 多列字段表头行字段标题位置修正 |
| v4.0.10 | — | 新式调用（options 为第 5 参数）；数字格式化；排序 Schwartzian 优化 |
| v4.0.18 | — | WPS Range + host array 双重转换 |
| v4.0.20 | — | `.val()` 对齐补齐，空位用 `''` |
| v4.0.23 | — | `__KJ_ARGS__` JSON 标记支持（解决 WPS 丢参） |
| v4.0.27 | — | `__KJ_ARGS__` 宽松 JSON 解析 |
| v4.0.28 | — | `__KJ_ARGS__` 引号内逗号保护 |
| v4.0.30 | — | fN 代理字符串自动 trim（修复 `.filter()` 匹配失败） |
| v4.0.31 | — | 单列字段表头每 colKey × numDataFields 重复对齐 |
| v4.0.32 | — | 单列字段表头改为 2 行 |
| v4.0.33 | — | `.val()` 空位用 `''` 代替 `null` |
| v4.0.34 | — | 默认 `nullValue` 从 `''` 改为 `0` |
| v4.0.35 | — | 单列字段表头行 0 row 字段位置不再重复 push |
| v4.0.36 | — | 单列字段表头行 1 col 字段标题位置不 push（避免重复） |
| v4.0.43 | — | macOS WPS function-type Range 支持 |
| v4.2.3 | — | 聚合函数支持计算列表达式（`f4*2`、`f3+f4`） |
| v4.2.5 | — | 无行字段时自动归入单一空键行 |

