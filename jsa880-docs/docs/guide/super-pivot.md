# superPivot 超级透视表

`z超级透视 / superPivot` 是 JSA880 框架最强大的功能，一行代码实现复杂的数据透视汇总。支持多层行列字段、多聚合函数、小计与总计、百分比显示等多种高级特性。

## 基本语法

```javascript
Array2D.z超级透视(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options)
```

## 参数详解

### 核心参数

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `arr` | Array/Array2D | 必填 | 源数据二维数组，支持带表头的数据 |
| `rowFields` | Array/String | 必填 | 行字段配置，决定垂直方向的分类 |
| `colFields` | Array/String | 必填 | 列字段配置，决定水平方向的分类 |
| `dataFields` | Array/String | 必填 | 数据字段配置，指定聚合操作 |
| `headerRows` | Number | 1 | 源数据表头行数，透视时会跳过这些行 |
| `outputHeader` | Number/String | 1 | 输出表头配置：1=输出, 0=不输出, -1=输出但隐藏行标题, 'map'=返回Map |
| `separator` | String | '@^@' | 多值键的分隔符，用于连接行列字段值 |
| `options` | Object | {} | 高级选项配置对象 |

### 字段配置格式

#### f模式（列选择器）

使用 `f1`, `f2`, `f3` 等表示第1、2、3列，索引从1开始：

```javascript
// 基本用法
['f1']           // 第1列
['f1,f2']        // 第1列和第2列
['f1,f2,f3']     // 前3列

// 带排序后缀
['f1+']          // 第1列，升序（默认）
['f2-']          // 第2列，降序
['f1+,f2-']      // 第1列升序，第2列降序
['f3#']          // 第3列，按原始顺序（保持原顺序）
```

**排序后缀说明：**

| 后缀 | 含义 | 说明 |
|------|------|------|
| `+` | 升序 | 数字从小到大，字母A-Z |
| `-` | 降序 | 数字从大到小，字母Z-A |
| `#` | 原始顺序 | 保持数据出现的原始顺序 |

#### 带标题的字段配置

```javascript
// 数组格式：[字段配置, 标题配置]
['f1,f2', '产品,地区']      // f1显示为"产品"，f2显示为"地区"
['f1+,f2-', '产品↑,地区↓'] // 带排序的标题
['f1,f2,f3', '大区,省份,城市']  // 多层级标题
```

### 数据字段配置

#### 聚合函数

| 函数 | 说明 | 示例 | 默认标题 |
|------|------|------|----------|
| `sum(col)` | 求和 | `'sum("f3")'` | 求和 |
| `count()` | 计数（记录数） | `'count()'` | 计数 |
| `average(col)` | 平均值 | `'average("f3")'` | 平均 |
| `max(col)` | 最大值 | `'max("f3")'` | 最大 |
| `min(col)` | 最小值 | `'min("f3")'` | 最小 |
| `countDistinct(col)` | 去重计数 | `'countDistinct("f2")'` | 去重计数 |

#### 组合使用

```javascript
// 单个聚合
'sum("f3")'
'count()'

// 多个聚合（逗号分隔）
'sum("f3"),count(),average("f4")'
'count(),sum("f5"),average("f5"),max("f5"),min("f5")'

// 带自定义标题
['sum("f3"),count()', '销售额,订单数']
```

#### 动态列引用

在聚合函数中，可以使用列名或索引：

```javascript
// 使用列索引（f1, f2, f3...）
'sum("f3")'      // 第3列求和
'sum("f5"),count()'  // 第5列求和，计数

// 动态函数形式
function(g) { return g.sum("f3"); }
```

### options 高级选项详解

```javascript
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 1, '@^@', {
    // 角标题
    cornerTitle: '销售汇总表',

    // 布局模式: 'outline' | 'compact' | 'tabular'
    layoutMode: 'outline',

    // 层级缩进
    rowFieldIndent: true,         // 是否启用缩进
    rowFieldIndentSize: 4,        // 缩进空格数

    // 行小计（v3.9.0+）
    subtotals: {
        row: true,               // 启用行小计
        col: true,              // 启用列小计
        label: '小计'           // 小计标签
    },

    // 总计（v3.9.0+）
    grandTotal: {
        row: true,              // 启用行总计
        col: true,              // 启用列总计
        label: '总计'           // 总计标签
    },

    // 百分比显示
    displayAs: {
        mode: 'percentOfGrandTotal',  // 'value' | 'percentOfGrandTotal' | 'percentOfRowTotal' | 'percentOfColTotal'
        decimals: 2             // 小数位数
    },

    // 列筛选（v3.9.0+）
    filterCols: {
        f1: ['北京', '上海'],   // 只保留 f1 列中值为"北京"或"上海"的行
        f2: ['2024']            // 只保留 f2 列中值为"2024"的行
    },

    // 空值处理
    emptyAs: 'zero'              // 'zero' | 'keep' | 'null'
});
```

#### 兼容性说明

v3.9.0+ 使用新的配置名（`subtotals`, `grandTotal`），但也兼容旧名称：

```javascript
// 旧版配置名（仍支持，但推荐使用新名称）
rowSubtotals: { enabled: true }    // → subtotals: { row: true }
colSubtotals: { enabled: true }    // → subtotals: { col: true }
grandTotals: { row: true }        // → grandTotal: { row: true }
```

## 快速示例

### 示例1：基本透视

```javascript
var data = [
    ['产品', '国家', '销售额'],
    ['手机', '中国', 10000],
    ['手机', '美国', 8000],
    ['电脑', '中国', 12000],
    ['电脑', '美国', 15000]
];

// 基础透视：按产品统计各国销售额
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")']);
result.toRange('A1');
```

**结果：**
| 产品 | 中国 | 美国 |
|------|------|------|
| 手机 | 10000 | 8000 |
| 电脑 | 12000 | 15000 |

### 示例2：多层列字段（多层表头）

```javascript
var data = [
    ['产品', '年份', '季度', '销售额'],
    ['手机', 2024, 'Q1', 5000],
    ['手机', 2024, 'Q2', 4000],
    ['手机', 2025, 'Q1', 5000],
    ['手机', 2025, 'Q2', 4000],
    ['电脑', 2024, 'Q1', 6000],
    ['电脑', 2024, 'Q2', 7000],
    ['电脑', 2025, 'Q1', 6000],
    ['电脑', 2025, 'Q2', 8000]
];

// 年份和季度作为列字段，产生多层表头
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],                    // 行字段
    ['f2,f3', '年份,季度'],            // 多列字段：年份×季度
    ['sum("f4")', '销售额']            // 数据字段
);
```

**结果（多层表头）：**
| 产品 | 2024 |  | 2025 |  |
|      | Q1 | Q2 | Q1 | Q2 |
|------|------|--|------|--|
| 手机 | 5000 | 4000 | 5000 | 4000 |
| 电脑 | 6000 | 7000 | 6000 | 8000 |

### 示例3：多数据字段（多聚合）

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],
    ['f2', '国家'],
    ['sum("f3"),count(),average("f3")', '销售额,订单数,平均单价']
);
```

**结果：**
| 产品 | 国家 | 销售额 | 订单数 | 平均单价 |
|------|------|--------|--------|----------|
| 手机 | 中国 | 10000 | 5 | 2000 |
| 手机 | 美国 | 8000 | 4 | 2000 |
| ... | ... | ... | ... | ... |

### 示例4：带小计和总计

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1+', '产品'],                   // 按产品名升序
    ['f2+', '地区'],
    ['sum("f3")', '销售额'],
    1, 1, '@^@',
    {
        subtotals: { row: true, col: true, label: '小计' },
        grandTotal: { row: true, col: true, label: '总计' }
    }
);
```

### 示例5：百分比显示

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1+'],
    ['f2+'],
    ['sum("f3")'],
    1, 1, '@^@',
    {
        displayAs: { mode: 'percentOfGrandTotal', decimals: 1 }
    }
);
// 数值将显示为占总计的百分比，如 "25.0%"
```

### 示例6：多级行字段

```javascript
var data = [
    ['大区', '省份', '城市', '销售额'],
    ['华东', '江苏', '南京', 5000],
    ['华东', '江苏', '苏州', 4000],
    ['华东', '浙江', '杭州', 6000],
    ['华南', '广东', '广州', 7000],
    ['华南', '广东', '深圳', 8000]
];

var result = Array2D.z超级透视(
    data,
    ['f1,f2,f3', '大区,省份,城市'],   // 三级行字段
    [],
    ['sum("f4"),count()', '销售额,城市数'],
    1, 1, '@^@',
    {
        layoutMode: 'outline',
        subtotals: { row: true, label: '小计' },
        grandTotal: { row: true, label: '总计' }
    }
);
```

### 示例7：链式调用

```javascript
// 先筛选再透视
var result = data
    .z筛选('f2 === "中国"')           // 筛选中国数据
    .z多列排序('f1+,f2-')             // 排序
    .z超级透视(['f1'], ['f2'], ['sum("f3")'])
    .toRange('A1');

// 连续筛选
var result = data
    .z筛选('f1 === "手机"')
    .z筛选('f2 > 5000')
    .z超级透视(['f3'], ['f4'], ['sum("f5")']);
```

## 输出配置详解

### outputHeader 参数

| 值 | 含义 | 示例 |
|----|------|------|
| `1` 或 `true` | 输出完整表头（默认） | 包含行列标题和多层表头 |
| `0` | 不输出表头 | 仅输出数据行 |
| `-1` | 输出表头但隐藏行标题 | 仅显示列字段标题 |
| `'map'` | 返回Map格式 | 可通过 `result.get('键')` 访问 |

```javascript
// 不输出表头（仅数据）
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 0);

// 返回Map格式
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 'map');
result.get('手机@^@中国');  // 获取手机在中国的销售额
```

## 与 Excel 数据透视表的对比

| 功能 | Excel 透视表 | JSA880 superPivot |
|------|-------------|------------------|
| 多层行列 | 拖拽字段到行列区域 | `['f1,f2', '标题']` |
| 聚合方式 | 值字段设置对话框 | `['sum("f3"),count()']` |
| 筛选 | 切片器筛选 | `.z筛选()` |
| 排序 | 手动拖拽调整顺序 | `'f1+,f2-'` 或 `'f3#'` |
| 小计/总计 | 字段设置 | `subtotals: { row: true }` |
| 输出位置 | 拖拽到区域 | `.toRange('A1')` |
| 自动化 | 需要VBA录制 | JavaScript直接调用 |
| 百分比显示 | 值显示方式 | `displayAs: { mode: 'percentOfRowTotal' }` |

## 布局模式详解

### outline（大纲模式）

默认模式，行字段值在每行显示，便于查看多层级结构。

```javascript
var result = Array2D.z超级透视(data, ['f1,f2'], [], ['sum("f5")'], 1, 1, '@^@', {
    layoutMode: 'outline'
});
```

### compact（紧凑模式）

行字段合并显示，相同值只显示一次，更紧凑。

```javascript
var result = Array2D.z超级透视(data, ['f1,f2'], [], ['sum("f5")'], 1, 1, '@^@', {
    layoutMode: 'compact'
});
```

### tabular（表格模式）

每个单元格显示完整键值，适合导出或打印。

```javascript
var result = Array2D.z超级透视(data, ['f1,f2'], [], ['sum("f5")'], 1, 1, '@^@', {
    layoutMode: 'tabular'
});
```

## 常见问题

### Q: 如何只输出数据不要表头？

```javascript
// 设置 outputHeader = 0
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 0);
```

### Q: 如何返回 Map 格式而不是数组？

```javascript
// 设置 outputHeader = 'map'
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 'map');
// 返回 Map 对象，可通过键访问
result.get('手机@^@中国');
```

### Q: 如何处理没有对应值的情况？

```javascript
// 默认返回 0，使用 'keep' 保留空值
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 1, '@^@', {
    emptyAs: 'keep'  // 'zero' | 'keep' | 'null'
});
```

### Q: 如何自定义分隔符避免冲突？

```javascript
// 如果数据中包含 "@^@"，使用其他分隔符
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 1, '|||');

// 或者使用空分隔符
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 1, '');
```

### Q: 如何只显示某些列的值？

```javascript
// 使用 filterCols 筛选列值
var result = Array2D.z超级透视(data, ['f1'], ['f2,f3'], ['sum("f4")'], 1, 0, '@^@', {
    filterCols: {
        f1: ['北京', '上海'],   // 只保留北京和上海的数据
        f2: ['2024']             // 只保留2024年的数据
    }
});
```

### Q: 如何保持原始数据顺序？

```javascript
// 使用 '#' 后缀保持原始顺序
var result = Array2D.z超级透视(data, ['f1#'], ['f2#'], ['sum("f3")']);
// 结果将按照数据出现的原始顺序排列
```

### Q: 如何获取透视表的元数据？

```javascript
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")']);
if (result.getMeta) {
    var meta = result.getMeta();
    Console.log('行键:', meta.rowKeys);
    Console.log('列键:', meta.colKeys);
}
```

## 性能建议

1. **预处理大数据**：在透视前使用 `z筛选` 减少数据量
2. **避免过多层级**：超过3层的行列字段可能影响可读性
3. **合理使用分隔符**：确保数据中不包含分隔符，或使用空分隔符
4. **批量操作**：如需多次透视，可先对数据进行分组处理

## 版本历史

- **v3.9.1**: 修复多层表头生成的边界情况
- **v3.9.0**: 新增 `subtotals`/`grandTotal` 配置，百分比显示增强
- **v3.8.x**: 基础透视功能完善