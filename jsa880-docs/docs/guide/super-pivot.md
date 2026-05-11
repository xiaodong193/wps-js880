# superPivot 超级透视表

z超级透视 / superPivot 是 JSA880 框架最强大的功能，一行代码实现复杂的数据透视汇总。

## 基本语法

```javascript
Array2D.z超级透视(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options)
```

### 参数说明

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `arr` | Array | 必填 | 源数据二维数组 |
| `rowFields` | Array/String | 必填 | 行字段配置 |
| `colFields` | Array/String | 必填 | 列字段配置 |
| `dataFields` | Array/String | 必填 | 数据字段配置 |
| `headerRows` | Number | 1 | 源数据表头行数 |
| `outputHeader` | Number/String | 1 | 输出表头配置 |
| `separator` | String | '@^@' | 分隔符 |
| `options` | Object | {} | 高级选项 |

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

var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")']);
result.toRange('A1');
```

**结果：**
| 产品 | 中国 | 美国 |
|------|------|------|
| 手机 | 10000 | 8000 |
| 电脑 | 12000 | 15000 |

### 示例2：多列字段（多层表头）

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],        // 行字段
    ['f2,f3', '国家,年份'], // 多列字段
    ['sum("f4")', '销售额']
);
```

**结果（多层表头）：**
| 产品 | 2024 |  | 2025 |  |
|------|------|--|------|--|
|      | 中国 | 美国 | 中国 | 美国 |
| 手机 | 5000 | 4000 | 5000 | 4000 |
| 电脑 | 6000 | 7000 | 6000 | 8000 |

### 示例3：多数据字段

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

## 字段配置格式

### 行/列字段

```javascript
// 单字段
['f1']
['f1', '产品']  // 带标题

// 多字段（带排序）
['f1+,f2-']  // f1升序, f2降序
['f1,f2', '产品,国家']  // 多字段+标题

// f模式（列选择器）
'f1,f2,f3'  // 第1,2,3列
```

### 数据字段

```javascript
// 单个聚合
'sum("f3")'
'count()'
'average("f4")'

// 多个聚合
'sum("f3"),count(),average("f4")'

// 带标题
['sum("f3"),count()', '销售额,订单数']
```

## 聚合函数

| 函数 | 说明 | 示例 |
|------|------|------|
| `count()` | 计数 | `'count()'` |
| `sum(col)` | 求和 | `'sum("f3")'` |
| `average(col)` | 平均值 | `'average("f4")'` |
| `max(col)` | 最大值 | `'max("f5")'` |
| `min(col)` | 最小值 | `'min("f5")'` |
| `countDistinct(col)` | 去重计数 | `'countDistinct("f2")'` |

## 高级选项

```javascript
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 1, '@^@', {
    // 角标题
    cornerTitle: '销售汇总表',

    // 布局模式: 'outline' | 'compact' | 'tabular'
    layoutMode: 'outline',

    // 层级缩进
    rowFieldIndent: true,
    rowFieldIndentSize: 4,

    // 行小计
    rowSubtotals: {
        enabled: true,
        label: '小计'
    },

    // 列小计
    colSubtotals: {
        enabled: true,
        label: '小计'
    },

    // 总计
    grandTotals: {
        row: true,      // 行总计
        column: true,   // 列总计
        label: '总计'
    },

    // 百分比显示
    displayAs: {
        mode: 'percentOfGrandTotal',  // 'value' | 'percentOfGrandTotal' | 'percentOfRowTotal'
        decimals: 2
    }
});
```

## 链式调用

```javascript
// 筛选 → 排序 → 透视
data.z筛选('f3 > 5000')
   .z多列排序('f1+,f2-')
   .z超级透视(['f1'], ['f2'], ['sum("f3")'])
   .toRange('A1');

// 多步筛选
data.z筛选('f1 === "手机"')
   .z筛选('f2 === "中国"')
   .z超级透视(['f3'], ['f4'], ['sum("f5")']);
```

## 与 Excel 数据透视表的对比

| 功能 | Excel 透视表 | JSA880 superPivot |
|------|-------------|------------------|
| 多层行列 | 拖拽字段 | `['f1,f2', '标题']` |
| 聚合方式 | 值字段设置 | `['sum("f3"),count()']` |
| 筛选 | 切片器 | `.z筛选()` |
| 排序 | 手动拖拽 | `'f1+,f2-'` |
| 输出位置 | 拖拽到区域 | `.toRange('A1')` |
| 自动化 | VBA | JavaScript |

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
// 默认返回 0 或空，使用 NaN 模式保留空值
var result = Array2D.z超级透视(data, ['f1'], ['f2'], ['sum("f3")'], 1, 1, '@^@', {
    emptyAs: 'keep'  // 'zero' | 'keep' | 'null'
});
```