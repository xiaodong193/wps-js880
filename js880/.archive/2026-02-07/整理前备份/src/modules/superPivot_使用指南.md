# superPivot v3.9.0 使用指南

## 快速开始

### 1. 基础透视表

```javascript
// 准备数据
var data = [
    ['产品', '年份', '地区', '销售额'],
    ['A', '2023', '华东', 100],
    ['A', '2023', '华南', 200],
    ['B', '2024', '华东', 150]
];

// 创建透视表
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],           // 行字段
    ['f2', '年份'],           // 列字段
    ['sum("f4")', '销售额']   // 数据字段
);

// 输出到工作表
result.toRange("A1");
```

### 2. 多层字段

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1,f2', '大区,省份'],   // 2个行字段
    ['f3,f4', '年份,季度'],   // 2个列字段
    ['sum("f5")', '销售额']
);
```

### 3. 带小计和总计

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],
    ['f2', '年份'],
    ['sum("f4")', '销售额'],
    1, 1, '@^@',
    {
        cornerTitle: '销售分析表',
        rowSubtotals: { enabled: true, label: '小计' },
        colSubtotals: { enabled: true, label: '小计' },
        grandTotals: { row: true, column: true, label: '总计' }
    }
);
```

### 4. 百分比显示

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],
    ['f2', '年份'],
    ['sum("f4")', '占比'],
    1, 1, '@^@',
    {
        displayAs: {
            mode: 'percentOfGrandTotal',  // 占总计百分比
            decimals: 2
        }
    }
);
```

## 参数详解

### options 对象

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| cornerTitle | String | '' | 左上角标题 |
| layoutMode | String | 'outline' | 布局模式：compact/outline/tabular |
| rowFieldIndent | Boolean | true | 是否启用层级缩进 |
| rowFieldIndentSize | Number | 4 | 缩进空格数 |
| rowSubtotals | Object | {enabled:false} | 行小计配置 |
| colSubtotals | Object | {enabled:false} | 列小计配置 |
| grandTotals | Object | {row:false,column:false} | 总计配置 |
| displayAs | Object | {mode:'value'} | 显示方式配置 |

### 排序符号

在字段配置中使用：
- `f1+` : 升序
- `f1-` : 降序
- `f1#` : 原始顺序

示例：`['f1+,f2-', '大区,省份']` 表示大区升序，省份降序

## 返回值方法

| 方法 | 说明 |
|------|------|
| toRange(rng, applyMerges) | 写入单元格 |
| getMeta() | 获取元数据 |
| val() / res() | 获取原始数组 |

## 调试技巧

```javascript
var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);

// 查看元数据
var meta = result.getMeta();
Console.log(JSON.stringify(meta, null, 2));

// 查看结果行数
Console.log("总行数: " + result.length);

// 查看第一行数据
Console.log("第一行: " + JSON.stringify(result[0]));
```
