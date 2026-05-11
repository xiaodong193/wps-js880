# superPivot v3.9.0 完整使用指南

## 📖 目录
1. [快速入门](#一快速入门)
2. [基础用法](#二基础用法)
3. [参数详解](#三参数详解)
4. [功能示例](#四功能示例)
5. [实际应用场景](#五实际应用场景)
6. [常见问题](#六常见问题)

---

## 一、快速入门

### 1.1 最简单示例

```javascript
// 步骤1: 准备数据（第一行是表头）
var data = [
    ['产品', '年份', '地区', '销售额'],
    ['手机', '2023', '华东', 1000],
    ['手机', '2023', '华南', 2000],
    ['电脑', '2024', '华东', 1500],
    ['电脑', '2024', '华南', 2500]
];

// 步骤2: 创建透视表
var result = Array2D.z超级透视(
    data,                    // 数据源
    ['f1', '产品'],          // 行字段：第1列，标题"产品"
    ['f2', '年份'],          // 列字段：第2列，标题"年份"
    ['sum("f4")', '销售额']  // 数据字段：第4列求和，标题"销售额"
);

// 步骤3: 输出到工作表
result.toRange("A1");
```

**输出效果：**
```
产品      2023    2024    
手机      3000            
电脑              4000    
```

---

## 二、基础用法

### 2.1 函数签名

```javascript
Array2D.z超级透视(
    arr,           // 数据源（二维数组或Array2D对象）
    rowFields,     // 行字段配置
    colFields,     // 列字段配置
    dataFields,    // 数据字段配置
    headerRows,    // 表头行数（默认1）
    outputHeader,  // 是否输出表头（默认1）
    separator,     // 分隔符（默认"@^@"）
    options        // 高级选项（v3.9.0新增）
)
```

### 2.2 字段配置格式

#### 行/列字段配置

**格式1：简单字符串**
```javascript
'f1,f2,f3'  // 第1、2、3列
```

**格式2：带标题的数组**
```javascript
['f1,f2', '大区,省份']  // 字段+自定义标题
```

**格式3：带排序符号**
```javascript
['f1+,f2-', '大区,省份']  // +升序，-降序，#原始顺序
```

#### 数据字段配置

**格式1：聚合函数字符串**
```javascript
'count(),sum("f4"),average("f5")'
```

**格式2：带标题的数组**
```javascript
['count(),sum("f4")', '订单数,销售额']
```

**支持的聚合函数：**
- `count()` - 计数
- `sum("fN")` - 求和
- `average("fN")` - 平均值
- `max("fN")` - 最大值
- `min("fN")` - 最小值

---

## 三、参数详解

### 3.1 options 对象（v3.9.0新增）

```javascript
{
    // 角标题（左上角显示的标题）
    cornerTitle: '销售分析表',
    
    // 布局模式
    layoutMode: 'outline',  // 'compact' | 'outline' | 'tabular'
    
    // 行字段层级缩进
    rowFieldIndent: true,       // 是否启用缩进
    rowFieldIndentSize: 4,      // 缩进空格数
    
    // 行小计
    rowSubtotals: {
        enabled: true,          // 是否启用
        position: 'bottom',     // 位置：'bottom'（固定）
        label: '小计',          // 显示标签
        aggregation: 'sum'      // 聚合方式
    },
    
    // 列小计
    colSubtotals: {
        enabled: true,
        position: 'right',      // 位置：'right'（固定）
        label: '小计'
    },
    
    // 总计
    grandTotals: {
        row: true,              // 显示行总计
        column: true,           // 显示列总计
        label: '总计'           // 标签文字
    },
    
    // 显示方式（百分比）
    displayAs: {
        mode: 'percentOfGrandTotal',  // 显示模式
        decimals: 2                    // 小数位数
    }
}
```

### 3.2 显示模式（displayAs.mode）

| 模式 | 说明 | 示例 |
|------|------|------|
| `'value'` | 原始值（默认） | 1000 |
| `'percentOfGrandTotal'` | 占总计百分比 | 25.50% |
| `'percentOfRowTotal'` | 占行总计百分比 | 40.00% |
| `'percentOfColTotal'` | 占列总计百分比 | 33.33% |

---

## 四、功能示例

### 4.1 多层行字段

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1,f2', '大区,省份'],      // 2层行字段
    ['f3', '年份'],
    ['sum("f5")', '销售额']
);
```

**输出效果：**
```
大区    省份    2023    2024
华东    江苏    1000    1500
        浙江    2000    2500
华南    广东    3000    3500
```

### 4.2 多层列字段

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],
    ['f2,f3', '年份,季度'],      // 2层列字段
    ['sum("f4")', '销售额']
);
```

**输出效果：**
```
产品    2023                2024
        Q1      Q2      Q1      Q2
手机    1000    2000    1500    2500
电脑    3000    4000    3500    4500
```

### 4.3 带小计和总计

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1,f2', '大区,省份'],
    ['f3', '年份'],
    ['sum("f4")', '销售额'],
    1, 1, '@^@',
    {
        cornerTitle: '销售分析',
        rowSubtotals: { enabled: true, label: '小计' },
        colSubtotals: { enabled: true, label: '小计' },
        grandTotals: { row: true, column: true, label: '总计' }
    }
);
```

**输出效果：**
```
销售分析        2023        2024        小计
大区    省份    
华东    江苏    1000        1500        2500
        浙江    2000        2500        4500
        小计    3000        4000        7000
华南    广东    3000        3500        6500
        小计    3000        3500        6500
总计            6000        7500        13500
```

### 4.4 百分比显示

```javascript
// 占总计百分比
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],
    ['f2', '年份'],
    ['sum("f3")', '占比'],
    1, 1, '@^@',
    {
        displayAs: { mode: 'percentOfGrandTotal', decimals: 2 }
    }
);

// 输出：25.50% 表示该值占总销售额的25.5%
```

### 4.5 多种聚合函数

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],
    ['f2', '年份'],
    ['count(),sum("f3"),average("f3"),max("f3"),min("f3")', 
     '订单数,总销售额,平均单价,最高单价,最低单价']
);
```

### 4.6 带排序

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1+,f2-', '大区,省份'],  // 大区升序，省份降序
    ['f3+', '年份'],            // 年份升序
    ['sum("f4")', '销售额']
);
```

---

## 五、实际应用场景

### 场景1：销售数据分析

```javascript
function 销售分析报表() {
    // 从工作表读取数据
    var ws = Application.ActiveWorkbook.Worksheets("销售数据");
    var data = ws.Range("A1:E1000").Value2;
    
    // 创建透视表
    var result = Array2D.z超级透视(
        data,
        ['f3,f4', '大区,省份'],      // 按地区分析
        ['f2,f5', '年份,月份'],      // 按时间分析
        ['sum("f6"),count()', '销售额,订单数'],
        1, 1, '@^@',
        {
            cornerTitle: '销售分析报表',
            grandTotals: { row: true, column: true },
            displayAs: { mode: 'value' }
        }
    );
    
    // 输出到新工作表
    var newWs = Application.ActiveWorkbook.Worksheets.Add();
    newWs.Name = "销售分析";
    result.toRange("A1", true);
    
    Console.log("报表已生成：销售分析");
}
```

### 场景2：库存周转分析

```javascript
function 库存周转分析() {
    var data = 获取数据("库存!A1:F500");
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f2', '仓库,商品类别'],
        ['f3', '月份'],
        ['sum("f4"),average("f5"),sum("f6")', 
         '入库数量,平均库存,出库数量'],
        1, 1, '@^@',
        {
            cornerTitle: '库存周转分析',
            rowSubtotals: { enabled: true, label: '类别合计' },
            colSubtotals: { enabled: true, label: '月度合计' }
        }
    );
    
    result.toRange("分析!A1");
}
```

### 场景3：员工业绩排名

```javascript
function 员工业绩报表() {
    var data = 获取数据("业绩!A1:D1000");
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f2', '部门,员工'],
        ['f3', '季度'],
        ['sum("f4"),max("f4"),min("f4")', 
         '总业绩,最高业绩,最低业绩'],
        1, 1, '@^@',
        {
            cornerTitle: '员工业绩分析',
            displayAs: { mode: 'percentOfColTotal', decimals: 1 }
        }
    );
    
    result.toRange("业绩分析!A1");
}
```

### 场景4：财务报表

```javascript
function 财务报表() {
    var data = 获取数据("财务!A1:G500");
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f2', '科目类型,科目'],
        ['f3,f4', '年份,月份'],
        ['sum("f5"),sum("f6"),sum("f7")', 
         '预算,实际,差额'],
        1, 1, '@^@',
        {
            cornerTitle: '财务预算执行分析',
            layoutMode: 'outline',
            rowFieldIndent: true,
            grandTotals: { row: true, column: true }
        }
    );
    
    result.toRange("财务分析!A1");
}
```

---

## 六、常见问题

### Q1: 如何获取透视表的元数据？

```javascript
var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);
var meta = result.getMeta();

Console.log("行数: " + meta.rowCount);
Console.log("列数: " + meta.colCount);
Console.log("配置: " + JSON.stringify(meta.options));
```

### Q2: 如何不输出表头？

```javascript
var result = Array2D.z超级透视(
    data,
    rowFields,
    colFields,
    dataFields,
    1,      // headerRows
    0       // outputHeader = 0 不输出表头
);
```

### Q3: 如何只输出数据，不包含行标题列？

```javascript
var result = Array2D.z超级透视(
    data,
    rowFields,
    colFields,
    dataFields,
    1,
    -1      // outputHeader = -1 输出表头但不包含行标题
);
```

### Q4: 如何处理大量数据？

```javascript
// 方法1：分批处理
var allData = 获取数据("A1:Z10000");
var batchSize = 1000;

for (var i = 0; i < allData.length; i += batchSize) {
    var batch = allData.slice(i, i + batchSize);
    // 处理批次...
}

// 方法2：禁用屏幕更新（toRange内部已自动处理）
var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);
result.toRange("A1");  // 自动优化性能
```

### Q5: 百分比显示不正确？

确保数据是数值类型：
```javascript
// 检查数据
for (var i = 0; i < data.length; i++) {
    var val = parseFloat(data[i][3]);
    if (isNaN(val)) {
        Console.log("第" + (i+1) + "行数据不是数值: " + data[i][3]);
    }
}
```

### Q6: 如何导出为HTML？

```javascript
var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);
var html = result.z输出HTML({
    header: true,
    className: 'pivot-table',
    style: 'border-collapse: collapse;'
});

// 保存为文件
IO.z写文件("C:\\report.html", html);
```

---

## 七、最佳实践

### 1. 数据准备
- 确保第一行是表头
- 确保数据区域没有空行
- 数值列不要包含文本

### 2. 性能优化
- 大数据量时禁用屏幕更新
- 使用合适的数据类型
- 避免过多的层级（建议行+列不超过5层）

### 3. 错误处理
```javascript
try {
    var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);
    result.toRange("A1");
} catch (e) {
    Console.log("错误: " + e.message);
    Console.log("堆栈: " + e.stack);
}
```

### 4. 调试技巧
```javascript
// 查看中间结果
var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);
Console.log("总行数: " + result.length);
Console.log("表头: " + JSON.stringify(result[0]));
Console.log("第一行数据: " + JSON.stringify(result[1]));

// 查看元数据
var meta = result.getMeta();
Console.log(JSON.stringify(meta, null, 2));
```

---

**文档版本**: v3.9.0  
**最后更新**: 2026-02-06
