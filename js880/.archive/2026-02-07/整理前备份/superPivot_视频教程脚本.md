# superPivot v3.9.0 视频教程脚本

## 教程1：从零开始（5分钟）

### 步骤1：准备数据（1分钟）

打开 WPS 表格，准备如下数据：

```
A列      B列      C列      D列
产品     年份     地区     销售额
手机     2023     华东     1000
手机     2023     华南     2000
手机     2024     华东     1500
电脑     2023     华南     3000
电脑     2024     华东     2500
电脑     2024     华南     3500
```

**操作**：在 Sheet1 中输入以上数据

---

### 步骤2：打开宏编辑器（1分钟）

1. 按 `Alt + F11` 打开宏编辑器
2. 点击「插入」→「模块」
3. 确保已加载 JSA880.js

**操作演示**：
```javascript
// 测试 JSA880 是否加载成功
Console.log(typeof Array2D);  // 应该输出 "function"
```

---

### 步骤3：编写第一个透视表代码（2分钟）

**代码**：
```javascript
function 我的第一个透视表() {
    // 1. 获取数据
    var ws = Application.ActiveWorkbook.Worksheets("Sheet1");
    var data = ws.Range("A1:D7").Value2;
    
    Console.log("数据行数: " + data.length);
    Console.log("第一行: " + JSON.stringify(data[0]));
    
    // 2. 创建透视表
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品'],          // 行：产品
        ['f2', '年份'],          // 列：年份
        ['sum("f4")', '销售额']  // 数据：销售额求和
    );
    
    // 3. 查看结果
    Console.log("结果行数: " + result.length);
    
    // 4. 输出到新工作表
    var newWs = Application.ActiveWorkbook.Worksheets.Add();
    newWs.Name = "透视结果";
    result.toRange("A1");
    
    Console.log("✅ 透视表已生成！");
}

// 运行
我的第一个透视表();
```

**预期结果**：
```
产品      2023    2024
手机      3000    1500
电脑      3000    6000
```

---

### 步骤4：查看结果（1分钟）

1. 查看「透视结果」工作表
2. 观察行列结构
3. 验证数据是否正确

---

## 教程2：添加小计和总计（5分钟）

### 目标

在上一个透视表基础上，添加：
- 行小计
- 列小计  
- 总计

### 代码

```javascript
function 带小计的透视表() {
    var ws = Application.ActiveWorkbook.Worksheets("Sheet1");
    var data = ws.Range("A1:D7").Value2;
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品'],
        ['f2', '年份'],
        ['sum("f4")', '销售额'],
        1,      // headerRows
        1,      // outputHeader
        '@^@',  // separator
        {
            // 角标题
            cornerTitle: '销售分析报表',
            
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
                row: true,
                column: true,
                label: '总计'
            }
        }
    );
    
    // 输出
    var newWs = Application.ActiveWorkbook.Worksheets.Add();
    newWs.Name = "小计报表";
    result.toRange("A1", true);  // true = 应用合并
    
    // 查看元数据
    var meta = result.getMeta();
    Console.log("行数: " + meta.rowCount);
    Console.log("列数: " + meta.colCount);
    Console.log("总计: " + meta.grandTotal);
}

带小计的透视表();
```

### 预期结果

```
销售分析报表      2023    2024    小计
产品                      
手机            3000    1500    4500
电脑            3000    6000    9000
小计            6000    7500    13500
```

---

## 教程3：多层字段分析（5分钟）

### 目标

实现多层行字段和列字段的复杂分析

### 数据准备

扩展数据，增加「地区」层级：

```
产品     年份     地区     销售额
手机     2023     华东-江苏  1000
手机     2023     华东-浙江  2000
手机     2024     华南-广东  1500
电脑     2023     华东-江苏  3000
电脑     2024     华东-浙江  2500
电脑     2024     华南-广东  3500
```

### 代码

```javascript
function 多层字段分析() {
    var data = Application.ActiveWorkbook.Worksheets("Sheet1").Range("A1:D7").Value2;
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f3', '产品,地区'],      // 2层行字段
        ['f2', '年份'],               // 1层列字段
        ['sum("f4"),count()', '销售额,订单数'],
        1, 1, '@^@',
        {
            cornerTitle: '多层分析',
            layoutMode: 'outline',      // 大纲模式
            rowFieldIndent: true,       // 启用缩进
            rowFieldIndentSize: 4,      // 4空格缩进
            rowSubtotals: { enabled: true, label: '小计' },
            grandTotals: { row: true, column: true }
        }
    );
    
    var newWs = Application.ActiveWorkbook.Worksheets.Add();
    newWs.Name = "多层分析";
    result.toRange("A1", true);
}

多层字段分析();
```

### 输出效果

```
多层分析          2023        2024
产品      地区        
手机      华东-江苏   1000        
          华东-浙江   2000        
          华南-广东               1500
          小计        3000        1500
电脑      ...
```

---

## 教程4：百分比分析（5分钟）

### 目标

将数值显示为百分比形式

### 三种百分比模式

```javascript
function 百分比分析() {
    var data = Application.ActiveWorkbook.Worksheets("Sheet1").Range("A1:D7").Value2;
    
    // 模式1：占总计百分比
    var result1 = Array2D.z超级透视(
        data, ['f1', '产品'], ['f2', '年份'], ['sum("f4")', '占比'],
        1, 1, '@^@',
        { displayAs: { mode: 'percentOfGrandTotal', decimals: 2 } }
    );
    
    // 模式2：占行总计百分比
    var result2 = Array2D.z超级透视(
        data, ['f1', '产品'], ['f2', '年份'], ['sum("f4")', '行占比'],
        1, 1, '@^@',
        { displayAs: { mode: 'percentOfRowTotal', decimals: 1 } }
    );
    
    // 模式3：占列总计百分比
    var result3 = Array2D.z超级透视(
        data, ['f1', '产品'], ['f2', '年份'], ['sum("f4")', '列占比'],
        1, 1, '@^@',
        { displayAs: { mode: 'percentOfColTotal', decimals: 1 } }
    );
    
    // 输出到不同工作表
    var ws1 = Application.ActiveWorkbook.Worksheets.Add(); ws1.Name = "占总计%";
    var ws2 = Application.ActiveWorkbook.Worksheets.Add(); ws2.Name = "占行%";
    var ws3 = Application.ActiveWorkbook.Worksheets.Add(); ws3.Name = "占列%";
    
    result1.toRange("占总计%!A1");
    result2.toRange("占行%!A1");
    result3.toRange("占列%!A1");
}

百分比分析();
```

### 结果对比

**占总计%**：
```
产品    2023        2024
手机    22.22%      11.11%
电脑    22.22%      44.44%
```

**占行%**：
```
产品    2023        2024
手机    66.67%      33.33%   ← 每行合计100%
电脑    33.33%      66.67%
```

**占列%**：
```
产品    2023        2024
手机    50.00%      20.00%   ← 每列合计100%
电脑    50.00%      80.00%
```

---

## 教程5：实战案例（10分钟）

### 场景：销售部门月度报表

#### 数据准备（Sheet1）

```
A        B        C        D        E
销售员   部门     月份     产品     销售额
张三     华东     1月      手机     5000
张三     华东     1月      电脑     8000
张三     华东     2月      手机     6000
李四     华南     1月      电脑     9000
李四     华南     2月      手机     7000
王五     华东     1月      手机     5500
王五     华东     2月      电脑     8500
```

#### 需求分析

1. 按「部门-销售员」行分析
2. 按「月份」列分析
3. 统计「销售额」和「订单数」
4. 显示小计和总计
5. 层级缩进显示

#### 完整代码

```javascript
function 销售月度报表() {
    Console.log("开始生成销售月度报表...");
    
    try {
        // 1. 获取数据
        var ws = Application.ActiveWorkbook.Worksheets("Sheet1");
        var data = ws.Range("A1:E8").Value2;
        Console.log("✓ 数据读取完成，共 " + (data.length - 1) + " 条记录");
        
        // 2. 创建透视表
        Console.log("正在生成透视表...");
        var result = Array2D.z超级透视(
            data,
            ['f2,f1', '部门,销售员'],      // 行：部门→销售员
            ['f3', '月份'],                 // 列：月份
            ['sum("f5"),count()', '销售额,订单数'],
            1, 1, '@^@',
            {
                cornerTitle: '2024年销售月度报表',
                layoutMode: 'outline',
                rowFieldIndent: true,
                rowFieldIndentSize: 4,
                rowSubtotals: { enabled: true, label: '部门合计' },
                colSubtotals: { enabled: true, label: '月度合计' },
                grandTotals: { row: true, column: true, label: '总计' }
            }
        );
        
        // 3. 输出结果
        var newWs;
        try {
            newWs = Application.ActiveWorkbook.Worksheets("销售报表");
            newWs.Cells.Clear();
        } catch (e) {
            newWs = Application.ActiveWorkbook.Worksheets.Add();
            newWs.Name = "销售报表";
        }
        
        result.toRange("销售报表!A1", true);
        Console.log("✓ 透视表已输出到【销售报表】工作表");
        
        // 4. 格式化
        newWs.Range("A1").Font.Bold = true;
        newWs.Range("A1").Font.Size = 14;
        
        // 5. 获取统计信息
        var meta = result.getMeta();
        Console.log("\n=== 报表统计 ===");
        Console.log("部门数: " + meta.rowCount);
        Console.log("月份数: " + meta.colCount);
        Console.log("总销售额: " + meta.grandTotal);
        
        Console.log("\n✅ 报表生成完成！");
        
    } catch (e) {
        Console.log("❌ 错误: " + e.message);
        Console.log("堆栈: " + e.stack);
    }
}

// 执行
销售月度报表();
```

#### 预期输出

```
2024年销售月度报表        1月         2月         月度合计
部门      销售员          
华东      张三            13000       6000        19000
          王五            5500        8500        14000
          部门合计        18500       14500       33000
华南      李四            9000        7000        16000
          部门合计        9000        7000        16000
总计                    27500       21500       49000
```

---

## 教程6：调试技巧（3分钟）

### 如何排查问题

```javascript
function 调试示例() {
    var data = 获取数据("A1:D10");
    
    // 技巧1：查看原始数据
    Console.log("数据前3行:");
    for (var i = 0; i < Math.min(3, data.length); i++) {
        Console.log(i + ": " + JSON.stringify(data[i]));
    }
    
    // 技巧2：逐步执行
    var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);
    
    Console.log("结果行数: " + result.length);
    Console.log("表头行: " + JSON.stringify(result[0]));
    Console.log("第一行数据: " + JSON.stringify(result[1]));
    
    // 技巧3：查看元数据
    var meta = result.getMeta();
    Console.log("元数据: " + JSON.stringify(meta, null, 2));
    
    // 技巧4：只输出部分数据测试
    result.toRange("测试!A1");
}
```

---

## 附录：完整参数参考

```javascript
Array2D.z超级透视(
    arr,                        // 数据源（二维数组）
    rowFields,                  // 行字段：'f1,f2' 或 ['f1,f2', '标题']
    colFields,                  // 列字段
    dataFields,                 // 数据字段：'sum("f3")' 或 ['sum("f3")', '标题']
    headerRows = 1,             // 表头行数
    outputHeader = 1,           // 1=输出, 0=不输出, -1=隐藏行标题
    separator = '@^@',          // 分隔符
    {                           // options (v3.9.0)
        cornerTitle: '',
        layoutMode: 'outline',
        rowFieldIndent: true,
        rowFieldIndentSize: 4,
        rowSubtotals: { enabled: false },
        colSubtotals: { enabled: false },
        grandTotals: { row: false, column: false },
        displayAs: { mode: 'value', decimals: 2 }
    }
);
```

---

**文档版本**: v3.9.0  
**适用环境**: WPS Office JSA
