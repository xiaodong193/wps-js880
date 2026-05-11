/**
 * =======================================================================
 * superPivot v3.9.0 示例测试
 * =======================================================================
 * 
 * 适合：初次使用，想快速了解功能
 * 
 * 使用方法：
 *   1. 在 WPS 中新建一个工作簿
 *   2. 在 Sheet1 中输入测试数据
 *   3. 按 Alt+F11 打开宏编辑器
 *   4. 粘贴此代码并运行
 * =======================================================================
 */

// ==================== 示例1: 最基础的透视表 ====================

function 示例1_基础透视() {
    Console.log("=== 示例1: 最基础的透视表 ===");
    
    // 准备数据
    var data = [
        ['产品', '年份', '销售额'],
        ['手机', '2023', 1000],
        ['手机', '2024', 2000],
        ['电脑', '2023', 3000],
        ['电脑', '2024', 4000]
    ];
    
    // 创建透视表
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品'],          // 行：产品
        ['f2', '年份'],          // 列：年份
        ['sum("f3")', '销售额']  // 数据：销售额求和
    );
    
    // 输出到工作表
    var ws = Application.ActiveWorkbook.Worksheets.Add();
    ws.Name = "示例1-基础";
    result.toRange("A1");
    
    Console.log("✓ 已在【示例1-基础】工作表中生成透视表");
    Console.log("  行字段: 产品");
    Console.log("  列字段: 年份");
    Console.log("  数据: 销售额");
}

// ==================== 示例2: 带小计和总计 ====================

function 示例2_小计总计() {
    Console.log("=== 示例2: 带小计和总计的透视表 ===");
    
    var data = [
        ['类别', '产品', '年份', '销售额'],
        ['电子', '手机', '2023', 1000],
        ['电子', '手机', '2024', 2000],
        ['电子', '电脑', '2023', 3000],
        ['电子', '电脑', '2024', 4000],
        ['家电', '电视', '2023', 2500],
        ['家电', '电视', '2024', 3500]
    ];
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f2', '类别,产品'],
        ['f3', '年份'],
        ['sum("f4")', '销售额'],
        1, 1, '@^@',
        {
            cornerTitle: '销售分析表',
            rowSubtotals: { enabled: true, label: '类别小计' },
            colSubtotals: { enabled: true, label: '年度小计' },
            grandTotals: { row: true, column: true, label: '总计' }
        }
    );
    
    var ws = Application.ActiveWorkbook.Worksheets.Add();
    ws.Name = "示例2-小计总计";
    result.toRange("A1", true);  // true = 应用合并
    
    // 查看元数据
    var meta = result.getMeta();
    Console.log("✓ 已在【示例2-小计总计】工作表中生成透视表");
    Console.log("  总行数（含表头）: " + result.length);
    Console.log("  角标题: " + meta.options.cornerTitle);
    Console.log("  总计值: " + meta.grandTotal);
}

// ==================== 示例3: 百分比显示 ====================

function 示例3_百分比显示() {
    Console.log("=== 示例3: 百分比显示模式 ===");
    
    var data = [
        ['产品', '年份', '销售额'],
        ['A', '2023', 100],
        ['A', '2024', 200],
        ['B', '2023', 300],
        ['B', '2024', 400]
    ];
    
    // 占总计百分比
    var result1 = Array2D.z超级透视(
        data,
        ['f1', '产品'],
        ['f2', '年份'],
        ['sum("f3")', '占总计%'],
        1, 1, '@^@',
        {
            displayAs: { mode: 'percentOfGrandTotal', decimals: 2 }
        }
    );
    
    var ws1 = Application.ActiveWorkbook.Worksheets.Add();
    ws1.Name = "示例3a-占总计";
    result1.toRange("A1");
    
    // 占行百分比
    var result2 = Array2D.z超级透视(
        data,
        ['f1', '产品'],
        ['f2', '年份'],
        ['sum("f3")', '占行%'],
        1, 1, '@^@',
        {
            displayAs: { mode: 'percentOfRowTotal', decimals: 1 }
        }
    );
    
    var ws2 = Application.ActiveWorkbook.Worksheets.Add();
    ws2.Name = "示例3b-占行";
    result2.toRange("A1");
    
    Console.log("✓ 已在以下工作表中生成百分比透视表:");
    Console.log("  【示例3a-占总计】- 每个值占总销售额的百分比");
    Console.log("  【示例3b-占行】- 每行内各列的百分比");
}

// ==================== 示例4: 多层行列字段 ====================

function 示例4_多层字段() {
    Console.log("=== 示例4: 多层行列字段 ===");
    
    var data = [
        ['大类', '小类', '年份', '季度', '销售额'],
        ['华东', '上海', '2023', 'Q1', 1000],
        ['华东', '上海', '2023', 'Q2', 1200],
        ['华东', '杭州', '2023', 'Q1', 800],
        ['华东', '杭州', '2024', 'Q1', 900],
        ['华南', '广州', '2023', 'Q1', 1500],
        ['华南', '广州', '2024', 'Q2', 1600]
    ];
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f2', '大区,城市'],
        ['f3,f4', '年份,季度'],
        ['sum("f5")', '销售额'],
        1, 1, '@^@',
        {
            cornerTitle: '区域销售分析',
            layoutMode: 'outline',
            rowFieldIndent: true,
            rowFieldIndentSize: 4,
            grandTotals: { row: true, column: true }
        }
    );
    
    var ws = Application.ActiveWorkbook.Worksheets.Add();
    ws.Name = "示例4-多层字段";
    result.toRange("A1", true);
    
    Console.log("✓ 已在【示例4-多层字段】工作表中生成透视表");
    Console.log("  行字段: 大区 → 城市（带层级缩进）");
    Console.log("  列字段: 年份 → 季度");
    Console.log("  布局模式: outline");
}

// ==================== 示例5: 多种聚合函数 ====================

function 示例5_多种聚合() {
    Console.log("=== 示例5: 多种聚合函数 ===");
    
    var data = [
        ['产品', '年份', '销售额', '数量'],
        ['A', '2023', 1000, 10],
        ['A', '2023', 2000, 20],
        ['A', '2024', 1500, 15],
        ['B', '2023', 3000, 30],
        ['B', '2024', 2500, 25]
    ];
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品'],
        ['f2', '年份'],
        ['count(),sum("f3"),average("f3"),max("f3"),min("f3")', 
         '订单数,总销售额,平均单价,最高单价,最低单价']
    );
    
    var ws = Application.ActiveWorkbook.Worksheets.Add();
    ws.Name = "示例5-多种聚合";
    result.toRange("A1");
    
    Console.log("✓ 已在【示例5-多种聚合】工作表中生成透视表");
    Console.log("  使用了5种聚合函数:");
    Console.log("    - count(): 订单数");
    Console.log("    - sum(): 总销售额");
    Console.log("    - average(): 平均单价");
    Console.log("    - max(): 最高单价");
    Console.log("    - min(): 最低单价");
}

// ==================== 示例6: 读取工作表数据 ====================

function 示例6_读取工作表() {
    Console.log("=== 示例6: 从当前工作表读取数据 ===");
    
    try {
        // 尝试从 Sheet1 读取数据
        var ws = Application.ActiveWorkbook.Worksheets("Sheet1");
        var usedRange = ws.UsedRange;
        
        if (usedRange.Rows.Count < 2) {
            Console.log("⚠ Sheet1 中没有足够的数据，使用示例数据");
            示例1_基础透视();
            return;
        }
        
        var data = usedRange.Value2;
        Console.log("✓ 从 Sheet1 读取了 " + (data.length - 1) + " 行数据");
        
        // 根据列数自动判断字段
        var colCount = data[0].length;
        Console.log("  列数: " + colCount);
        
        // 使用第一列和第二列作为行列字段，最后一列作为数据
        var rowField = 'f1';
        var colField = colCount > 2 ? 'f2' : 'f1';
        var dataField = 'f' + colCount;
        
        var result = Array2D.z超级透视(
            data,
            [rowField, String(data[0][0])],
            [colField, String(data[0][1] || data[0][0])],
            ['sum("' + dataField + '")', String(data[0][colCount - 1])]
        );
        
        var newWs = Application.ActiveWorkbook.Worksheets.Add();
        newWs.Name = "示例6-工作表透视";
        result.toRange("A1");
        
        Console.log("✓ 已在【示例6-工作表透视】工作表中生成透视表");
        Console.log("  行字段: " + data[0][0]);
        Console.log("  列字段: " + (data[0][1] || data[0][0]));
        Console.log("  数据字段: " + data[0][colCount - 1]);
        
    } catch (e) {
        Console.log("❌ 读取失败: " + e.message);
        Console.log("  请确保 Sheet1 中有数据");
    }
}

// ==================== 示例7: 完整功能演示 ====================

function 示例7_完整演示() {
    Console.log("=== 示例7: 完整功能演示 ===");
    
    // 生成更丰富的测试数据
    var data = [
        ['销售大区', '销售城市', '产品类别', '产品名称', '年份', '季度', '销售额', '利润'],
        ['华东', '上海', '电子', '手机', '2023', 'Q1', 10000, 2000],
        ['华东', '上海', '电子', '手机', '2023', 'Q2', 12000, 2400],
        ['华东', '上海', '电子', '电脑', '2023', 'Q1', 25000, 5000],
        ['华东', '杭州', '电子', '手机', '2023', 'Q1', 8000, 1600],
        ['华东', '杭州', '家电', '电视', '2023', 'Q2', 15000, 3000],
        ['华东', '杭州', '家电', '电视', '2024', 'Q1', 16000, 3200],
        ['华南', '广州', '电子', '电脑', '2023', 'Q1', 30000, 6000],
        ['华南', '广州', '电子', '电脑', '2024', 'Q2', 32000, 6400],
        ['华南', '深圳', '家电', '空调', '2023', 'Q1', 20000, 4000],
        ['华南', '深圳', '家电', '空调', '2024', 'Q2', 22000, 4400]
    ];
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f2', '大区,城市'],
        ['f5,f6', '年份,季度'],
        ['sum("f7"),sum("f8")', '销售额,利润'],
        1, 1, '@^@',
        {
            cornerTitle: '2023-2024年度销售分析报表',
            layoutMode: 'outline',
            rowFieldIndent: true,
            rowFieldIndentSize: 4,
            rowSubtotals: { enabled: true, label: '小计' },
            colSubtotals: { enabled: true, label: '小计' },
            grandTotals: { row: true, column: true, label: '总计' }
        }
    );
    
    var ws = Application.ActiveWorkbook.Worksheets.Add();
    ws.Name = "示例7-完整演示";
    result.toRange("A1", true);
    
    // 格式化
    ws.Range("A1").Font.Bold = true;
    ws.Range("A1").Font.Size = 14;
    ws.Columns.AutoFit();
    
    // 获取并显示元数据
    var meta = result.getMeta();
    
    Console.log("✓ 已在【示例7-完整演示】工作表中生成完整透视表");
    Console.log("");
    Console.log("=== 报表元数据 ===");
    Console.log("版本: " + meta.version);
    Console.log("行字段: " + meta.rowTitles.join(' → '));
    Console.log("列字段: " + meta.colTitles.join(' → '));
    Console.log("数据字段: " + meta.dataTitles.join(', '));
    Console.log("数据行数: " + meta.rowCount);
    Console.log("表头行数: " + meta.headerRowCount);
    Console.log("");
    Console.log("配置选项:");
    Console.log("  角标题: " + meta.options.cornerTitle);
    Console.log("  布局模式: " + meta.options.layoutMode);
    Console.log("  层级缩进: " + (meta.options.rowFieldIndent ? '启用' : '禁用'));
    Console.log("  行小计: " + (meta.options.rowSubtotals.enabled ? '启用' : '禁用'));
    Console.log("  列小计: " + (meta.options.colSubtotals.enabled ? '启用' : '禁用'));
    Console.log("  总计: 行=" + meta.options.grandTotals.row + ", 列=" + meta.options.grandTotals.column);
}

// ==================== 主入口 ====================

function 运行所有示例() {
    Console.log("╔══════════════════════════════════════════════════════╗");
    Console.log("║     superPivot v3.9.0 示例测试                       ║");
    Console.log("║     将依次生成7个示例工作表                          ║");
    Console.log("╚══════════════════════════════════════════════════════╝");
    Console.log("");
    
    var examples = [
        示例1_基础透视,
        示例2_小计总计,
        示例3_百分比显示,
        示例4_多层字段,
        示例5_多种聚合,
        示例6_读取工作表,
        示例7_完整演示
    ];
    
    for (var i = 0; i < examples.length; i++) {
        try {
            examples[i]();
        } catch (e) {
            Console.log("❌ 示例" + (i + 1) + "失败: " + e.message);
        }
        Console.log("");
    }
    
    Console.log("╔══════════════════════════════════════════════════════╗");
    Console.log("║     所有示例已生成完成！                             ║");
    Console.log("╚══════════════════════════════════════════════════════╝");
    Console.log("");
    Console.log("生成的7个工作表:");
    Console.log("  1. 示例1-基础");
    Console.log("  2. 示例2-小计总计");
    Console.log("  3. 示例3a-占总计");
    Console.log("  4. 示例3b-占行");
    Console.log("  5. 示例4-多层字段");
    Console.log("  6. 示例5-多种聚合");
    Console.log("  7. 示例6-工作表透视（如果Sheet1有数据）");
    Console.log("  8. 示例7-完整演示");
    Console.log("");
    Console.log("建议: 逐个查看每个工作表，对比代码和结果");
}

// 快捷入口
function 快速示例() {
    示例1_基础透视();
    示例2_小计总计();
    Console.log("✓ 已生成2个基础示例，请查看工作表");
}

// 如果直接运行此文件，执行所有示例
if (typeof Application !== 'undefined') {
    Console.log("superPivot 示例测试已加载");
    Console.log("可用函数:");
    Console.log("  运行所有示例() - 运行全部7个示例");
    Console.log("  快速示例() - 运行前2个基础示例");
    Console.log("  示例1_基础透视() - 单个示例");
    Console.log("  ...");
}
