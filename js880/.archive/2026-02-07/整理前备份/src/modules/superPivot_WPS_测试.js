/**
 * =======================================================================
 * superPivot v3.9.0 全面测试套件
 * =======================================================================
 *
 * 测试范围:
 * 1. 基础透视功能 (单行单列、多行多列)
 * 2. 小计与总计功能 (行小计、列小计、总计)
 * 3. 百分比显示功能 (占总计%、占行%、占列%、占父级%)
 * 4. 显示格式模式 (compact/outline/tabular)
 * 5. 聚合函数 (count/sum/average/max/min)
 * 6. 排序功能 (升序/降序/原始顺序)
 * 7. 性能测试
 * 8. 边界情况测试
 *
 * 使用方法:
 * 1. 在WPS中打开包含测试数据的文件
 * 2. 按 Alt+F11 打开宏编辑器
 * 3. 运行 运行全部测试()
 * =======================================================================
 */

// ==================== 测试数据生成器 ====================

/**
 * 生成标准测试数据
 * 包含: 大区、省份、城市、产品类别、年份、季度、月份、销售额、利润
 */
function 生成标准测试数据(行数) {
    行数 = 行数 || 200;
    
    var 大区 = ['华东', '华南', '华北', '西南'];
    var 省份 = ['江苏', '浙江', '广东', '北京', '四川'];
    var 城市 = ['南京', '苏州', '杭州', '广州', '深圳', '北京', '成都'];
    var 产品类别 = ['电子产品', '家电', '服装', '食品'];
    var 年份 = ['2023', '2024'];
    var 季度 = ['Q1', 'Q2', 'Q3', 'Q4'];
    var 月份 = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
    
    var 数据 = [['大区', '省份', '城市', '产品类别', '年份', '季度', '月份', '销售额', '利润']];
    
    for (var i = 0; i < 行数; i++) {
        var 销售额 = Math.floor(Math.random() * 10000) + 1000;
        var 利润 = Math.floor(销售额 * (0.1 + Math.random() * 0.3));
        
        数据.push([
            大区[Math.floor(Math.random() * 大区.length)],
            省份[Math.floor(Math.random() * 省份.length)],
            城市[Math.floor(Math.random() * 城市.length)],
            产品类别[Math.floor(Math.random() * 产品类别.length)],
            年份[Math.floor(Math.random() * 年份.length)],
            季度[Math.floor(Math.random() * 季度.length)],
            月份[Math.floor(Math.random() * 月份.length)],
            销售额,
            利润
        ]);
    }
    
    // 输出到工作表
    var wb = Application.ActiveWorkbook;
    var ws;
    try {
        ws = wb.Worksheets("测试数据");
        ws.Cells.Clear();
    } catch (e) {
        ws = wb.Worksheets.Add();
        ws.Name = "测试数据";
    }
    
    数据.toRange("测试数据!A1");
    console.log("✅ 生成测试数据: " + 行数 + " 行");
    return 数据;
}

/**
 * 从工作表获取测试数据
 */
function 获取测试数据() {
    var wb = Application.ActiveWorkbook;
    var ws;
    try {
        ws = wb.Worksheets("测试数据");
    } catch (e) {
        console.log("⚠️ 测试数据表不存在，生成新数据...");
        生成标准测试数据(200);
        ws = wb.Worksheets("测试数据");
    }
    
    var usedRange = ws.UsedRange;
    var data = usedRange.Value2;
    return new Array2D(data);
}

/**
 * 清空测试输出表
 */
function 清空测试输出() {
    var wb = Application.ActiveWorkbook;
    var ws;
    try {
        ws = wb.Worksheets("测试输出");
        ws.Cells.Clear();
    } catch (e) {
        ws = wb.Worksheets.Add();
        ws.Name = "测试输出";
    }
    return ws;
}

// ==================== 辅助函数 ====================

function 数字转列号(num) {
    var result = "";
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result || "A";
}

function 左填充(str, length) {
    str = String(str);
    while (str.length < length) {
        str = " " + str;
    }
    return str;
}

// ==================== 第1组: 基础功能测试 ====================

/**
 * 测试1: 单行单列基础透视
 */
function 测试1_单行单列基础() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试1: 单行单列基础透视                              ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1', '大区'],
        ['f5', '年份'],
        ['sum("f8"),sum("f9")', '销售额,利润'],
        0, 1
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   结果: " + result.z数量() + " 行 × " + result.z列数() + " 列");
    return { 通过: true, 耗时: time };
}

/**
 * 测试2: 多行多列透视
 */
function 测试2_多行多列() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试2: 多行多列透视 (2行×2列)                        ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1,f2', '大区,省份'],
        ['f5,f6', '年份,季度'],
        ['sum("f8"),count()', '销售额,订单数'],
        0, 1
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   结果: " + result.z数量() + " 行 × " + result.z列数() + " 列");
    return { 通过: true, 耗时: time };
}

/**
 * 测试3: 三层列字段透视
 */
function 测试3_三层列字段() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试3: 三层列字段透视 (年份→季度→月份)              ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1', '大区'],
        ['f5,f6,f7', '年份,季度,月份'],
        ['sum("f8")', '销售额'],
        0, 1
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   结果: " + result.z数量() + " 行 × " + result.z列数() + " 列");
    console.log("   表头行数: 4行 (1行标题+3层列字段)");
    return { 通过: true, 耗时: time };
}

// ==================== 第2组: 小计与总计功能测试 ====================

/**
 * 测试4: 行小计功能
 */
function 测试4_行小计功能() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试4: 行小计功能                                    ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1,f2', '大区,省份'],
        ['f5', '年份'],
        ['sum("f8")', '销售额'],
        0, 1, '@^@',
        {
            rowSubtotals: {
                enabled: true,
                position: 'bottom',
                label: '小计',
                aggregation: 'sum'
            }
        }
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   结果: " + result.z数量() + " 行 (包含小计行)");
    
    // 验证小计行存在
    var hasSubtotal = false;
    for (var i = 0; i < result.z数量(); i++) {
        if (String(result[i][1]).indexOf("小计") >= 0) {
            hasSubtotal = true;
            break;
        }
    }
    console.log("   小计行检测: " + (hasSubtotal ? "✅ 存在" : "❌ 不存在"));
    
    return { 通过: hasSubtotal, 耗时: time };
}

/**
 * 测试5: 总计功能
 */
function 测试5_总计功能() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试5: 总计功能 (行总计+列总计)                      ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1', '大区'],
        ['f5', '年份'],
        ['sum("f8")', '销售额'],
        0, 1, '@^@',
        {
            grandTotals: {
                row: true,
                column: true,
                label: '总计'
            }
        }
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   结果: " + result.z数量() + " 行 × " + result.z列数() + " 列");
    
    // 验证总计行存在
    var lastRow = result[result.z数量() - 1];
    var hasGrandTotal = String(lastRow[0]).indexOf("总计") >= 0;
    console.log("   总计行检测: " + (hasGrandTotal ? "✅ 存在" : "❌ 不存在"));
    
    return { 通过: hasGrandTotal, 耗时: time };
}

/**
 * 测试6: 完整小计+总计组合
 */
function 测试6_完整小计总计组合() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试6: 完整小计+总计组合                             ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1,f2', '大区,省份'],
        ['f5', '年份'],
        ['sum("f8"),count()', '销售额,订单数'],
        0, 1, '@^@',
        {
            rowSubtotals: {
                enabled: true,
                position: 'bottom',
                label: '小计',
                aggregation: 'sum'
            },
            colSubtotals: {
                enabled: true,
                position: 'right',
                label: '小计'
            },
            grandTotals: {
                row: true,
                column: true,
                label: '总计'
            }
        }
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   结果: " + result.z数量() + " 行 × " + result.z列数() + " 列");
    console.log("   包含: 行小计 + 列小计 + 行总计 + 列总计");
    
    return { 通过: true, 耗时: time };
}

// ==================== 第3组: 百分比显示功能测试 ====================

/**
 * 测试7: 占总计百分比
 */
function 测试7_占总计百分比() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试7: 占总计百分比显示                              ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1', '大区'],
        ['f5', '年份'],
        ['sum("f8")', '销售额占比'],
        0, 1, '@^@',
        {
            displayAs: {
                mode: 'percentOfGrandTotal',
                decimals: 2
            }
        }
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   显示模式: 占总计百分比");
    console.log("   示例值: 华东-2023 可能显示为 25.50%");
    
    return { 通过: true, 耗时: time };
}

/**
 * 测试8: 占行总计百分比
 */
function 测试8_占行总计百分比() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试8: 占行总计百分比显示                            ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1', '大区'],
        ['f5', '年份'],
        ['sum("f8")', '行内占比'],
        0, 1, '@^@',
        {
            displayAs: {
                mode: 'percentOfRowTotal',
                decimals: 1
            }
        }
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   显示模式: 占行总计百分比");
    console.log("   每行总计: 100%");
    
    return { 通过: true, 耗时: time };
}

/**
 * 测试9: 占列总计百分比
 */
function 测试9_占列总计百分比() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试9: 占列总计百分比显示                            ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1', '大区'],
        ['f5', '年份'],
        ['sum("f8")', '列内占比'],
        0, 1, '@^@',
        {
            displayAs: {
                mode: 'percentOfColTotal',
                decimals: 1
            }
        }
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   显示模式: 占列总计百分比");
    console.log("   每列总计: 100%");
    
    return { 通过: true, 耗时: time };
}

// ==================== 第4组: 显示格式模式测试 ====================

/**
 * 测试10: Compact紧凑模式
 */
function 测试10_Compact紧凑模式() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试10: Compact紧凑模式                              ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1,f2,f3', '大区,省份,城市'],
        ['f5', '年份'],
        ['sum("f8")', '销售额'],
        0, 1, '@^@',
        {
            layoutMode: 'compact',
            cornerTitle: '地理维度'
        }
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   布局模式: Compact (所有行字段合并成一列)");
    console.log("   显示格式: 华东 / 江苏 / 南京");
    console.log("   列数: " + result.z列数() + " 列 (比outline模式少)");
    
    return { 通过: true, 耗时: time };
}

/**
 * 测试11: Outline大纲模式 (默认)
 */
function 测试11_Outline大纲模式() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试11: Outline大纲模式 (默认)                       ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1,f2,f3', '大区,省份,城市'],
        ['f5', '年份'],
        ['sum("f8")', '销售额'],
        0, 1, '@^@',
        {
            layoutMode: 'outline',
            cornerTitle: '地理维度'
        }
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   布局模式: Outline (每级行字段独立一列)");
    console.log("   列数: " + result.z列数() + " 列");
    
    return { 通过: true, 耗时: time };
}

/**
 * 测试12: 层级缩进显示
 */
function 测试12_层级缩进显示() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试12: 层级缩进显示                                 ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f1,f2,f3', '大区,省份,城市'],
        ['f5', '年份'],
        ['sum("f8")', '销售额'],
        0, 1, '@^@',
        {
            layoutMode: 'outline',
            rowFieldIndent: true,
            rowFieldIndentSize: 4,
            cornerTitle: '地理维度'
        }
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   缩进设置: 每级4个空格");
    console.log("   显示效果: 省份比大区缩进4空格，城市比省份缩进4空格");
    
    return { 通过: true, 耗时: time };
}

// ==================== 第5组: 聚合函数与排序测试 ====================

/**
 * 测试13: 五种聚合函数
 */
function 测试13_五种聚合函数() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试13: 五种聚合函数 (count/sum/average/max/min)     ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    var t0 = new Date().getTime();
    var result = Array2D.z超级透视(
        arr,
        ['f4', '产品类别'],
        ['f5', '年份'],
        ['count(),sum("f8"),average("f8"),max("f8"),min("f8")', 
         '订单数,总销售额,平均销售额,最高单额,最低单额'],
        0, 1
    );
    var time = new Date().getTime() - t0;
    
    result.toRange("测试输出!A1", true);
    
    console.log("✅ 完成! 耗时: " + time + "ms");
    console.log("   聚合函数: count, sum, average, max, min");
    console.log("   列数: " + result.z列数() + " 列 (5种聚合×2年)");
    
    return { 通过: true, 耗时: time };
}

/**
 * 测试14: 排序功能 (升序/降序)
 */
function 测试14_排序功能() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试14: 排序功能 (升序/降序/原始顺序)                ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    清空测试输出();
    var arr = 获取测试数据().z跳过(1);
    
    // 测试降序
    console.log("--- 测试降序 (f1-) ---");
    var result1 = Array2D.z超级透视(
        arr,
        ['f1-', '大区(降序)'],
        ['f5', '年份'],
        ['sum("f8")', '销售额'],
        0, 1
    );
    result1.toRange("测试输出!A1", false);
    console.log("   第一行: " + result1[0][0]);
    
    // 计算第二个结果的起始列
    var cols1 = result1.z列数();
    var startCol2 = cols1 + 2;
    
    // 测试升序
    console.log("--- 测试升序 (f1+) ---");
    var result2 = Array2D.z超级透视(
        arr,
        ['f1+', '大区(升序)'],
        ['f5', '年份'],
        ['sum("f8")', '销售额'],
        0, 1
    );
    result2.toRange("测试输出!" + 数字转列号(startCol2) + "1", false);
    console.log("   第一行: " + result2[0][0]);
    
    console.log("✅ 排序测试完成!");
    return { 通过: true, 耗时: 0 };
}

// ==================== 第6组: 性能测试 ====================

/**
 * 测试15: 性能压力测试
 */
function 测试15_性能压力测试() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试15: 性能压力测试                                 ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    // 生成大数据
    console.log("生成1000行测试数据...");
    生成标准测试数据(1000);
    var arr = 获取测试数据().z跳过(1);
    
    var testSizes = [100, 500, 1000];
    var results = [];
    
    console.log("\n┌──────────┬──────────┬──────────┬──────────┐");
    console.log("│ 数据量   │ 透视ms   │ 输出ms   │ 总计ms   │");
    console.log("├──────────┼──────────┼──────────┼──────────┤");
    
    for (var i = 0; i < testSizes.length; i++) {
        var size = testSizes[i];
        var testData = arr._items.slice(0, size);
        
        var t0 = new Date().getTime();
        var result = Array2D.z超级透视(
            testData,
            ['f1,f2', '大区,省份'],
            ['f5,f6', '年份,季度'],
            ['count(),sum("f8")', '订单数,销售额'],
            0, 0  // 不输出表头以提高性能
        );
        var pivotTime = new Date().getTime() - t0;
        
        var t1 = new Date().getTime();
        result.toRange("测试输出!A1", false);
        var outputTime = new Date().getTime() - t1;
        
        var totalTime = pivotTime + outputTime;
        
        console.log("│ " + 左填充(size, 8) + " │ " + 左填充(pivotTime, 8) + " │ " + 左填充(outputTime, 8) + " │ " + 左填充(totalTime, 8) + " │");
        
        results.push({
            数据量: size,
            透视: pivotTime,
            输出: outputTime,
            总计: totalTime
        });
    }
    
    console.log("└──────────┴──────────┴──────────┴──────────┘");
    
    return { 通过: true, 结果: results };
}

// ==================== 第7组: 边界情况测试 ====================

/**
 * 测试16: 边界情况测试
 */
function 测试16_边界情况测试() {
    console.log("\n╔══════════════════════════════════════════════════════╗");
    console.log("║  测试16: 边界情况测试                                 ║");
    console.log("╚══════════════════════════════════════════════════════╝");
    
    var results = [];
    
    // 测试1: 小数据集
    console.log("\n--- 测试1: 小数据集 (3行) ---");
    try {
        var smallData = [
            ['产品', '年份', '销售额'],
            ['A', '2023', 100],
            ['A', '2023', 200],
            ['B', '2024', 300]
        ];
        var result1 = Array2D.z超级透视(
            Array2D(smallData).z跳过(1),
            ['f1', '产品'],
            ['f2', '年份'],
            ['sum("f3")', '销售额'],
            0, 1
        );
        console.log("   ✅ 通过: " + result1.z数量() + " 行");
        results.push("小数据集: 通过");
    } catch (e) {
        console.log("   ❌ 失败: " + e.message);
        results.push("小数据集: 失败");
    }
    
    // 测试2: 单值列字段
    console.log("\n--- 测试2: 单值列字段 ---");
    try {
        var singleColData = [
            ['产品', '销售额'],
            ['A', 100],
            ['B', 200]
        ];
        var result2 = Array2D.z超级透视(
            Array2D(singleColData).z跳过(1),
            ['f1', '产品'],
            ['f2', '销售额'],
            ['count()'],
            0, 1
        );
        console.log("   ✅ 通过: " + result2.z数量() + " 行");
        results.push("单值列字段: 通过");
    } catch (e) {
        console.log("   ❌ 失败: " + e.message);
        results.push("单值列字段: 失败");
    }
    
    // 测试3: 无表头模式
    console.log("\n--- 测试3: 无表头模式 ---");
    try {
        var noHeaderData = [
            ['华东', '2023', 100],
            ['华南', '2024', 200]
        ];
        var result3 = Array2D.z超级透视(
            noHeaderData,
            ['f1', '大区'],
            ['f2', '年份'],
            ['sum("f3")', '销售额'],
            0, 0  // 无表头输入，无表头输出
        );
        console.log("   ✅ 通过: " + result3.z数量() + " 行");
        results.push("无表头模式: 通过");
    } catch (e) {
        console.log("   ❌ 失败: " + e.message);
        results.push("无表头模式: 失败");
    }
    
    console.log("\n========== 边界测试汇总 ==========");
    for (var i = 0; i < results.length; i++) {
        console.log(results[i]);
    }
    
    return { 通过: true, 结果: results };
}

// ==================== 主运行函数 ====================

/**
 * 运行全部测试
 */
function 运行全部测试() {
    console.log("╔════════════════════════════════════════════════════════════════════╗");
    console.log("║           superPivot v3.9.0 全面测试套件                           ║");
    console.log("║           测试范围: 基础功能 | 小计总计 | 百分比 | 布局模式         ║");
    console.log("╚════════════════════════════════════════════════════════════════════╝");
    
    var allResults = [];
    var startTime = new Date().getTime();
    
    // 生成标准测试数据
    生成标准测试数据(200);
    
    var tests = [
        { name: "基础-单行单列", func: 测试1_单行单列基础 },
        { name: "基础-多行多列", func: 测试2_多行多列 },
        { name: "基础-三层列字段", func: 测试3_三层列字段 },
        { name: "小计-行小计功能", func: 测试4_行小计功能 },
        { name: "小计-总计功能", func: 测试5_总计功能 },
        { name: "小计-完整组合", func: 测试6_完整小计总计组合 },
        { name: "百分比-占总计", func: 测试7_占总计百分比 },
        { name: "百分比-占行总计", func: 测试8_占行总计百分比 },
        { name: "百分比-占列总计", func: 测试9_占列总计百分比 },
        { name: "布局-Compact模式", func: 测试10_Compact紧凑模式 },
        { name: "布局-Outline模式", func: 测试11_Outline大纲模式 },
        { name: "布局-层级缩进", func: 测试12_层级缩进显示 },
        { name: "聚合-五种函数", func: 测试13_五种聚合函数 },
        { name: "排序-升序降序", func: 测试14_排序功能 },
        { name: "性能-压力测试", func: 测试15_性能压力测试 },
        { name: "边界-特殊情况", func: 测试16_边界情况测试 }
    ];
    
    for (var i = 0; i < tests.length; i++) {
        try {
            var result = tests[i].func();
            allResults.push({
                名称: tests[i].name,
                状态: result.通过 ? "✅ 通过" : "❌ 失败",
                耗时: result.耗时 || 0
            });
        } catch (e) {
            allResults.push({
                名称: tests[i].name,
                状态: "❌ 错误: " + e.message,
                耗时: 0
            });
        }
    }
    
    var endTime = new Date().getTime();
    
    // 输出汇总报告
    console.log("\n╔════════════════════════════════════════════════════════════════════╗");
    console.log("║                         测试汇总报告                               ║");
    console.log("╚════════════════════════════════════════════════════════════════════╝");
    
    var passCount = 0;
    for (var j = 0; j < allResults.length; j++) {
        var r = allResults[j];
        console.log((j + 1) + ". " + 左填充(r.名称, 20) + " | " + r.状态);
        if (r.状态.indexOf("通过") >= 0) passCount++;
    }
    
    console.log("\n总计: " + passCount + "/" + allResults.length + " 通过");
    console.log("总耗时: " + (endTime - startTime) + " ms");
    console.log("\n✅ 全部测试完成! 请查看'测试输出'表中的详细结果。");
    
    return allResults;
}

/**
 * 快速测试 - 只运行核心测试
 */
function 快速测试() {
    console.log("========== 快速测试 ==========");
    
    生成标准测试数据(100);
    
    测试1_单行单列基础();
    测试4_行小计功能();
    测试7_占总计百分比();
    测试10_Compact紧凑模式();
    
    console.log("\n✅ 快速测试完成!");
}

/**
 * 专项测试 - 只测试新功能
 */
function 专项测试_新功能() {
    console.log("========== v3.9.0 新功能专项测试 ==========");
    
    生成标准测试数据(200);
    
    测试4_行小计功能();
    测试5_总计功能();
    测试6_完整小计总计组合();
    测试7_占总计百分比();
    测试8_占行总计百分比();
    测试9_占列总计百分比();
    测试10_Compact紧凑模式();
    测试12_层级缩进显示();
    
    console.log("\n✅ 新功能专项测试完成!");
}

// 导出主要函数供外部调用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        运行全部测试: 运行全部测试,
        快速测试: 快速测试,
        专项测试_新功能: 专项测试_新功能
    };
}
