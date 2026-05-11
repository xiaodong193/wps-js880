/**
 * =======================================================================
 * superPivot v3.9.0 全面测试套件
 * =======================================================================
 * 
 * 版本: 1.0.0
 * 日期: 2026-02-06
 * 适用环境: WPS Office JavaScript API (JSA)
 * 
 * 测试范围:
 *   - 基础透视功能
 *   - 多层行列字段
 *   - 小计与总计功能
 *   - 百分比显示模式
 *   - 布局模式（compact/outline/tabular）
 *   - 层级缩进
 *   - 角标题
 *   - 元数据获取
 *   - 性能测试
 *   - 边界情况
 * 
 * 使用方法:
 *   1. 在 WPS 中打开测试文件
 *   2. 按 Alt+F11 打开宏编辑器
 *   3. 运行 运行全部测试()
 * =======================================================================
 */

// ==================== 测试配置 ====================

var TEST_CONFIG = {
    version: '3.9.0',
    outputSheet: '测试结果',
    dataSheet: '测试数据',
    verbose: true,           // 是否输出详细日志
    stopOnError: false       // 出错时是否停止
};

var TEST_RESULTS = [];
var TEST_START_TIME = 0;
var TEST_END_TIME = 0;

// ==================== 测试数据生成器 ====================

/**
 * 生成标准测试数据
 * @returns {Array} 测试数据二维数组
 */
function 生成测试数据() {
    var data = [];
    
    // 表头
    data.push(['产品类别', '产品名称', '年份', '季度', '地区', '销售员', '销售额', '数量']);
    
    // 测试数据（20行）
    var categories = ['电子产品', '家电', '服装'];
    var products = {
        '电子产品': ['手机', '电脑', '平板'],
        '家电': ['电视', '冰箱', '空调'],
        '服装': ['T恤', '裤子', '外套']
    };
    var years = ['2023', '2024'];
    var quarters = ['Q1', 'Q2', 'Q3', 'Q4'];
    var regions = ['华东', '华南', '华北'];
    var salesmen = ['张三', '李四', '王五'];
    
    for (var i = 0; i < 20; i++) {
        var cat = categories[Math.floor(Math.random() * categories.length)];
        var prod = products[cat][Math.floor(Math.random() * products[cat].length)];
        var year = years[Math.floor(Math.random() * years.length)];
        var quarter = quarters[Math.floor(Math.random() * quarters.length)];
        var region = regions[Math.floor(Math.random() * regions.length)];
        var salesman = salesmen[Math.floor(Math.random() * salesmen.length)];
        var amount = Math.floor(Math.random() * 10000) + 1000;
        var qty = Math.floor(Math.random() * 100) + 10;
        
        data.push([cat, prod, year, quarter, region, salesman, amount, qty]);
    }
    
    return data;
}

/**
 * 生成大数据集（用于性能测试）
 * @param {Number} rowCount - 行数
 * @returns {Array} 大数据集
 */
function 生成大数据(rowCount) {
    rowCount = rowCount || 1000;
    var data = [];
    
    data.push(['类别', '年份', '月份', '销售额']);
    
    var categories = ['A', 'B', 'C', 'D', 'E'];
    var years = ['2023', '2024'];
    var months = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
    
    for (var i = 0; i < rowCount; i++) {
        var cat = categories[Math.floor(Math.random() * categories.length)];
        var year = years[Math.floor(Math.random() * years.length)];
        var month = months[Math.floor(Math.random() * months.length)];
        var amount = Math.floor(Math.random() * 10000) + 1000;
        
        data.push([cat, year, month, amount]);
    }
    
    return data;
}

/**
 * 输出测试数据到工作表
 */
function 准备测试数据() {
    var wb = Application.ActiveWorkbook;
    var ws;
    
    try {
        ws = wb.Worksheets(TEST_CONFIG.dataSheet);
        ws.Cells.Clear();
    } catch (e) {
        ws = wb.Worksheets.Add();
        ws.Name = TEST_CONFIG.dataSheet;
    }
    
    var data = 生成测试数据();
    ws.Range("A1").Resize(data.length, data[0].length).Value2 = data;
    
    if (TEST_CONFIG.verbose) {
        Console.log("✓ 测试数据已生成: " + (data.length - 1) + " 行");
    }
    
    return data;
}

// ==================== 测试框架核心 ====================

/**
 * 记录测试结果
 */
function 记录结果(testName, passed, message, duration) {
    var result = {
        name: testName,
        passed: passed,
        message: message || '',
        duration: duration || 0,
        timestamp: new Date().toISOString()
    };
    
    TEST_RESULTS.push(result);
    
    if (TEST_CONFIG.verbose) {
        var status = passed ? '✅' : '❌';
        Console.log(status + ' ' + testName + 
            (duration ? ' (' + duration + 'ms)' : '') +
            (message ? ' - ' + message : ''));
    }
}

/**
 * 断言函数
 */
function 断言(condition, message) {
    if (!condition) {
        throw new Error(message || '断言失败');
    }
}

/**
 * 测试包装器
 */
function 运行测试(testName, testFunc) {
    var startTime = new Date().getTime();
    
    try {
        testFunc();
        var duration = new Date().getTime() - startTime;
        记录结果(testName, true, '', duration);
        return true;
    } catch (e) {
        var duration = new Date().getTime() - startTime;
        记录结果(testName, false, e.message, duration);
        
        if (TEST_CONFIG.stopOnError) {
            throw e;
        }
        return false;
    }
}

// ==================== 基础功能测试 ====================

/**
 * 测试1: 单行单列基础透视
 */
function 测试_基础透视() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品类别'],
        ['f3', '年份'],
        ['sum("f7")', '销售额']
    );
    
    断言(result && result.length > 0, '结果不能为空');
    断言(result.length > 1, '结果应包含表头和数据');
    
    // 验证元数据
    var meta = result.getMeta();
    断言(meta && meta.version === '3.9.0', '版本号应为3.9.0');
    断言(meta.rowFields.length === 1, '应只有1个行字段');
    断言(meta.colFields.length === 1, '应只有1个列字段');
}

/**
 * 测试2: 多行字段透视
 */
function 测试_多行字段() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f2', '产品类别,产品名称'],
        ['f3', '年份'],
        ['sum("f7")', '销售额']
    );
    
    var meta = result.getMeta();
    断言(meta.rowFields.length === 2, '应有2个行字段');
}

/**
 * 测试3: 多列字段透视
 */
function 测试_多列字段() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品类别'],
        ['f3,f4', '年份,季度'],
        ['sum("f7")', '销售额']
    );
    
    var meta = result.getMeta();
    断言(meta.colFields.length === 2, '应有2个列字段');
}

/**
 * 测试4: 多数据字段
 */
function 测试_多数据字段() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品类别'],
        ['f3', '年份'],
        ['count(),sum("f7"),average("f7")', '订单数,总销售额,平均单价']
    );
    
    var meta = result.getMeta();
    断言(meta.dataFields.length === 3, '应有3个数据字段');
}

// ==================== v3.9.0 新功能测试 ====================

/**
 * 测试5: 行小计功能
 */
function 测试_行小计() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品类别'],
        ['f3', '年份'],
        ['sum("f7")', '销售额'],
        1, 1, '@^@',
        {
            rowSubtotals: {
                enabled: true,
                label: '小计'
            }
        }
    );
    
    var meta = result.getMeta();
    断言(meta.options.rowSubtotals.enabled === true, '行小计应启用');
    
    // 验证结果中包含小计行
    var hasSubtotal = false;
    for (var i = 0; i < result.length; i++) {
        for (var j = 0; j < result[i].length; j++) {
            if (String(result[i][j]).indexOf('小计') >= 0) {
                hasSubtotal = true;
                break;
            }
        }
    }
    断言(hasSubtotal, '结果中应包含"小计"字样');
}

/**
 * 测试6: 列小计功能
 */
function 测试_列小计() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品类别'],
        ['f3', '年份'],
        ['sum("f7")', '销售额'],
        1, 1, '@^@',
        {
            colSubtotals: {
                enabled: true,
                label: '小计'
            }
        }
    );
    
    var meta = result.getMeta();
    断言(meta.options.colSubtotals.enabled === true, '列小计应启用');
}

/**
 * 测试7: 总计功能
 */
function 测试_总计() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品类别'],
        ['f3', '年份'],
        ['sum("f7")', '销售额'],
        1, 1, '@^@',
        {
            grandTotals: {
                row: true,
                column: true,
                label: '总计'
            }
        }
    );
    
    var meta = result.getMeta();
    断言(meta.options.grandTotals.row === true, '行总计应启用');
    断言(meta.options.grandTotals.column === true, '列总计应启用');
    断言(meta.grandTotal !== null, '应有总计值');
    
    // 验证最后一行包含"总计"
    var lastRow = result[result.length - 1];
    var hasGrandTotal = false;
    for (var i = 0; i < lastRow.length; i++) {
        if (String(lastRow[i]).indexOf('总计') >= 0) {
            hasGrandTotal = true;
            break;
        }
    }
    断言(hasGrandTotal, '最后一行应包含"总计"字样');
}

/**
 * 测试8: 百分比显示 - 占总计
 */
function 测试_百分比_占总计() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品类别'],
        ['f3', '年份'],
        ['sum("f7")', '占比'],
        1, 1, '@^@',
        {
            displayAs: {
                mode: 'percentOfGrandTotal',
                decimals: 2
            }
        }
    );
    
    // 验证结果包含百分号
    var hasPercent = false;
    for (var i = 1; i < result.length; i++) {
        for (var j = 0; j < result[i].length; j++) {
            if (String(result[i][j]).indexOf('%') >= 0) {
                hasPercent = true;
                break;
            }
        }
    }
    断言(hasPercent, '结果中应包含百分号');
}

/**
 * 测试9: 百分比显示 - 占行
 */
function 测试_百分比_占行() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品类别'],
        ['f3', '年份'],
        ['sum("f7")', '行占比'],
        1, 1, '@^@',
        {
            displayAs: {
                mode: 'percentOfRowTotal',
                decimals: 1
            }
        }
    );
    
    var meta = result.getMeta();
    断言(meta.options.displayAs.mode === 'percentOfRowTotal', '应为占行百分比模式');
}

/**
 * 测试10: 角标题
 */
function 测试_角标题() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品类别'],
        ['f3', '年份'],
        ['sum("f7")', '销售额'],
        1, 1, '@^@',
        {
            cornerTitle: '销售分析表'
        }
    );
    
    var meta = result.getMeta();
    断言(meta.options.cornerTitle === '销售分析表', '角标题应设置正确');
    
    // 验证第一行第一列包含角标题
    var firstCell = String(result[0][0]);
    断言(firstCell.indexOf('销售分析表') >= 0, '左上角应显示角标题');
}

/**
 * 测试11: 层级缩进
 */
function 测试_层级缩进() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f2', '类别,产品'],
        ['f3', '年份'],
        ['sum("f7")', '销售额'],
        1, 1, '@^@',
        {
            layoutMode: 'outline',
            rowFieldIndent: true,
            rowFieldIndentSize: 4
        }
    );
    
    var meta = result.getMeta();
    断言(meta.options.rowFieldIndent === true, '层级缩进应启用');
    断言(meta.options.rowFieldIndentSize === 4, '缩进空格数应为4');
}

/**
 * 测试12: 完整功能组合
 */
function 测试_完整功能组合() {
    var data = 生成测试数据();
    
    var result = Array2D.z超级透视(
        data,
        ['f1,f5', '产品类别,地区'],
        ['f3,f4', '年份,季度'],
        ['sum("f7"),count()', '销售额,订单数'],
        1, 1, '@^@',
        {
            cornerTitle: '综合销售分析',
            layoutMode: 'outline',
            rowFieldIndent: true,
            rowSubtotals: { enabled: true, label: '小计' },
            colSubtotals: { enabled: true, label: '小计' },
            grandTotals: { row: true, column: true, label: '总计' }
        }
    );
    
    var meta = result.getMeta();
    断言(meta.options.cornerTitle === '综合销售分析', '角标题应正确');
    断言(meta.options.rowSubtotals.enabled === true, '行小计应启用');
    断言(meta.options.colSubtotals.enabled === true, '列小计应启用');
    断言(meta.options.grandTotals.row === true, '行总计应启用');
    断言(meta.options.grandTotals.column === true, '列总计应启用');
}

// ==================== 边界情况测试 ====================

/**
 * 测试13: 空数据处理
 */
function 测试_空数据() {
    var data = [['产品', '销售额']];
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '产品'],
        ['f1', '产品'],
        ['count()']
    );
    
    断言(result && result.length > 0, '空数据应返回表头');
}

/**
 * 测试14: 单值列
 */
function 测试_单值列() {
    var data = [
        ['类别', '值'],
        ['A', 100],
        ['A', 200]
    ];
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '类别'],
        ['f1', '类别'],
        ['sum("f2")', '总和']
    );
    
    断言(result.length > 0, '单值列应正常处理');
}

/**
 * 测试15: 无表头模式
 */
function 测试_无表头模式() {
    var data = [
        ['A', '2023', 100],
        ['B', '2024', 200]
    ];
    
    var result = Array2D.z超级透视(
        data,
        ['f1', '类别'],
        ['f2', '年份'],
        ['sum("f3")', '值'],
        0,  // 无表头输入
        0   // 无表头输出
    );
    
    断言(result.length > 0, '无表头模式应正常处理');
}

// ==================== 性能测试 ====================

/**
 * 测试16: 大数据量性能
 */
function 测试_大数据性能() {
    var testSizes = [100, 500, 1000];
    var results = [];
    
    for (var i = 0; i < testSizes.length; i++) {
        var size = testSizes[i];
        var data = 生成大数据(size);
        
        var start = new Date().getTime();
        var result = Array2D.z超级透视(
            data,
            ['f1', '类别'],
            ['f2', '年份'],
            ['sum("f4")', '销售额'],
            0, 0
        );
        var duration = new Date().getTime() - start;
        
        results.push({ size: size, duration: duration });
    }
    
    if (TEST_CONFIG.verbose) {
        Console.log("性能测试结果:");
        for (var i = 0; i < results.length; i++) {
            Console.log("  " + results[i].size + "行: " + results[i].duration + "ms");
        }
    }
    
    断言(results[0].duration < 1000, '100行应在1秒内完成');
}

// ==================== 结果输出 ====================

/**
 * 生成测试报告
 */
function 生成测试报告() {
    var passed = 0;
    var failed = 0;
    var totalDuration = 0;
    
    for (var i = 0; i < TEST_RESULTS.length; i++) {
        if (TEST_RESULTS[i].passed) {
            passed++;
        } else {
            failed++;
        }
        totalDuration += TEST_RESULTS[i].duration;
    }
    
    var report = [];
    report.push(['superPivot v3.9.0 测试报告']);
    report.push(['===================']);
    report.push(['测试时间: ' + new Date().toLocaleString()]);
    report.push(['测试数量: ' + TEST_RESULTS.length]);
    report.push(['通过: ' + passed]);
    report.push(['失败: ' + failed]);
    report.push(['总耗时: ' + totalDuration + 'ms']);
    report.push(['']);
    report.push(['详细结果:']);
    report.push(['测试名称', '状态', '耗时(ms)', '消息']);
    
    for (var i = 0; i < TEST_RESULTS.length; i++) {
        var r = TEST_RESULTS[i];
        report.push([
            r.name,
            r.passed ? '通过' : '失败',
            r.duration,
            r.message
        ]);
    }
    
    return report;
}

/**
 * 输出报告到工作表
 */
function 输出报告到工作表(report) {
    var wb = Application.ActiveWorkbook;
    var ws;
    
    try {
        ws = wb.Worksheets(TEST_CONFIG.outputSheet);
        ws.Cells.Clear();
    } catch (e) {
        ws = wb.Worksheets.Add();
        ws.Name = TEST_CONFIG.outputSheet;
    }
    
    // 找到最大列数
    var maxCols = 0;
    for (var i = 0; i < report.length; i++) {
        if (report[i].length > maxCols) {
            maxCols = report[i].length;
        }
    }
    
    // 输出到工作表
    for (var i = 0; i < report.length; i++) {
        for (var j = 0; j < report[i].length; j++) {
            ws.Cells(i + 1, j + 1).Value2 = report[i][j];
        }
    }
    
    // 格式化
    ws.Range("A1").Font.Bold = true;
    ws.Range("A1").Font.Size = 14;
    ws.Range("A10:D10").Font.Bold = true;
    ws.Columns.AutoFit();
    
    Console.log("✓ 测试报告已输出到【" + TEST_CONFIG.outputSheet + "】工作表");
}

// ==================== 主入口 ====================

/**
 * 运行全部测试
 */
function 运行全部测试() {
    Console.log("╔══════════════════════════════════════════════════════╗");
    Console.log("║     superPivot v3.9.0 全面测试套件 v1.0              ║");
    Console.log("║     开始时间: " + new Date().toLocaleString() + "          ║");
    Console.log("╚══════════════════════════════════════════════════════╝");
    Console.log("");
    
    TEST_RESULTS = [];
    TEST_START_TIME = new Date().getTime();
    
    // 准备测试数据
    准备测试数据();
    Console.log("");
    
    // 基础功能测试
    Console.log("【基础功能测试】");
    运行测试("基础透视", 测试_基础透视);
    运行测试("多行字段", 测试_多行字段);
    运行测试("多列字段", 测试_多列字段);
    运行测试("多数据字段", 测试_多数据字段);
    Console.log("");
    
    // v3.9.0 新功能测试
    Console.log("【v3.9.0 新功能测试】");
    运行测试("行小计", 测试_行小计);
    运行测试("列小计", 测试_列小计);
    运行测试("总计", 测试_总计);
    运行测试("百分比-占总计", 测试_百分比_占总计);
    运行测试("百分比-占行", 测试_百分比_占行);
    运行测试("角标题", 测试_角标题);
    运行测试("层级缩进", 测试_层级缩进);
    运行测试("完整功能组合", 测试_完整功能组合);
    Console.log("");
    
    // 边界情况测试
    Console.log("【边界情况测试】");
    运行测试("空数据", 测试_空数据);
    运行测试("单值列", 测试_单值列);
    运行测试("无表头模式", 测试_无表头模式);
    Console.log("");
    
    // 性能测试
    Console.log("【性能测试】");
    运行测试("大数据性能", 测试_大数据性能);
    Console.log("");
    
    TEST_END_TIME = new Date().getTime();
    var totalTime = TEST_END_TIME - TEST_START_TIME;
    
    // 生成并输出报告
    var report = 生成测试报告();
    输出报告到工作表(report);
    
    // 控制台汇总
    Console.log("╔══════════════════════════════════════════════════════╗");
    Console.log("║                   测试完成汇总                        ║");
    Console.log("╠══════════════════════════════════════════════════════╣");
    
    var passed = 0;
    var failed = 0;
    for (var i = 0; i < TEST_RESULTS.length; i++) {
        if (TEST_RESULTS[i].passed) passed++;
        else failed++;
    }
    
    Console.log("║ 总测试数: " + 左填充(TEST_RESULTS.length, 4) + "                                 ║");
    Console.log("║ 通过:     " + 左填充(passed, 4) + " ✓                               ║");
    Console.log("║ 失败:     " + 左填充(failed, 4) + (failed > 0 ? " ✗" : " ✓") + "                               ║");
    Console.log("║ 总耗时:   " + 左填充(totalTime, 6) + " ms                           ║");
    Console.log("╚══════════════════════════════════════════════════════╝");
    
    if (failed > 0) {
        Console.log("");
        Console.log("失败的测试:");
        for (var i = 0; i < TEST_RESULTS.length; i++) {
            if (!TEST_RESULTS[i].passed) {
                Console.log("  ❌ " + TEST_RESULTS[i].name + ": " + TEST_RESULTS[i].message);
            }
        }
    }
    
    Console.log("");
    Console.log("✓ 详细报告已输出到【" + TEST_CONFIG.outputSheet + "】工作表");
    
    return {
        total: TEST_RESULTS.length,
        passed: passed,
        failed: failed,
        duration: totalTime,
        results: TEST_RESULTS
    };
}

/**
 * 运行单项测试（用于调试）
 */
function 运行单项测试(testName) {
    TEST_RESULTS = [];
    
    var tests = {
        '基础透视': 测试_基础透视,
        '多行字段': 测试_多行字段,
        '多列字段': 测试_多列字段,
        '多数据字段': 测试_多数据字段,
        '行小计': 测试_行小计,
        '列小计': 测试_列小计,
        '总计': 测试_总计,
        '百分比-占总计': 测试_百分比_占总计,
        '百分比-占行': 测试_百分比_占行,
        '角标题': 测试_角标题,
        '层级缩进': 测试_层级缩进,
        '完整功能组合': 测试_完整功能组合,
        '空数据': 测试_空数据,
        '单值列': 测试_单值列,
        '无表头模式': 测试_无表头模式,
        '大数据性能': 测试_大数据性能
    };
    
    if (tests[testName]) {
        Console.log("运行单项测试: " + testName);
        准备测试数据();
        var success = 运行测试(testName, tests[testName]);
        return success;
    } else {
        Console.log("未知测试: " + testName);
        Console.log("可用测试: " + Object.keys(tests).join(", "));
        return false;
    }
}

// 辅助函数
function 左填充(str, length) {
    str = String(str);
    while (str.length < length) {
        str = str + " ";
    }
    return str;
}

// 快捷入口
function 快速测试() {
    运行单项测试('基础透视');
}

function 测试v390() {
    TEST_RESULTS = [];
    准备测试数据();
    
    Console.log("【v3.9.0 核心功能快速验证】");
    运行测试("行小计", 测试_行小计);
    运行测试("列小计", 测试_列小计);
    运行测试("总计", 测试_总计);
    运行测试("百分比-占总计", 测试_百分比_占总计);
    运行测试("角标题", 测试_角标题);
    
    Console.log("");
    Console.log("验证完成，详细结果请查看【测试结果】工作表");
}

// 导出
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        运行全部测试: 运行全部测试,
        运行单项测试: 运行单项测试,
        快速测试: 快速测试,
        测试v390: 测试v390
    };
}
