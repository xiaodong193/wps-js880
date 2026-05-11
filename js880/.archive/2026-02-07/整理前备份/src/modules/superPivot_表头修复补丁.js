/**
 * =======================================================================
 * superPivot v3.9.1 表头修复补丁
 * =======================================================================
 * 
 * 使用说明:
 *   1. 在 JSA880.js 中找到 superPivot 函数的表头生成部分
 *   2. 用此文件中的代码替换对应部分
 *   3. 注意保留周围的上下文代码
 * =======================================================================
 */

// ===== 修复1: 表头行数计算 (替换第 6677 行) =====

// 【替换前】
// var headerRowCount = (numColFieldLevels === 1) ? 3 : numColFieldLevels + 1;

// 【替换后】
var headerRowCount = numColFieldLevels + 1;
var leadingBlankCols = Math.max(0, numRowFieldLevels - numColFieldLevels);


// ===== 修复2: 前导空白列计算 (替换第 6862-6867 行) =====

// 【替换前】
// if (!hideRowTitles) {
//     for (var rfIdx = 0; rfIdx < numRowFieldLevels - 1; rfIdx++) {
//         headerRows[targetRow].push('');
//     }
// }

// 【替换后】
if (!hideRowTitles) {
    for (var i = 0; i < leadingBlankCols; i++) {
        headerRows[targetRow].push('');
    }
}


// ===== 修复3: 完整的表头生成逻辑 (替换第 6767-6958 行) =====

// 【完整替换代码】
if (numColFieldLevels === 1) {
    // ===== 单列字段情况 =====
    // 表头结构:
    // Row 0: [前导空白] [列字段标题] [列值, 列值, ...] [小计]
    // Row 1: [行字段标题] [数据标题, 数据标题, ...] [小计标题]
    
    if (!hideRowTitles) {
        // Row 0: 前导空白 + 列字段标题
        for (var i = 0; i < leadingBlankCols; i++) {
            headerRows[0].push('');
        }
        
        // 角标题或列字段标题
        if (cornerTitle && leadingBlankCols === 0) {
            headerRows[0].push(cornerTitle);
        } else {
            var colTitle = '';
            if (hasColTitles) {
                colTitle = colConfig.titles[0] || '';
            } else if (_originalHeader) {
                var match = colConfig.fields[0].field.match(/^f(\d+)$/);
                if (match) {
                    colTitle = _originalHeader[parseInt(match[1]) - 1] || '';
                }
            }
            headerRows[0].push(colTitle);
        }
        
        // 列值
        for (var ck = 0; ck < colKeys.length; ck++) {
            headerRows[0].push(colKeys[ck]);
        }
        
        // 列小计
        if (colSubtotals.enabled) {
            headerRows[0].push(colSubtotals.label || '小计');
        }
        
        // Row 1: 行字段标题
        for (var rfIdx = 0; rfIdx < numRowFieldLevels; rfIdx++) {
            var rowTitle = '';
            if (hasRowTitles) {
                rowTitle = rowConfig.titles[rfIdx] || '';
            } else if (_originalHeader) {
                var match = rowConfig.fields[rfIdx].field.match(/^f(\d+)$/);
                if (match) {
                    rowTitle = _originalHeader[parseInt(match[1]) - 1] || '';
                }
            }
            headerRows[1].push(rowTitle);
        }
        
        // 数据字段标题
        for (var ck = 0; ck < colKeys.length; ck++) {
            for (var dfIdx = 0; dfIdx < numDataFields; dfIdx++) {
                headerRows[1].push(defaultDataTitles[dfIdx] || '');
            }
        }
        
        // 列小计的数据字段标题
        if (colSubtotals.enabled) {
            for (var dfIdx = 0; dfIdx < numDataFields; dfIdx++) {
                headerRows[1].push(defaultDataTitles[dfIdx] || '');
            }
        }
    }
    
} else {
    // ===== 多列字段情况 =====
    // 表头结构:
    // Row 0: [前导空白] [列字段1标题] [列值1, 列值1, ...]
    // Row 1: [前导空白] [列字段2标题] [列值2, 列值2, ...]
    // ...
    // Row N: [行字段标题] [数据标题, ...]
    
    for (var cfIdx = 0; cfIdx < numColFieldLevels; cfIdx++) {
        var targetRow = cfIdx;
        
        // 前导空白
        if (!hideRowTitles) {
            for (var i = 0; i < leadingBlankCols; i++) {
                headerRows[targetRow].push('');
            }
        }
        
        // 列字段标题（第一行可能是角标题）
        if (cfIdx === 0 && cornerTitle && !hideRowTitles) {
            headerRows[targetRow].push(cornerTitle);
        } else if (!hideRowTitles || cfIdx > 0) {
            var colTitle = '';
            if (hasColTitles) {
                colTitle = colConfig.titles[cfIdx] || '';
            } else if (_originalHeader) {
                var match = colConfig.fields[cfIdx].field.match(/^f(\d+)$/);
                if (match) {
                    colTitle = _originalHeader[parseInt(match[1]) - 1] || '';
                }
            }
            headerRows[targetRow].push(colTitle);
        }
        
        // 列值（每个值重复 numDataFields 次）
        for (var ck = 0; ck < colKeys.length; ck++) {
            var colKeyParts = colKeys[ck].split(separator);
            if (cfIdx < colKeyParts.length) {
                for (var df = 0; df < numDataFields; df++) {
                    headerRows[targetRow].push(colKeyParts[cfIdx]);
                }
            } else {
                for (var df = 0; df < numDataFields; df++) {
                    headerRows[targetRow].push('');
                }
            }
        }
        
        // 列小计
        if (colSubtotals.enabled) {
            if (cfIdx === numColFieldLevels - 1) {
                headerRows[targetRow].push(colSubtotals.label || '小计');
            } else {
                headerRows[targetRow].push('');
            }
        }
    }
    
    // 最后一行: 行字段标题 + 数据字段标题
    var lastRow = numColFieldLevels;
    
    if (!hideRowTitles) {
        for (var rfIdx = 0; rfIdx < numRowFieldLevels; rfIdx++) {
            var rowTitle = '';
            if (hasRowTitles) {
                rowTitle = rowConfig.titles[rfIdx] || '';
            } else if (_originalHeader) {
                var match = rowConfig.fields[rfIdx].field.match(/^f(\d+)$/);
                if (match) {
                    rowTitle = _originalHeader[parseInt(match[1]) - 1] || '';
                }
            }
            headerRows[lastRow].push(rowTitle);
        }
    }
    
    // 数据字段标题
    for (var ck = 0; ck < colKeys.length; ck++) {
        for (var dfIdx = 0; dfIdx < numDataFields; dfIdx++) {
            headerRows[lastRow].push(defaultDataTitles[dfIdx] || '');
        }
    }
    
    // 列小计的数据字段标题
    if (colSubtotals.enabled) {
        for (var dfIdx = 0; dfIdx < numDataFields; dfIdx++) {
            headerRows[lastRow].push(defaultDataTitles[dfIdx] || '');
        }
    }
}


// ===== 使用示例 =====

/**
 * 应用此补丁后，测试代码：
 */
function 测试修复后的表头() {
    var data = [
        ['类别', '产品', '年份', '季度', '销售额'],
        ['电子', '手机', '2023', 'Q1', 100],
        ['电子', '手机', '2023', 'Q2', 200],
        ['电子', '电脑', '2024', 'Q1', 300]
    ];
    
    // 测试1: 双列字段 + 双行字段
    Console.log("=== 测试: 双列字段 + 双行字段 ===");
    var result1 = Array2D.z超级透视(
        data,
        ['f1,f2', '类别,产品'],
        ['f3,f4', '年份,季度'],
        ['sum("f5")', '销售额']
    );
    
    Console.log("表头行数: " + (result1.length - result1.getMeta().rowCount));
    Console.log("预期: 3行");
    
    // 输出到新工作表查看
    var ws1 = Application.ActiveWorkbook.Worksheets.Add();
    ws1.Name = "修复测试1";
    result1.toRange("A1", true);
    
    // 测试2: 单列字段 + 双行字段
    Console.log("\n=== 测试: 单列字段 + 双行字段 ===");
    var result2 = Array2D.z超级透视(
        data,
        ['f1,f2', '类别,产品'],
        ['f3', '年份'],
        ['sum("f5")', '销售额']
    );
    
    Console.log("表头行数: " + (result2.length - result2.getMeta().rowCount));
    Console.log("预期: 2行");
    
    var ws2 = Application.ActiveWorkbook.Worksheets.Add();
    ws2.Name = "修复测试2";
    result2.toRange("A2", true);
    
    Console.log("\n✓ 测试完成，请查看【修复测试1】和【修复测试2】工作表");
}

// 运行测试
// 测试修复后的表头();
