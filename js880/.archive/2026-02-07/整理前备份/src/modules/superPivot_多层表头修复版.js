/**
 * =======================================================================
 * superPivot v3.9.1 多层表头修复版
 * =======================================================================
 * 
 * 修复内容:
 *   - 修复多层列字段表头结构
 *   - 修复前导空白列计算
 *   - 修正表头行数和列对齐
 * 
 * 正确的表头结构:
 *   - 单列字段 + 单行字段: 2行表头 (列值行 + 字段名/数据标题行)
 *   - 双列字段 + 单行字段: 3行表头 (列值行1 + 列值行2 + 字段名/数据标题行)
 *   - 单列字段 + 双行字段: 2行表头 + 前导空白
 *   - 双列字段 + 双行字段: 3行表头 + 前导空白
 * =======================================================================
 */

/**
 * 修复版 superPivot 函数 - 仅表头生成部分
 * 用于替换 JSA880.js 中的对应部分
 */
function 修复版_superPivot_表头生成(data, rowFields, colFields, dataFields, options) {
    options = options || {};
    
    var separator = '@^@';
    var rowFieldIndent = options.rowFieldIndent !== false;
    var rowFieldIndentSize = options.rowFieldIndentSize || 4;
    var layoutMode = options.layoutMode || 'outline';
    var cornerTitle = options.cornerTitle || '';
    var colSubtotals = options.colSubtotals || { enabled: false };
    
    // 解析字段配置
    function parseFields(fields) {
        var result = { fields: [], titles: [], hasTitles: false };
        if (typeof fields === 'string') {
            var parts = fields.split(',');
            for (var i = 0; i < parts.length; i++) {
                var match = parts[i].trim().match(/^(f\d+)([+\-#]*)$/);
                if (match) {
                    result.fields.push({ field: match[1], sort: match[2] || '+' });
                    result.titles.push('');
                }
            }
        } else if (Array.isArray(fields) && fields.length >= 2) {
            var fieldParts = fields[0].split(',');
            var titleParts = fields[1].split(',');
            for (var i = 0; i < fieldParts.length; i++) {
                var match = fieldParts[i].trim().match(/^(f\d+)([+\-#]*)$/);
                if (match) {
                    result.fields.push({ field: match[1], sort: match[2] || '+' });
                    result.titles.push(titleParts[i] || '');
                }
            }
            result.hasTitles = true;
        }
        return result;
    }
    
    var rowConfig = parseFields(rowFields);
    var colConfig = parseFields(colFields);
    
    var numRowFieldLevels = rowConfig.fields.length;
    var numColFieldLevels = colConfig.fields.length;
    
    // 解析数据字段
    var numDataFields = 1;
    if (typeof dataFields === 'string') {
        var matches = dataFields.match(/\w+\s*\(/g);
        if (matches) numDataFields = matches.length;
    } else if (Array.isArray(dataFields) && dataFields.length >= 2) {
        var titles = dataFields[1].split(',');
        numDataFields = titles.length;
    }
    
    // ========== 关键修复1: 正确计算表头结构 ==========
    
    // 表头行数 = 列字段数 + 1 (数据标题行)
    // 注意：不是简单的 numColFieldLevels + 1
    // 而是要根据实际情况调整
    var headerRowCount;
    if (numColFieldLevels === 1) {
        // 单列字段: 
        // - 如果有行字段：2行 (列值行 + 字段名行)
        // - 行字段标题放在第2行
        headerRowCount = 2;
    } else {
        // 多列字段: 列字段数 + 1
        headerRowCount = numColFieldLevels + 1;
    }
    
    // 前导空白列数 = max(0, 行字段数 - 列字段数)
    var leadingBlankCols = Math.max(0, numRowFieldLevels - numColFieldLevels);
    
    Console.log("=== 表头结构分析 ===");
    Console.log("行字段数: " + numRowFieldLevels);
    Console.log("列字段数: " + numColFieldLevels);
    Console.log("前导空白列: " + leadingBlankCols);
    Console.log("表头行数: " + headerRowCount);
    Console.log("数据字段数: " + numDataFields);
    
    // ========== 关键修复2: 正确构建表头 ==========
    
    var headerRows = [];
    for (var h = 0; h < headerRowCount; h++) {
        headerRows.push([]);
    }
    
    // 模拟 colKeys (实际应从数据中提取)
    var colKeys = [];
    if (numColFieldLevels === 1) {
        colKeys = ['2023', '2024'];
    } else if (numColFieldLevels === 2) {
        colKeys = ['2023@^@Q1', '2023@^@Q2', '2024@^@Q1', '2024@^@Q2'];
    }
    
    // 填充表头
    if (numColFieldLevels === 1) {
        // ===== 单列字段情况 =====
        // 结构:
        // Row 0: [前导空白] [列字段标题] [列值, 列值, ...] [小计]
        // Row 1: [行字段标题] [数据标题, 数据标题, ...] [小计标题]
        
        // Row 0: 前导空白 + 列字段标题 + 列值
        for (var i = 0; i < leadingBlankCols; i++) {
            headerRows[0].push('');
        }
        headerRows[0].push(colConfig.titles[0] || '列字段');
        
        for (var ck = 0; ck < colKeys.length; ck++) {
            headerRows[0].push(colKeys[ck]);
        }
        if (colSubtotals.enabled) {
            headerRows[0].push(colSubtotals.label || '小计');
        }
        
        // Row 1: 行字段标题 + 数据字段标题
        for (var i = 0; i < numRowFieldLevels; i++) {
            headerRows[1].push(rowConfig.titles[i] || '行字段' + (i+1));
        }
        
        for (var ck = 0; ck < colKeys.length; ck++) {
            for (var df = 0; df < numDataFields; df++) {
                headerRows[1].push('数据' + (df+1));
            }
        }
        if (colSubtotals.enabled) {
            for (var df = 0; df < numDataFields; df++) {
                headerRows[1].push('小计');
            }
        }
        
    } else {
        // ===== 多列字段情况 =====
        // 结构:
        // Row 0: [前导空白] [列字段1标题] [列值1, 列值1, ...]
        // Row 1: [前导空白] [列字段2标题] [列值2, 列值2, ...]
        // ...
        // Row N: [行字段标题] [数据标题, ...]
        
        for (var cfIdx = 0; cfIdx < numColFieldLevels; cfIdx++) {
            var targetRow = cfIdx;
            
            // 前导空白
            for (var i = 0; i < leadingBlankCols; i++) {
                headerRows[targetRow].push('');
            }
            
            // 列字段标题
            if (cfIdx === 0 && cornerTitle) {
                headerRows[targetRow].push(cornerTitle);
            } else {
                headerRows[targetRow].push(colConfig.titles[cfIdx] || '列字段' + (cfIdx+1));
            }
            
            // 列值 (每个值重复 numDataFields 次)
            for (var ck = 0; ck < colKeys.length; ck++) {
                var parts = colKeys[ck].split(separator);
                for (var df = 0; df < numDataFields; df++) {
                    headerRows[targetRow].push(parts[cfIdx] || '');
                }
            }
            
            // 小计标题
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
        
        for (var i = 0; i < numRowFieldLevels; i++) {
            headerRows[lastRow].push(rowConfig.titles[i] || '行字段' + (i+1));
        }
        
        for (var ck = 0; ck < colKeys.length; ck++) {
            for (var df = 0; df < numDataFields; df++) {
                headerRows[lastRow].push('数据' + (df+1));
            }
        }
        if (colSubtotals.enabled) {
            for (var df = 0; df < numDataFields; df++) {
                headerRows[lastRow].push('小计');
            }
        }
    }
    
    // 输出表头结构
    Console.log("\n=== 生成的表头结构 ===");
    for (var i = 0; i < headerRows.length; i++) {
        Console.log("Row " + i + ": [" + headerRows[i].join(", ") + "]");
    }
    
    return headerRows;
}

// ==================== 测试用例 ====================

function 测试_单列字段_单行字段() {
    Console.log("\n========== 测试: 单列字段 + 单行字段 ==========");
    var data = [['产品', '年份', '销售额']];
    return 修复版_superPivot_表头生成(
        data,
        ['f1', '产品'],
        ['f2', '年份'],
        ['sum("f3")', '销售额']
    );
}

function 测试_双列字段_单行字段() {
    Console.log("\n========== 测试: 双列字段 + 单行字段 ==========");
    var data = [['产品', '年份', '季度', '销售额']];
    return 修复版_superPivot_表头生成(
        data,
        ['f1', '产品'],
        ['f2,f3', '年份,季度'],
        ['sum("f4")', '销售额']
    );
}

function 测试_单列字段_双行字段() {
    Console.log("\n========== 测试: 单列字段 + 双行字段 ==========");
    var data = [['类别', '产品', '年份', '销售额']];
    return 修复版_superPivot_表头生成(
        data,
        ['f1,f2', '类别,产品'],
        ['f3', '年份'],
        ['sum("f4")', '销售额']
    );
}

function 测试_双列字段_双行字段() {
    Console.log("\n========== 测试: 双列字段 + 双行字段 ==========");
    var data = [['类别', '产品', '年份', '季度', '销售额']];
    return 修复版_superPivot_表头生成(
        data,
        ['f1,f2', '类别,产品'],
        ['f3,f4', '年份,季度'],
        ['sum("f5")', '销售额']
    );
}

function 测试_双列字段_双行字段_多数据() {
    Console.log("\n========== 测试: 双列字段 + 双行字段 + 多数据字段 ==========");
    var data = [['类别', '产品', '年份', '季度', '销售额', '数量']];
    return 修复版_superPivot_表头生成(
        data,
        ['f1,f2', '类别,产品'],
        ['f3,f4', '年份,季度'],
        ['sum("f5"),sum("f6")', '销售额,数量']
    );
}

function 测试_带角标题() {
    Console.log("\n========== 测试: 双列字段 + 双行字段 + 角标题 ==========");
    var data = [['类别', '产品', '年份', '季度', '销售额']];
    return 修复版_superPivot_表头生成(
        data,
        ['f1,f2', '类别,产品'],
        ['f3,f4', '年份,季度'],
        ['sum("f5")', '销售额'],
        { cornerTitle: '销售分析表' }
    );
}

function 测试_带列小计() {
    Console.log("\n========== 测试: 单列字段 + 单行字段 + 列小计 ==========");
    var data = [['产品', '年份', '销售额']];
    return 修复版_superPivot_表头生成(
        data,
        ['f1', '产品'],
        ['f2', '年份'],
        ['sum("f3")', '销售额'],
        { colSubtotals: { enabled: true, label: '年度小计' } }
    );
}

// ==================== 运行所有测试 ====================

function 运行所有表头测试() {
    Console.log("╔══════════════════════════════════════════════════════╗");
    Console.log("║     superPivot v3.9.1 多层表头结构测试               ║");
    Console.log("╚══════════════════════════════════════════════════════╝");
    
    测试_单列字段_单行字段();
    测试_双列字段_单行字段();
    测试_单列字段_双行字段();
    测试_双列字段_双行字段();
    测试_双列字段_双行字段_多数据();
    测试_带角标题();
    测试_带列小计();
    
    Console.log("\n✓ 所有表头结构测试完成");
}

// 导出
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        运行所有表头测试: 运行所有表头测试,
        测试_单列字段_单行字段: 测试_单列字段_单行字段,
        测试_双列字段_单行字段: 测试_双列字段_单行字段,
        测试_单列字段_双行字段: 测试_单列字段_双行字段,
        测试_双列字段_双行字段: 测试_双列字段_双行字段
    };
}

// 如果直接运行，执行测试
if (typeof Application !== 'undefined') {
    运行所有表头测试();
}
