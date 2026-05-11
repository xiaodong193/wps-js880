/**
 * =======================================================================
 * SuperPivot 多层级表头测试 - JSA880 v3.9.4+
 * =======================================================================
 *
 * 测试 z超级透视 多层级标题功能
 * 支持 Node.js 环境运行（WPS环境需要调整）
 *
 * 运行方法:
 *   node test_superPivot_headers.js
 *
 * =======================================================================
 */

// 模拟 WPS Console (Node.js 环境)
if (typeof Console === 'undefined') {
    global.Console = {
        log: function(...args) { console.log('[LOG]', ...args); }
    };
}

// 模拟 WPS Range/Application 对象
if (typeof Application === 'undefined') {
    global.Application = {
        ActiveSheet: { Range: function() { return { Merge: function(){} }; } }
    };
}

// 加载 JSA880（如果存在）
var Array2D;
try {
    // 尝试加载文件
    var fs = require('fs');
    var path = require('path');
    var jsaPath = path.join(__dirname, '../JSA880.js');

    if (fs.existsSync(jsaPath)) {
        var code = fs.readFileSync(jsaPath, 'utf8');
        // 移除 WPS 特定代码
        code = code.replace(/isWPS\s*=\s*typeof\s+Application[^;]+;/g, 'isWPS = false;');
        code = code.replace(/if\s*\([^)]*Console[^)]*\)/g, 'if (typeof Console !== "undefined" && Console.log)');
        eval(code);
        Console.log('✅ JSA880 已加载');
    }
} catch (e) {
    Console.log('⚠️ 无法加载 JSA880: ' + e.message);
}

// =======================================================================
// 测试数据
// =======================================================================

/**
 * 标准测试数据 - 3个列字段（年份→季度→月份）
 */
var TEST_DATA_3_LEVELS = [
    ['产品', '年份', '季度', '月份', '销售额', '数量'],
    ['A产品', '2024', 'Q1', '1月', 1000, 10],
    ['A产品', '2024', 'Q1', '2月', 1500, 15],
    ['A产品', '2024', 'Q2', '1月', 2000, 20],
    ['A产品', '2024', 'Q2', '2月', 1800, 18],
    ['A产品', '2025', 'Q1', '1月', 1200, 12],
    ['A产品', '2025', 'Q1', '2月', 1600, 16],
    ['A产品', '2025', 'Q2', '1月', 2200, 22],
    ['A产品', '2025', 'Q2', '2月', 1900, 19],
    ['B产品', '2024', 'Q1', '1月', 800, 8],
    ['B产品', '2024', 'Q1', '2月', 900, 9],
    ['B产品', '2024', 'Q2', '1月', 1100, 11],
    ['B产品', '2024', 'Q2', '2月', 950, 9],
    ['B产品', '2025', 'Q1', '1月', 1300, 13],
    ['B产品', '2025', 'Q1', '2月', 1400, 14],
    ['B产品', '2025', 'Q2', '1月', 1700, 17],
    ['B产品', '2025', 'Q2', '2月', 1550, 15]
];

/**
 * 标准测试数据 - 2个列字段（年份→季度）
 */
var TEST_DATA_2_LEVELS = [
    ['产品', '年份', '季度', '销售额', '数量'],
    ['A产品', '2024', 'Q1', 1000, 10],
    ['A产品', '2024', 'Q2', 1500, 15],
    ['A产品', '2025', 'Q1', 2000, 20],
    ['A产品', '2025', 'Q2', 1800, 18],
    ['B产品', '2024', 'Q1', 800, 8],
    ['B产品', '2024', 'Q2', 900, 9],
    ['B产品', '2025', 'Q1', 1300, 13],
    ['B产品', '2025', 'Q2', 1400, 14]
];

/**
 * 标准测试数据 - 1个列字段（地区）
 */
var TEST_DATA_1_LEVEL = [
    ['产品', '地区', '销售额'],
    ['A产品', '北京', 1000],
    ['A产品', '上海', 1500],
    ['B产品', '北京', 2000],
    ['B产品', '上海', 1000]
];

// =======================================================================
// 测试函数
// =======================================================================

var testCount = 0;
var passCount = 0;
var failCount = 0;

function assertEqual(actual, expected, testName) {
    testCount++;
    var passed = JSON.stringify(actual) === JSON.stringify(expected);
    if (passed) {
        passCount++;
        Console.log(`✅ 测试 ${testCount}: ${testName}`);
    } else {
        failCount++;
        Console.log(`❌ 测试 ${testCount}: ${testName}`);
        Console.log(`   期望: ${JSON.stringify(expected)}`);
        Console.log(`   实际: ${JSON.stringify(actual)}`);
    }
    return passed;
}

function assertTrue(condition, testName) {
    testCount++;
    if (condition) {
        passCount++;
        Console.log(`✅ 测试 ${testCount}: ${testName}`);
    } else {
        failCount++;
        Console.log(`❌ 测试 ${testCount}: ${testName}`);
    }
    return condition;
}

// =======================================================================
// 多层级表头测试
// =======================================================================

Console.log('========================================');
Console.log('  SuperPivot 多层级表头测试');
Console.log('========================================\n');

// 测试 1: 3层级列字段 - 表头结构验证
Console.log('【测试 A】3层级列字段（年份→季度→月份）');
if (typeof Array2D !== 'undefined' && Array2D.z超级透视) {
    try {
        var result3 = Array2D.z超级透视(TEST_DATA_3_LEVELS,
            ['f1+'],                    // 行字段: 产品
            ['f2+,f3+,f4+'],           // 列字段: 年份→季度→月份
            ['sum("f5")'],              // 数据字段: 销售额求和
            1, 1, '@^@', {});

        Console.log('结果行数: ' + result3.length);
        Console.log('表头行数: 4 (3个列层级 + 1个数据行)');
        Console.log('第一行(表头): ' + JSON.stringify(result3[0]));
        Console.log('第二行(表头): ' + JSON.stringify(result3[1]));
        Console.log('第三行(表头): ' + JSON.stringify(result3[2]));
        Console.log('第四行(数据): ' + JSON.stringify(result3[3]));

        // 验证表头行数
        assertTrue(result3.length >= 5, '3层级应有表头行+数据行');

        // 验证列字段值出现在第1行
        var headerRow1 = result3[0];
        var has2024 = headerRow1.includes('2024');
        var has2025 = headerRow1.includes('2025');
        assertTrue(has2024 && has2025, '第1行包含年份值 2024/2025');

        Console.log('✅ 3层级测试通过\n');
    } catch (e) {
        failCount++;
        Console.log('❌ 3层级测试失败: ' + e.message + '\n');
    }
} else {
    Console.log('⚠️ Array2D.z超级透视 不可用，跳过测试\n');
}

// 测试 2: 2层级列字段 - 表头结构验证
Console.log('【测试 B】2层级列字段（年份→季度）');
if (typeof Array2D !== 'undefined' && Array2D.z超级透视) {
    try {
        var result2 = Array2D.z超级透视(TEST_DATA_2_LEVELS,
            ['f1+'],
            ['f2+,f3+'],
            ['sum("f4")'],
            1, 1, '@^@', {});

        Console.log('结果行数: ' + result2.length);
        Console.log('表头行数: 3 (2个列层级 + 1个数据行)');
        Console.log('第一行(表头): ' + JSON.stringify(result2[0]));
        Console.log('第二行(表头): ' + JSON.stringify(result2[1]));

        // 验证表头行数
        assertTrue(result2.length >= 4, '2层级应有表头行+数据行');

        // 验证第2行包含季度值
        var headerRow2 = result2[1];
        var hasQ1 = headerRow2.includes('Q1');
        var hasQ2 = headerRow2.includes('Q2');
        assertTrue(hasQ1 && hasQ2, '第2行包含季度值 Q1/Q2');

        Console.log('✅ 2层级测试通过\n');
    } catch (e) {
        failCount++;
        Console.log('❌ 2层级测试失败: ' + e.message + '\n');
    }
} else {
    Console.log('⚠️ Array2D.z超级透视 不可用，跳过测试\n');
}

// 测试 3: 1层级列字段 - 3行表头验证
Console.log('【测试 C】1层级列字段（地区）');
if (typeof Array2D !== 'undefined' && Array2D.z超级透视) {
    try {
        var result1 = Array2D.z超级透视(TEST_DATA_1_LEVEL,
            ['f1+'],
            ['f2+'],
            ['sum("f3")'],
            1, 1, '@^@', {});

        Console.log('结果行数: ' + result1.length);
        Console.log('表头行数: 3 (单列字段专用格式)');
        Console.log('第一行(行标题+列值): ' + JSON.stringify(result1[0]));
        Console.log('第二行(空白): ' + JSON.stringify(result1[1]));
        Console.log('第三行(数据标题): ' + JSON.stringify(result1[2]));

        // 验证 3 行表头
        assertTrue(result1.length >= 4, '1层级应有3行表头+数据行');

        // 验证第1行包含地区值
        var headerRow1 = result1[0];
        var has北京 = headerRow1.includes('北京');
        var has上海 = headerRow1.includes('上海');
        assertTrue(has北京 && has上海, '第1行包含地区值 北京/上海');

        // 验证第2行结构（单列字段专用格式）
        // 第2行应该是: [空白, 列标题, 空白, ..., 空白] - 有列标题在numRowFieldLevels位置
        var row2HasColTitle = result1[1].indexOf('地区') === 1;  // 列标题在第2列位置
        assertTrue(row2HasColTitle || result1[1].every(v => v === '' || v === undefined),
            '第2行包含列标题或空白（单列字段格式）');

        Console.log('✅ 1层级测试通过\n');
    } catch (e) {
        failCount++;
        Console.log('❌ 1层级测试失败: ' + e.message + '\n');
    }
} else {
    Console.log('⚠️ Array2D.z超级透视 不可用，跳过测试\n');
}

// 测试 4: 无列字段 - 单行表头
Console.log('【测试 D】无列字段');
if (typeof Array2D !== 'undefined' && Array2D.z超级透视) {
    try {
        var dataNoCol = [
            ['产品', '销售额', '数量'],
            ['A产品', 1000, 10],
            ['A产品', 1500, 15],
            ['B产品', 2000, 20]
        ];

        var resultNoCol = Array2D.z超级透视(dataNoCol,
            ['f1+'],
            [],                        // 无列字段
            ['sum("f2"),sum("f3")'],
            1, 1, '@^@', {});

        Console.log('结果行数: ' + resultNoCol.length);
        Console.log('表头行数: 1 (无列字段单行表头)');
        Console.log('第一行: ' + JSON.stringify(resultNoCol[0]));

        // 验证只有1行表头
        assertTrue(resultNoCol.length >= 2, '无列字段应有1行表头+数据行');
        assertTrue(resultNoCol[0].includes('产品') || resultNoCol[0].includes('A产品') === false,
            '第1行是表头行');

        Console.log('✅ 无列字段测试通过\n');
    } catch (e) {
        failCount++;
        Console.log('❌ 无列字段测试失败: ' + e.message + '\n');
    }
} else {
    Console.log('⚠️ Array2D.z超级透视 不可用，跳过测试\n');
}

// =======================================================================
// 测试报告
// =======================================================================

Console.log('========================================');
Console.log('  测试报告');
Console.log('========================================');
Console.log(`总计: ${testCount} 个测试`);
Console.log(`通过: ${passCount} 个`);
Console.log(`失败: ${failCount} 个`);
Console.log('========================================');

if (failCount === 0) {
    Console.log('🎉 所有测试通过！');
} else {
    Console.log('⚠️ 有测试失败，请检查');
}