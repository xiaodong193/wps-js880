/**
 * SuperPivot (z超级透视) Node.js 独立测试脚本
 * 
 * 目的：在 Node.js 环境中测试超级透视核心功能是否正常
 * 方法：mock WPS JSA 全局对象，加载 JSA880.js，运行测试用例
 */

// ==================== Mock WPS 全局对象 ====================
// JSA880.js 引用了这些 WPS 特有的全局对象

// Console mock (WPS 专用控制台)
global.Console = {
    log: function(...args) {
        // 只在有实际输出时打印
        const msg = args.join(' ');
        if (msg && !msg.includes('[superPivot DEBUG]')) {
            // console.log('[WPS Console]', ...args);
        }
    }
};

// Application mock
global.Application = {
    ActiveSheet: {
        Name: '测试工作表'
    },
    Worksheets: function(name) { return { Name: name, Cells: {} }; },
    Range: function() { return { Value2: null }; },
    ScreenUpdating: true,
    Calculation: -4105,
    EnableEvents: true
};

// Console 别名 (某些代码用 console，某些用 Console)
// Node.js 已有 console，不需要额外 mock

// Worksheets 全局函数
global.Worksheets = function(name) {
    return {
        Name: name,
        Cells: {
            ClearContents: function() {},
            ClearFormats: function() {}
        },
        Activate: function() {},
        Range: function() { return { Merge: function() {} }; },
        Columns: { AutoFit: function() {} }
    };
};

// ==================== 加载 JSA880.js ====================
console.log('正在加载 JSA880.js ...');
try {
    const fs = require('fs');
    const path = require('path');
    const jsaCode = fs.readFileSync(path.join(__dirname, 'JSA880.js'), 'utf8');
    // 使用 vm.runInThisContext 在全局作用域执行，使 Array2D 等成为全局变量
    const vm = require('vm');
    vm.runInThisContext(jsaCode, { filename: 'JSA880.js' });
    console.log('✅ JSA880.js 加载成功');
} catch (e) {
    console.error('❌ JSA880.js 加载失败:', e.message);
    console.error(e.stack);
    process.exit(1);
}

// ==================== 验证 Array2D 可用 ====================
if (typeof Array2D === 'undefined') {
    console.error('❌ Array2D 未定义');
    process.exit(1);
}
console.log('✅ Array2D 已定义');

if (typeof Array2D.z超级透视 === 'function') {
    console.log('✅ Array2D.z超级透视 方法存在');
} else {
    console.error('❌ Array2D.z超级透视 方法不存在');
    process.exit(1);
}

if (typeof Array2D.superPivot === 'function') {
    console.log('✅ Array2D.superPivot (英文别名) 方法存在');
}

// ==================== 测试辅助函数 ====================
let testCount = 0;
let passCount = 0;
let failCount = 0;

function test(name, fn) {
    testCount++;
    try {
        fn();
        passCount++;
        console.log(`  ✅ [${testCount}] ${name}`);
    } catch (e) {
        failCount++;
        console.log(`  ❌ [${testCount}] ${name}`);
        console.log(`     错误: ${e.message}`);
    }
}

function assertEqual(actual, expected, msg) {
    if (actual !== expected) {
        throw new Error(`${msg || ''} (期望: ${expected}, 实际: ${actual})`);
    }
}

function assertTrue(condition, msg) {
    if (!condition) {
        throw new Error(msg || '期望为真，实际为假');
    }
}

function printResult(result, title) {
    console.log(`\n  📊 [${title}]`);
    if (result && result.length > 0) {
        for (let i = 0; i < result.length; i++) {
            const row = result[i];
            if (Array.isArray(row)) {
                console.log('     ' + row.map(v => v === null || v === undefined ? '(空)' : String(v)).join(' | '));
            }
        }
    } else {
        console.log('     (无数据)');
    }
}

// ==================== 测试数据 ====================
function createSimpleData() {
    return [
        ["产品", "地区", "销售额"],
        ["A", "北京", 100],
        ["A", "上海", 200],
        ["B", "北京", 300],
        ["B", "上海", 400]
    ];
}

function createTestData() {
    return [
        ["产品", "地区", "年份", "季度", "销售额", "数量"],
        ["A", "北京", 2023, "Q1", 1000, 10],
        ["A", "北京", 2023, "Q2", 1500, 15],
        ["A", "上海", 2023, "Q1", 1200, 12],
        ["A", "上海", 2023, "Q2", 1800, 18],
        ["B", "北京", 2023, "Q1", 2000, 20],
        ["B", "北京", 2023, "Q2", 2500, 25],
        ["B", "上海", 2023, "Q1", 2200, 22],
        ["B", "上海", 2023, "Q2", 2800, 28]
    ];
}

function createMultiLevelData() {
    return [
        ["大区", "省份", "城市", "产品", "销售额"],
        ["华北", "北京", "北京市", "A", 1000],
        ["华北", "北京", "北京市", "B", 2000],
        ["华北", "天津", "天津市", "A", 1500],
        ["华东", "上海", "上海市", "A", 2500],
        ["华东", "上海", "上海市", "B", 3000],
        ["华东", "江苏", "南京市", "A", 1800]
    ];
}

// ==================== 开始测试 ====================
console.log('\n╔══════════════════════════════════════════════════╗');
console.log('║   SuperPivot (z超级透视) 功能测试               ║');
console.log('╚══════════════════════════════════════════════════╝\n');

// ==================== 测试组1: 基础功能 ====================
console.log('━━━ 测试组 1: 基础功能测试 ━━━');

test('基础透视 - 单行字段单列字段', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '基础透视 - 单行单列');
});

test('无列字段 - 仅行字段', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '仅行字段');
});

test('无行字段 - 仅列字段', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, [], ["f2+"], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '仅列字段');
});

// ==================== 测试组2: 多字段测试 ====================
console.log('\n━━━ 测试组 2: 多字段测试 ━━━');

test('多行字段 - 产品和地区', function() {
    const data = createTestData();
    const result = Array2D.z超级透视(data, ["f1+,f2+"], ["f3+"], ["sum(\"f5\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '多行字段');
});

test('多列字段 - 年份和季度', function() {
    const data = createTestData();
    const result = Array2D.z超级透视(data, ["f1+"], ["f3+,f4+"], ["sum(\"f5\")"], 1);
    assertTrue(result.length >= 4, '多列字段应有至少4行');
    printResult(result, '多列字段（多层表头）');
});

test('多行多列字段组合', function() {
    const data = createTestData();
    const result = Array2D.z超级透视(data, ["f1+,f2+"], ["f3+,f4+"], ["sum(\"f5\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '多行多列组合');
});

test('三层行字段 - 大区省份城市', function() {
    const data = createMultiLevelData();
    const result = Array2D.z超级透视(data, ["f1+,f2+,f3+"], [], ["sum(\"f5\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '三层行字段');
});

// ==================== 测试组3: 排序功能 ====================
console.log('\n━━━ 测试组 3: 排序功能测试 ━━━');

test('行字段升序 (+)', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '行字段升序');
});

test('行字段降序 (-)', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1-"], [], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '行字段降序');
});

// ==================== 测试组4: 聚合函数 ====================
console.log('\n━━━ 测试组 4: 聚合函数测试 ━━━');

test('聚合 - count()', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["count()"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, 'count计数');
});

test('聚合 - sum()', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, 'sum求和');
});

test('聚合 - average()', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], [], ["average(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, 'average平均值');
});

test('聚合 - max()', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], [], ["max(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, 'max最大值');
});

test('聚合 - min()', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], [], ["min(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, 'min最小值');
});

test('多聚合函数组合', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], [], ["count(),sum(\"f3\"),average(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '多聚合组合');
});

// ==================== 测试组5: 自定义标题 ====================
console.log('\n━━━ 测试组 5: 自定义标题测试 ━━━');

test('自定义行字段标题', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+", "产品名称"], ["f2+"], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '自定义行字段标题');
});

test('自定义数据字段标题', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\")", "销售总额"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '自定义数据字段标题');
});

test('多数据字段自定义标题', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\"),count()", "销售额,订单数"], 1);
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '多数据字段自定义标题');
});

// ==================== 测试组6: 边界情况 ====================
console.log('\n━━━ 测试组 6: 边界情况测试 ━━━');

test('边界 - 空数据（仅表头）', function() {
    const data = [["产品", "地区", "销售额"]];
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '应返回至少表头');
    printResult(result, '空数据');
});

test('边界 - 单行数据', function() {
    const data = [["产品", "地区", "销售额"], ["A", "北京", 100]];
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '应能处理单行数据');
    printResult(result, '单行数据');
});

test('边界 - 重复数据聚合', function() {
    const data = [["产品", "地区", "销售额"], ["A", "北京", 100], ["A", "北京", 200], ["A", "北京", 300]];
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '应能处理重复数据');
    printResult(result, '重复数据聚合');
});

test('边界 - null/undefined 值', function() {
    const data = [["产品", "地区", "销售额"], ["A", "北京", 100], ["A", null, 200], [null, "上海", 300]];
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
    assertTrue(result.length > 0, '应能处理null值');
    printResult(result, 'null值处理');
});

// ==================== 测试组7: 选项测试 ====================
console.log('\n━━━ 测试组 7: 选项测试 ━━━');

test('选项 - cornerTitle 角落标题', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { cornerTitle: "销售分析" });
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, 'cornerTitle角落标题');
});

test('选项 - grandTotal 总计行', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { grandTotal: { row: true, label: "总计" } });
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '总计行');
});

test('选项 - grandTotal 总计列', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { grandTotal: { col: true, label: "总计" } });
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '总计列');
});

test('选项 - 行列总计', function() {
    const data = createSimpleData();
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { grandTotal: { row: true, col: true, label: "总计" } });
    assertTrue(result.length > 0, '结果不应为空');
    printResult(result, '行列总计');
});

// ==================== 测试组8: 性能测试 ====================
console.log('\n━━━ 测试组 8: 性能测试 ━━━');

test('性能 - 1000行数据', function() {
    const data = [["产品", "地区", "销售额"]];
    const products = ["A", "B", "C", "D", "E"];
    const regions = ["北京", "上海", "广州", "深圳", "杭州"];
    for (let i = 0; i < 1000; i++) {
        data.push([products[i % products.length], regions[i % regions.length], Math.floor(Math.random() * 10000)]);
    }
    const startTime = Date.now();
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
    const duration = Date.now() - startTime;
    assertTrue(result.length > 0, '应能处理大数据量');
    console.log(`     处理时间: ${duration}ms, 结果行数: ${result.length}`);
});

test('性能 - 5000行数据', function() {
    const data = [["产品", "地区", "销售额"]];
    const products = ["A", "B", "C", "D", "E"];
    const regions = ["北京", "上海", "广州", "深圳", "杭州"];
    for (let i = 0; i < 5000; i++) {
        data.push([products[i % products.length], regions[i % regions.length], Math.floor(Math.random() * 10000)]);
    }
    const startTime = Date.now();
    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
    const duration = Date.now() - startTime;
    assertTrue(result.length > 0, '应能处理大数据量');
    console.log(`     处理时间: ${duration}ms, 结果行数: ${result.length}`);
});

// ==================== 测试汇总 ====================
console.log('\n╔══════════════════════════════════════════════════╗');
console.log('║   测试结果汇总                                   ║');
console.log('╠══════════════════════════════════════════════════╣');
console.log(`║   总计: ${testCount} 个测试`);
console.log(`║   通过: ${passCount} 个 ✅`);
console.log(`║   失败: ${failCount} 个 ${failCount > 0 ? '❌' : ''}`);
console.log(`║   通过率: ${(passCount / testCount * 100).toFixed(1)}%`);
console.log('╚══════════════════════════════════════════════════╝');

if (failCount > 0) {
    process.exit(1);
}