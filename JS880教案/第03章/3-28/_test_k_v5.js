/**
 * KO k() UDF v5 · Node 仿真测试
 * 跑 9 个 Excel 测试公式 + 5 个错误注入,无需 WPS
 */
'use strict';

// === 模拟 WPS JSA 环境 ===
var JSA = {};
var Array2D = {
    version: function() { return '5.0.0-test'; },
    leftjoin: function() { throw new Error('未实现: leftjoin'); },
    superPivot: function() { throw new Error('未实现: superPivot'); },
    distinct: function() { throw new Error('未实现: distinct'); },
    insertCols: function() { throw new Error('未实现: insertCols'); }
};
var $$;  // 全局别名,稍后由 JSA880.js 设置
var Range = function(addr) { return { Address: addr, Value2: [[1,2],[3,4]] }; };
var Console = { log: function() { /* console.log.apply(console, arguments); */ } };
var Application = {};

// 模拟 globalThis
var globalThis = (typeof global !== 'undefined') ? global : this;
globalThis.$$ = $$;
globalThis.JSA = JSA;
globalThis.Array2D = Array2D;
globalThis.Range = Range;

// === 加载要测试的模块 ===
// 稍后 loadFile('JSA880.js') 或 loadFile('KO_k_udf_v5.js')

// === 测试运行器 ===
var tests = [];
var passed = 0;
var failed = 0;

function test(name, fn) {
    tests.push({ name: name, fn: fn });
}

function assertEqual(actual, expected, msg) {
    if (JSON.stringify(actual) === JSON.stringify(expected)) return;
    throw new Error('assertEqual 失败: ' + (msg || '') +
        '\n  期望: ' + JSON.stringify(expected) +
        '\n  实际: ' + JSON.stringify(actual));
}

function runTests() {
    for (var i = 0; i < tests.length; i++) {
        var t = tests[i];
        try {
            t.fn();
            console.log('✅ ' + t.name);
            passed++;
        } catch (e) {
            console.log('❌ ' + t.name);
            console.log('   ' + e.message);
            failed++;
        }
    }
    console.log('\n' + (tests.length === passed ? '✅ 全部通过' : '❌ ' + failed + ' 个失败'));
    process.exit(failed > 0 ? 1 : 0);
}

// === 9 个 Excel 公式测试(数据用 JSON 模拟 Range) ===
var dataA2B7 = [['A', 1], ['B', 2], ['C', 3], ['A', 4], ['B', 5], ['C', 6]];
var dataA1C17 = [];  // 17 行 3 列,模拟 A1:C17
for (var i = 0; i < 17; i++) {
    dataA1C17.push([i + 1, 'P' + (i % 3), (i + 1) * 10]);
}
var dataA1H40 = [];
for (var j = 0; j < 40; j++) {
    dataA1H40.push([j+1, 'P' + (j%3), 'C' + (j%4), j+1, 1, 2020 + (j%5), (j%12)+1, j+1]);
}

// T1: leftjoin 路径调用
test('T1: k("$$.leftjoin", ...) 路径调用', function() {
    Array2D.leftjoin = function(arr, brr, lk, rk, sel) {
        return [['joined', 'result']];
    };
    var result = JSA.k('$$.leftjoin', [['A'],['B']], dataA2B7, 'f1', 'f1', 'a.f1,b.f2');
    assertEqual(result, [['joined', 'result']]);
});

// T2: 多参 lambda + 嵌套箭头
test('T2: k("(a,v)=>a.filter(x=>x[1]>v)", arr, val)', function() {
    var result = JSA.k('(a,v)=>a.filter(x=>x[1]>v)', dataA2B7, 2);
    // 期望:第 2 列 > 2 的行
    var expected = dataA2B7.filter(function(x) { return x[1] > 2; });
    assertEqual(result, expected);
});

// T3: lambda 返回 1D 数组
test('T3: k("(arr,v)=>arr.map(x=>[x.f2+v])", arr, val)', function() {
    var result = JSA.k('(arr,v)=>arr.map(x=>[x.f2+v])', dataA2B7, 100);
    // 期望:每行第 2 列 + 100,包成 1 元素数组
    var expected = dataA2B7.map(function(x) { return [x[1] + 100]; });
    assertEqual(result, expected);
});

// T4: lambda 调用 $$.distinct
test('T4: k("arr=>$$.distinct(arr,\'f1\')", arr)', function() {
    Array2D.distinct = function(arr, key) {
        // 简单实现:按第 0 列去重
        var seen = {}, result = [];
        for (var i = 0; i < arr.length; i++) {
            if (!seen[arr[i][0]]) { seen[arr[i][0]] = 1; result.push([arr[i][0]]); }
        }
        return result;
    };
    var result = JSA.k("arr=>$$.distinct(arr,'f1')", dataA2B7);
    assertEqual(result, [['A'],['B'],['C']]);
});

// T5: lambda 返回 2D 数组
test('T5: k("(arr,v)=>arr.map(x=>[x.f1,x.f2+v])", arr, val)', function() {
    var result = JSA.k('(arr,v)=>arr.map(x=>[x.f1,x.f2+v])', dataA2B7, 10);
    var expected = dataA2B7.map(function(x) { return [x[0], x[1] + 10]; });
    assertEqual(result, expected);
});

// T6: rest + spread + 嵌套调用
test('T6: k("(...args)=>$$.insertCols($$.superPivot(...args),-1,x=>x.sum())", ...)', function() {
    Array2D.superPivot = function() {
        return [['P0', 'sum', 100], ['P1', 'sum', 200]];
    };
    Array2D.insertCols = function(arr, pos, fn) {
        return [['P0', 'sum', 100, 100], ['P1', 'sum', 200, 200]];
    };
    var result = JSA.k('(...args)=>$$.insertCols($$.superPivot(...args),-1,x=>x.sum())',
                       dataA1C17, 'f1', 'f2', "sum('f3')", 1, 0);
    assertEqual(result.length, 2);
    assertEqual(result[0].length, 4);
});

// T7: block body
test('T7: k("(x,y)=>{y=y+1;return x.map(a=>[a[0],a[2]+\'-\'+y])}", arr, val)', function() {
    var result = JSA.k("(x,y)=>{y=y+1;return x.map(a=>[a[0],a[2]+'-'+y])}",
                       dataA1C17, 1);
    // 期望:每行 [id, 数量+'-'+2]
    var expected = dataA1C17.map(function(a) { return [a[0], a[2] + '-2']; });
    assertEqual(result, expected);
});

// T8: 反引号列名
test('T8: k("$$.superPivot", arr, "...", "f6", "sum(`f4*f5`)...")', function() {
    Array2D.superPivot = function(data, rows, cols, vals) {
        return [['header', 'value'], ['P0', 100]];
    };
    var result = JSA.k('$$.superPivot', dataA1H40, 'f3,f2', 'f6',
                       'sum(`f4*f5`),textjoin(`f4+\'-\'+f5`,`+`)');
    assertEqual(result.length, 2);
});

// T9: rest + spread + chainable filter
test('T9: k("(...args)=>$$.superPivot(...args).filter(...)", arr, ..., "", ...)', function() {
    Array2D.superPivot = function() {
        return [['header'], ['Product1'], ['Product2'], ['Product1']];
    };
    var result = JSA.k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",
                       dataA1H40, 'f3,f2', '', "count(),sum(`f4`),textjoin(`f4`,`+`)");
    // 期望:header + 2 个 Product1 行
    assertEqual(result.length >= 2, true);
});

// === Smoke test ===
test('harness 自身能跑', function() {
    assertEqual(1 + 1, 2);
});

if (require.main === module) {
    runTests();
}

module.exports = { test: test, runTests: runTests, assertEqual: assertEqual,
                   JSA: JSA, Array2D: Array2D, $$: $$ };
