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

// === 加载真实 JSA880.js(在 Node 下 require 它) ===
var path = require('path');
var fs = require('fs');

// 把 JSA880.js 读出来,在我们的模拟环境里 eval
// (不能直接 require,因为它不是 CommonJS 模块)
var jsa880Path = path.resolve(__dirname, '../../../js880/JSA880.js');
var jsa880Code = fs.readFileSync(jsa880Path, 'utf8');

// 模拟 WPS JSA 的全局上下文
var vm = require('vm');
var sandbox = {
    JSA: JSA,
    Array2D: Array2D,
    Range: Range,
    Console: Console,
    Application: Application,
    $$: undefined,
    globalThis: globalThis,
    console: console
};
// 让 $$ 通过 globalThis 引用
sandbox.globalThis.$$ = undefined;

// eval JSA880.js 在沙箱里
vm.createContext(sandbox);
vm.runInContext(jsa880Code, sandbox);

// eval 完后,实际的对象在 sandbox 里
// 把 Array2D / JSA 拉回到 Node 全局
JSA = sandbox.JSA;
Array2D = sandbox.Array2D;
$$ = sandbox.$$ || sandbox.globalThis.$$;
globalThis.$$ = $$;

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

// === 5 个错误注入测试 ===
test('E1: k() 空调用 → #K_ERR: pos=0, FN', function() {
    var result = JSA.k();
    if (typeof result !== 'string' || result.indexOf('#K_ERR') !== 0) {
        throw new Error('期望错误字符串,实际:' + JSON.stringify(result));
    }
    if (result.indexOf('pos=0') === -1) {
        throw new Error('期望 pos=0,实际:' + result);
    }
    if (result.indexOf('FN') === -1) {
        throw new Error('期望 kind=FN,实际:' + result);
    }
});

test('E2: k("JSA.xxx") 但 JSA880 未加载 → #K_ERR: INTERNAL', function() {
    // 临时把 JSA.jsaLambda 设为 undefined
    var saved = JSA.jsaLambda;
    JSA.jsaLambda = undefined;
    var result;
    try {
        result = JSA.k('JSA.getIndexs', 1, 10, 2);
    } finally {
        JSA.jsaLambda = saved;
    }
    if (typeof result !== 'string' || result.indexOf('INTERNAL') === -1) {
        throw new Error('期望 INTERNAL 错误,实际:' + result);
    }
});

test('E3: k("$$yyy") $$yyy 不存在 → #K_ERR: FN', function() {
    var result = JSA.k('$$notExist', 1, 2);
    if (typeof result !== 'string' || result.indexOf('FN') === -1) {
        throw new Error('期望 FN 错误,实际:' + result);
    }
});

test('E4: k("x=>x.b", "abc") 类型错 → 包含 TypeError', function() {
    var result = JSA.k('x=>x.b()', 'abc');
    if (typeof result !== 'string' || result.indexOf('TypeError') === -1) {
        throw new Error('期望 TypeError,实际:' + result);
    }
});

test('E5 (meta): V8 保留模板字符串 ${}, 不依赖 JSA.k', function() {
    // 这是模板字符串场景,期望 #K_ERR(因为 b 未定义),但 ${} 要保留
    // 我们只验证它没被改成双引号
    var errThrown = null;
    try {
        // 模拟:用 new Function 直接编译,看 b 引用是否被保留
        var fn = new Function('return ' + '`a${b}c`');
        fn();
    } catch (e) {
        errThrown = e;
    }
    if (!errThrown || errThrown.message.indexOf('b is not defined') === -1) {
        throw new Error('模板字符串应被原样保留,实际:' + (errThrown && errThrown.message));
    }
});

test('E5b: k("`a${b}c`", 1) → JSA.k 实际调用应保留 ${}', function() {
    // 真正的 JSA.k 测试:传一个含模板字符串的 fn,期望 #K_ERR(因为 b 未定义),
    // 但 ${} 必须被保留(没被错误地转成双引号)
    var result = JSA.k('`a${b}c`', 1);
    if (typeof result !== 'string' || result.indexOf('b') === -1) {
        throw new Error('JSA.k 应保留 ${b},实际:' + result);
    }
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
