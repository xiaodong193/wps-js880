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

// === Smoke test ===
test('harness 自身能跑', function() {
    assertEqual(1 + 1, 2);
});

if (require.main === module) {
    runTests();
}

module.exports = { test: test, runTests: runTests, assertEqual: assertEqual,
                   JSA: JSA, Array2D: Array2D, $$: $$ };
