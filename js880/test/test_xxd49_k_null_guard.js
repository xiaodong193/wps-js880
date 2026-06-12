/**
 * XXD-49: JSA.k 不应把 fn 主动返回的 null/undefined 兜底成 #K_ERR
 *
 * 5 个验收用例:
 *   1. JSA.k('x => null', 1)            → null
 *   2. JSA.k('x => undefined', 1)       → undefined
 *   3. JSA.k('x => -1', [1,2,3])        → -1
 *   4. JSA.k('x => throw new TypeError()', 1) → '#K_ERR: pos=0, TYPE, msg="..."'
 *   5. JSA.k() / JSA.k('') / JSA.k(null)   → '#K_ERR: pos=0, FN, msg="fn 不能为空..."'
 *
 * 运行: node test/test_xxd49_k_null_guard.js
 */

'use strict';

var passed = 0, failed = 0, failures = [];

function assertEq(actual, expected, msg) {
    var a = JSON.stringify(actual);
    var e = JSON.stringify(expected);
    if (a === e) {
        passed++;
        console.log('  ✅ ' + msg);
    } else {
        failed++;
        failures.push(msg + ' (实际: ' + a + ', 期望: ' + e + ')');
        console.log('  ❌ ' + msg + '\n     实际: ' + a + '\n     期望: ' + e);
    }
}

function assertMatch(actual, pattern, msg) {
    var s = (actual === null || actual === undefined) ? String(actual) : String(actual);
    if (pattern.test(s)) {
        passed++;
        console.log('  ✅ ' + msg);
    } else {
        failed++;
        failures.push(msg + ' (实际: ' + s + ', 期望匹配: ' + pattern + ')');
        console.log('  ❌ ' + msg + '\n     实际: ' + s);
    }
}

console.log('🚀 XXD-49 — JSA.k null/undefined 兜底修复验收');
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

// ==================== 模拟 WPS / JSA880 环境 ====================
// 只 mock JSA.k 真正访问的全局对象,保持测试聚焦

global.Console = {
    log: function () { /* swallow */ }
};
global.Application = {
    ActiveSheet: { Name: 'test' },
    Worksheets: function () { return { Name: 'x', Cells: {} }; },
    Range: function () { return { Value2: null }; },
    ScreenUpdating: true,
    Calculation: -4105
};

// ==================== 加载 JSA880.js ====================
// JSA880.js 顶层定义 function k/jsaLambda (WPS UDF shim),这是宿主层 entry
// 我们只关心 JSA.k 自身,顶层 k shim 会调用 JSA.k.apply,Node 环境没在宿主层
// 被调,但 JSA.k 已挂在 JSA 命名空间上,直接测 JSA.k 即可

var path = require('path');
var fs = require('fs');

var jsaSrc = fs.readFileSync(path.join(__dirname, '..', 'JSA880.js'), 'utf8');
// 在 global 作用域 eval,使顶层 var JSA / function k / 等都挂到 global
// (Node 直接 eval(...) 是 indirect eval 的话会进 global 域)

try {
    (0, eval)(jsaSrc);
} catch (e) {
    console.error('❌ 加载 JSA880.js 失败: ' + e.message);
    console.error(e.stack);
    process.exit(1);
}

if (typeof global.JSA === 'undefined' || typeof global.JSA.k !== 'function') {
    console.error('❌ JSA.k 未注册');
    process.exit(1);
}
var JSA = global.JSA;

console.log('  ℹ️  JSA880.js 已加载, JSA.k 已是函数');
console.log('');

// ==================== 用例 1: 主动返回 null ====================
console.log('用例 1: JSA.k("x => null", 1) → null');
assertEq(JSA.k('x => null', 1), null, '返回 null 而非 #K_ERR');

// ==================== 用例 2: 主动返回 undefined ====================
console.log('\n用例 2: JSA.k("x => undefined", 1) → undefined');
assertEq(JSA.k('x => undefined', 1), undefined, '返回 undefined 而非 #K_ERR');

// ==================== 用例 3: 主动返回 -1 ====================
console.log('\n用例 3: JSA.k("x => -1", [1,2,3]) → -1');
assertEq(JSA.k('x => -1', [1, 2, 3]), -1, '返回 -1');

// ==================== 用例 4: throw 仍然走 #K_ERR ====================
console.log('\n用例 4: JSA.k("x => throw new TypeError()", 1) → #K_ERR');
var ret4 = JSA.k('x => { throw new TypeError("boom"); return 1; }', 1);
assertMatch(
    ret4,
    /^#K_ERR: pos=0, TYPE, msg=".*boom.*"/,
    'TypeError 走 #K_ERR TYPE 格式'
);

// ==================== 用例 5: 空 fn 仍然走 #K_ERR FN ====================
console.log('\n用例 5: JSA.k() / JSA.k("") / JSA.k(null) → #K_ERR FN_EMPTY');
assertMatch(JSA.k(),         /^#K_ERR: pos=0, FN, msg="fn 不能为空.*"/, 'JSA.k() 返回 #K_ERR FN');
assertMatch(JSA.k(''),       /^#K_ERR: pos=0, FN, msg="fn 不能为空.*"/, 'JSA.k("") 返回 #K_ERR FN');
assertMatch(JSA.k(null),     /^#K_ERR: pos=0, FN, msg="fn 不能为空.*"/, 'JSA.k(null) 返回 #K_ERR FN');

// ==================== 额外回归: 0 / false / 空串 仍然合法透传 ====================
console.log('\n回归: 0 / false / "" 仍是合法结果 (issue XXD-44 注释承诺)');
assertEq(JSA.k('x => 0', 1), 0, 'JSA.k("x => 0", 1) === 0');
assertEq(JSA.k('x => false', 1), false, 'JSA.k("x => false", 1) === false');
assertEq(JSA.k('x => ""', 1), '', 'JSA.k("x => """, 1) === ""');

// ==================== 测试结果 ====================
console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log('  通过: ' + passed);
console.log('  失败: ' + failed);
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

if (failed > 0) {
    console.log('\n❌ 失败用例:');
    failures.forEach(function (f, i) { console.log('  ' + (i + 1) + '. ' + f); });
    process.exit(1);
} else {
    console.log('\n🎉 全部验收通过 (XXD-49 修复有效)');
    process.exit(0);
}
