/**
 * XXD-51: __KJ_ARGS__ 解析失败时静默走默认路径,fn 标记仍被剥 — 改为严格 JSON
 *
 * 3 个验收用例(来自 issue 验收段):
 *   1. 拼错 `__KJ_ARGS__={rowFields}` (无冒号) → 抛 `#K_ERR: pos=0, FN, msg="...__KJ_ARGS__ 解析失败: ..."` 而不是静默成功
 *   2. 正确 `__KJ_ARGS__={"rowFields":"f3,f2"}` → 走原路径,fn 正常执行(返回 x+1)
 *   3. `__KJ_ARGS__={}` 空对象 → __kjExtracted=[] 且 fn 标记**不剥**(下游按 fn 解析错兜底,而非悄悄剥掉标记)
 *
 * 运行: node test/test_xxd51_kj_args_strict.js
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

function assertNotMatch(actual, pattern, msg) {
    var s = (actual === null || actual === undefined) ? String(actual) : String(actual);
    if (!pattern.test(s)) {
        passed++;
        console.log('  ✅ ' + msg);
    } else {
        failed++;
        failures.push(msg + ' (实际: ' + s + ', 不应匹配: ' + pattern + ')');
        console.log('  ❌ ' + msg + '\n     实际: ' + s);
    }
}

console.log('🚀 XXD-51 — __KJ_ARGS__ 严格 JSON 解析验收');
console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

// ==================== 模拟 WPS / JSA880 环境 ====================
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
var path = require('path');
var fs = require('fs');

var jsaSrc = fs.readFileSync(path.join(__dirname, '..', 'JSA880.js'), 'utf8');
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

// ==================== 用例 1: 拼错 __KJ_ARGS__={rowFields} 抛 K_ERR ====================
console.log('用例 1: JSA.k("__KJ_ARGS__={rowFields} (x)=>x+1", 1) → #K_ERR FN "__KJ_ARGS__ 解析失败"');
//   - 严格 JSON.parse("{rowFields}") 失败(没有冒号)
//   - 宽松解析:protected = "rowFields" → __kjEntries=["rowFields"] → 无 : 跳过 → __kj={}
//   - 期望:触发新的 throw,__kCode=PARSE,外层 JSA.k 兜成 #K_ERR 字符串
var ret1 = JSA.k('__KJ_ARGS__={rowFields} (x)=>x+1', 1);
assertMatch(
    ret1,
    /^#K_ERR: pos=0, FN, msg=".*__KJ_ARGS__ 解析失败: .*rowFields.*"/,
    '拼错的 __KJ_ARGS__={rowFields} 抛 #K_ERR 并包含 "__KJ_ARGS__ 解析失败"'
);

// ==================== 用例 2: 正确 __KJ_ARGS__={"rowFields":"f3,f2"} 走原路径 ====================
console.log('\n用例 2: JSA.k("__KJ_ARGS__={\\"rowFields\\":\\"f3,f2\\"} (x)=>x+1", 1) → 2');
//   - 严格 JSON.parse 成功
//   - __kjExtracted=["f3,f2"]
//   - 标记被剥,fn 变成 "(x)=>x+1",传入 1 → 返回 2
var ret2 = JSA.k('__KJ_ARGS__={"rowFields":"f3,f2"} (x)=>x+1', 1);
assertEq(ret2, 2, '正确 __KJ_ARGS__ JSON 走原路径,返回 x+1');

// ==================== 用例 3: __KJ_ARGS__={} 空对象,fn 标记不剥 ====================
console.log('\n用例 3: JSA.k("__KJ_ARGS__={} (x)=>x+1", 1) → 不抛 __KJ_ARGS__ 解析失败');
//   - 严格 JSON.parse("{}") 成功(__kjStrictOk=true,不进入 throw 分支)
//   - __kjExtracted=[] → 守卫生效,fn 标记**不剥** → fn 仍是 "__KJ_ARGS__={} (x)=>x+1"
//   - 下游 z解析函数表达式 看到 __KJ_ARGS__={} 不是合法 fn 语法 → 抛 SyntaxError
//   - 外层 JSA.k 兜成 #K_ERR PARSE,**不应**包含 "__KJ_ARGS__ 解析失败"(那是我们的提示,不是 fn 解析错)
var ret3 = JSA.k('__KJ_ARGS__={} (x)=>x+1', 1);
assertNotMatch(
    ret3,
    /__KJ_ARGS__ 解析失败/,
    '__KJ_ARGS__={} 不应触发 __KJ_ARGS__ 解析失败提示(改走下游 fn 解析错路径)'
);
assertMatch(
    ret3,
    /^#K_ERR: pos=0, FN, msg=".*"/,
    '__KJ_ARGS__={} 仍走 #K_ERR FN 路径(下游 fn 解析错兜底)'
);

// ==================== 用例 4 (回归): 宽松解析能拿到键时,仍走原路径(不抛) ====================
console.log('\n用例 4 (回归): 宽松解析能拿到键时不抛(原 v4.0.27 行为保留)');
//   - 输入:__KJ_ARGS__={rowFields:"f3,f2"} (缺字段名引号,缺右外层引号)
//   - 严格 JSON.parse 失败;宽松解析成功:__kj = {rowFields: '"f3,f2"'} (有键,非空)
//   - 期望:不抛 __KJ_ARGS__ 解析失败;走后续提取路径
var ret4 = JSA.k('__KJ_ARGS__={rowFields:"f3,f2"} (x)=>x+1', 1);
assertNotMatch(
    ret4,
    /^#K_ERR:.*__KJ_ARGS__ 解析失败/,
    '宽松解析能拿到键时不应抛 __KJ_ARGS__ 解析失败'
);
assertEq(ret4, 2, '宽松解析成功时,(x)=>x+1 正常返回 2');

// ==================== 用例 5 (回归): 没 __KJ_ARGS__ 标记的老公式不受影响 ====================
console.log('\n用例 5 (回归): 无 __KJ_ARGS__ 标记的公式不受影响');
assertEq(JSA.k('x => x+1', 1), 2, '(x)=>x+1 仍正常工作');

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
    console.log('\n🎉 全部验收通过 (XXD-51 修复有效)');
    process.exit(0);
}
