/**
 * XXD-162 回归测试: Array2D.z求和 字符串 sel 抛 ReferenceError (A is not defined)
 *
 * 修复前:
 *   new Array2D([['A','B'],[1,2]]).z求和('A') → THROW "A is not defined"
 *
 * 修复后:
 *   new Array2D([['A','B'],[1,2]]).z求和('A') → 1 (按首行表头名解析)
 *   new Array2D([['A','B'],[1,2]]).z求和('foo') → THROW TypeError 友好错误
 *
 * 解析流程:
 *   1) 已知 lambda 语法 (f1/$0/箭头函数/方括号) → 走 parseLambda
 *   2) 否则尝试作为 header 名 (匹配首行表头)
 *   3) 都不匹配抛 TypeError (而非 ReferenceError)
 *
 * 运行方式:
 *   node test_xxd162_zqiuhe_header.js
 *
 * 依赖：纯 ES5（不依赖 WPS 任何对象），所以可以在 Node 中跑。
 */

var fs = require('fs');
var path = require('path');

// 加载 JSA880.js
var JSA_PATH = path.join(__dirname, '..', 'JSA880.js');
eval(fs.readFileSync(JSA_PATH, 'utf8'));

var passed = 0;
var failed = 0;

function assertEq(name, actual, expected) {
    if (actual === expected) {
        passed++;
        console.log('  ✓ ' + name);
    } else {
        failed++;
        console.log('  ✗ ' + name + '  expected=' + JSON.stringify(expected) + '  actual=' + JSON.stringify(actual));
    }
}

function assertThrows(name, fn, errType) {
    try {
        var r = fn();
        failed++;
        console.log('  ✗ ' + name + '  expected throw, got ' + JSON.stringify(r));
    } catch (e) {
        if (errType && !(e instanceof errType)) {
            failed++;
            console.log('  ✗ ' + name + '  expected ' + errType.name + ', got ' + e.constructor.name + ': ' + e.message);
        } else {
            passed++;
            console.log('  ✓ ' + name + '  (threw: ' + e.constructor.name + ': ' + e.message + ')');
        }
    }
}

console.log('=== XXD-162: z求和 header-name resolution ===\n');

console.log('[1] Header name resolution (the fix)');
assertEq('z求和("A") on [[A,B],[1,2]]',
    new Array2D([['A','B'],[1,2]]).z求和('A'), 1);
assertEq('z求和("B") on [[A,B],[1,2],[3,4]]',
    new Array2D([['A','B'],[1,2],[3,4]]).z求和('B'), 6);
assertEq('z求和("age") on real header',
    new Array2D([['name','age','score'],['alice',25,90],['bob',30,85]]).z求和('age'), 55);
assertEq('z求和("score") on real header',
    new Array2D([['name','age','score'],['alice',25,90],['bob',30,85]]).z求和('score'), 175);

console.log('\n[2] Existing lambda syntax preserved');
assertEq('z求和("f1")',
    new Array2D([['A','B'],[1,2],[3,4]]).z求和('f1'), 4);
assertEq('z求和("f2")',
    new Array2D([['A','B'],[1,2],[3,4]]).z求和('f2'), 6);
assertEq('z求和("f1*f2") expression',
    new Array2D([['A','B'],[1,2],[3,4]]).z求和('f1*f2'), 14);
assertEq('z求和("row=>row[0]") arrow',
    new Array2D([['A','B'],[1,2],[3,4]]).z求和('row=>row[0]'), 4);
assertEq('z求和("$0") index syntax',
    new Array2D([['A','B'],[1,2],[3,4]]).z求和('$0'), 4);
assertEq('z求和("[f1,f2]") array bracket',
    new Array2D([['A','B'],[1,2],[3,4]]).z求和('[f1,f2]'), 46);
assertEq('z求和() default flatten',
    new Array2D([['A','B'],[1,2]]).z求和(), 3);
assertEq('z求和(function) callback',
    new Array2D([['A','B'],[1,2]]).z求和(function(r){return r[0];}), 1);
assertEq('z求和(1) numeric 1-based',
    new Array2D([['A','B'],[1,2],[3,4]]).z求和(1), 4);
assertEq('z求和(2) numeric 1-based',
    new Array2D([['A','B'],[1,2],[3,4]]).z求和(2), 6);
assertEq('sum alias ("A")',
    new Array2D([['A','B'],[1,2]]).sum('A'), 1);

console.log('\n[3] Friendly TypeError (not ReferenceError) for unknown strings');
assertThrows('z求和("foo") on header data',
    function() { new Array2D([['A','B'],[1,2]]).z求和('foo'); }, TypeError);
assertThrows('z求和("xyz") on no-header data',
    function() { new Array2D([[1,2],[3,4]]).z求和('xyz'); }, TypeError);
assertThrows('z求和("missing") on non-matching header',
    function() { new Array2D([['foo','bar'],[1,2]]).z求和('baz'); }, TypeError);

console.log('\n[4] Numeric boundary preserved (XXD-161 fix not regressed)');
assertThrows('z求和(0) out of range',
    function() { new Array2D([['A','B'],[1,2]]).z求和(0); }, RangeError);
assertThrows('z求和(3) out of range',
    function() { new Array2D([['A','B'],[1,2]]).z求和(3); }, RangeError);

console.log('\n[5] Issue repro from XXD-162 description');
// 复现 issue 描述中的命令:
//   node -e "var fs=require('fs');eval(fs.readFileSync('JSA880.js','utf8'));try{console.log(new Array2D([['A','B'],[1,2]]).z求和('A'))}catch(e){console.log('THROW',e.message);}"
// 期望: 1 (而非 "THROW A is not defined")
var reproResult;
try {
    reproResult = new Array2D([['A','B'],[1,2]]).z求和('A');
} catch(e) {
    reproResult = 'THROW: ' + e.message;
}
assertEq('Issue XXD-162 exact repro', reproResult, 1);

console.log('\n=== Summary ===');
console.log('Passed: ' + passed);
console.log('Failed: ' + failed);
process.exit(failed > 0 ? 1 : 0);
