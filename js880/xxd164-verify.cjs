// XXD-164 / XXD-165 regression test:
//   JSA should expose 15 常用工具 z* aliases (z是否为空 ... z转小写).
//   Run from the js880 dir:
//     node xxd164-verify.cjs
var fs = require('fs');
var path = '/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880/JSA880.js';
eval(fs.readFileSync(path, 'utf8'));

var failures = 0;
function assertEq(name, actual, expected) {
  if (actual === expected) { console.log('  PASS', name, '=>', actual); }
  else { console.log('  FAIL', name, 'expected', expected, 'got', actual); failures++; }
}
function assertType(name, val, t) {
  if (typeof val === t) { console.log('  PASS', name, 'typeof=' + t); }
  else { console.log('  FAIL', name, 'expected typeof ' + t + ' got ' + typeof val); failures++; }
}

var names = [
  'z是否为空','z非空','z默认值','z当前时间戳','z当前日期',
  'z日期差','z包含','z开始于','z结束于','z去空白',
  'z分割','z连接','z替换全部','z转大写','z转小写'
];
console.log('XXD-164 / XXD-165 reproducer (issue body command):');
names.forEach(function(n){
  assertType(n + ' typeof function', JSA[n], 'function');
});

console.log('--- functional spot checks:');
assertEq('z是否为空(null)',       JSA.z是否为空(null),       true);
assertEq('z是否为空("")',          JSA.z是否为空(''),          true);
assertEq('z是否为空([])',          JSA.z是否为空([]),          true);
assertEq('z是否为空({})',          JSA.z是否为空({}),          false);
assertEq('z非空("x")',             JSA.z非空('x'),             true);
assertEq('z非空("")',              JSA.z非空(''),              false);
assertEq('z默认值(null,9)',        JSA.z默认值(null, 9),       9);
assertEq('z默认值(0,9)  (zero kept)', JSA.z默认值(0, 9),       0);
assertEq('z当前时间戳() > 0',      JSA.z当前时间戳() > 0,      true);
assertEq('z当前日期() is Date',    JSA.z当前日期() instanceof Date, true);
assertEq('z日期差 forward 10d',    JSA.z日期差('2026-01-01', '2026-01-11'), 10);
assertEq('z日期差 reverse -10d',   JSA.z日期差('2026-01-11', '2026-01-01'), -10);
assertEq('z包含("abc","b")',       JSA.z包含('abc', 'b'),      true);
assertEq('z开始于("abc","a")',     JSA.z开始于('abc', 'a'),    true);
assertEq('z结束于("abc","c")',     JSA.z结束于('abc', 'c'),    true);
assertEq('z去空白("  x  ")="x"',   JSA.z去空白('  x  '),       'x');
assertEq('z分割("a,b,c",",")',     JSON.stringify(JSA.z分割('a,b,c', ',')), '["a","b","c"]');
assertEq('z连接(["a","b"],"-")',   JSA.z连接(['a','b'], '-'),  'a-b');
assertEq('z替换全部("a-b-a","-","/")', JSA.z替换全部('a-b-a','-','/'), 'a/b/a');
assertEq('z转大写("abc")="ABC"',   JSA.z转大写('abc'),         'ABC');
assertEq('z转小写("ABC")="abc"',   JSA.z转小写('ABC'),         'abc');

// English aliases also resolve
assertEq('isEmpty === z是否为空',   JSA.isEmpty, JSA.z是否为空);
assertEq('replaceAll === z替换全部',JSA.replaceAll, JSA.z替换全部);
assertEq('startsWith === z开始于',  JSA.startsWith, JSA.z开始于);
assertEq('endsWith === z结束于',    JSA.endsWith, JSA.z结束于);
assertEq('trim === z去空白',        JSA.trim, JSA.z去空白);
assertEq('toUpperCase === z转大写', JSA.toUpperCase, JSA.z转大写);
assertEq('toLowerCase === z转小写', JSA.toLowerCase, JSA.z转小写);

// Pre-existing globals must NOT be clobbered
assertEq('JSA.now() still Date',   JSA.now() instanceof Date, true);
assertEq('JSA.today() still string', typeof JSA.today(), 'string');

console.log('---');
console.log('XXD-164 result: ' + (failures === 0 ? 'PASS' : (failures + ' FAILURES')));
process.exit(failures === 0 ? 0 : 1);
