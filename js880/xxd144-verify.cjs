// XXD-144 regression test: z最大值 / z最小值 on raw Array2D must skip header row
var fs = require('fs');
var path = '/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880/JSA880.js';
eval(fs.readFileSync(path, 'utf8'));

var failures = 0;
function assertEq(name, actual, expected) {
  if (actual === expected) { console.log('  PASS', name, '=>', actual); }
  else { console.log('  FAIL', name, 'expected', expected, 'got', actual); failures++; }
}

console.log('XXD-144 reproducer:');
assertEq('raw [[v],[1],[2],[3]] z最大值', new Array2D([['v'],[1],[2],[3]]).z最大值(), 3);
assertEq('raw [[v],[1],[2],[3]] z最小值', new Array2D([['v'],[1],[2],[3]]).z最小值(), 1);

console.log('Aliases:');
assertEq('prototype.max', new Array2D([['v'],[1],[2],[3]]).max(), 3);
assertEq('prototype.min', new Array2D([['v'],[1],[2],[3]]).min(), 1);
assertEq('Array2D.z最大值 static', Array2D.z最大值([['v'],[1],[2],[3]]), 3);
assertEq('Array2D.max static', Array2D.max([['v'],[1],[2],[3]]), 3);

console.log('With explicit _header (header is NOT in data — do not skip):');
var withHeader = new Array2D([[1],[2],[3]]);
withHeader._header = ['col1'];
assertEq('_header set, max [[1],[2],[3]]', withHeader.z最大值(), 3);
assertEq('_header set, min [[1],[2],[3]]', withHeader.z最小值(), 1);

console.log('Numeric "header" (still skipped — raw mode):');
assertEq('raw [[10],[1],[2],[3]] max', new Array2D([[10],[1],[2],[3]]).z最大值(), 3);
assertEq('raw [[10],[1],[2],[3]] min', new Array2D([[10],[1],[2],[3]]).z最小值(), 1);

console.log('Edge cases:');
assertEq('empty []', new Array2D([]).z最大值(), undefined);
assertEq('single row [[v]]', new Array2D([['v']]).z最大值(), undefined);

console.log('Multi-col raw:');
assertEq('raw [[a,b],[1,4],[2,5]] max', new Array2D([['a','b'],[1,4],[2,5]]).z最大值(), 5);

if (failures > 0) { console.log('\n*** ' + failures + ' FAILURES ***'); process.exit(1); }
console.log('\nAll passed.');
