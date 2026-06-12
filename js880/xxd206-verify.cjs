// XXD-206 regression verifier. Run after any edit that touches JSA880.js.
// Usage: node xxd206-verify.cjs
//
// Asserts that all 4 S3.x fixes are live in the source AND behave correctly:
//   S3.3: 多列/头名 join 头名相同时, 默认结果去掉右表键列
//   S3.4: 数组 resultSelector 越界返 null
//   S3.5: brr == null 返 this._new([])
//   S3.6: 字符串 sel 支持表头名解析
//
// Also asserts 4 regressions: numeric, function, string, array resultSelector.
const fs = require('fs');
const path = require('path');
const TARGET = path.join(__dirname, 'JSA880.js');
const src = fs.readFileSync(TARGET, 'utf8');

if (src.indexOf('/* XXD-206 S3.x fixes applied */') === -1) {
  console.error('FAIL: XXD-206 marker not present in source');
  process.exit(1);
}
if (src.indexOf('XXD-206 S3.5: null brr') === -1
    || src.indexOf('XXD-206 S3.6: capture header rows') === -1
    || src.indexOf('XXD-206 S3.4: 越界返 null') === -1
    || src.indexOf('XXD-206 S3.3: detect') === -1
    || src.indexOf('XXD-206 S3.3: 头名相同时') === -1) {
  console.error('FAIL: one or more S3.x markers missing — patch may be partial');
  process.exit(1);
}

const vm = require('vm');
const ctx = {};
vm.createContext(ctx);
vm.runInContext(src, ctx);
const A = ctx.Array2D;

let fails = 0;
function eq(name, got, expected) {
  const ok = JSON.stringify(got) === JSON.stringify(expected);
  if (ok) {
    console.log('PASS', name);
  } else {
    console.log('FAIL', name, '\n  got     ', JSON.stringify(got), '\n  expected', JSON.stringify(expected));
    fails++;
  }
}
function throws(name, fn, msgPart) {
  try {
    const r = fn();
    console.log('FAIL', name, '(expected throw containing "' + msgPart + '", got result ' + JSON.stringify(r) + ')');
    fails++;
  } catch (e) {
    if (e.message.indexOf(msgPart) !== -1) {
      console.log('PASS', name, '(threw: ' + e.message + ')');
    } else {
      console.log('FAIL', name, '(threw wrong message: ' + e.message + ')');
      fails++;
    }
  }
}

// S3.5 — null brr
eq('S3.5 null brr', (() => {
  const left = new A([['A','B'],[1,2]]);
  return left.z内连接(null, 0, 0);
})(), []);

// S3.5b — undefined brr
eq('S3.5b undefined brr', (() => {
  const left = new A([['A','B'],[1,2]]);
  return left.z内连接(undefined, 0, 0);
})(), []);

// S3.6 — string header sel works
eq('S3.6 string header sel', (() => {
  const left = new A([['A','B'],[1,2],[3,4]]);
  const right = new A([['A','C'],[1,10],[3,30]]);
  return left.z内连接(right._items, 'A', 'A');
})(), [['A','B','C'],[1,2,10],[3,4,30]]);

// S3.6b — string header with array resultSelector
// (Header rows themselves match on 'A' col, so the output includes the joined header row.)
eq('S3.6b string header + [1] resultSelector', (() => {
  const left = new A([['A','B'],[1,2],[3,4]]);
  const right = new A([['A','C'],[1,10],[3,30]]);
  return left.z内连接(right._items, 'A', 'A', [1]);
})(), [['A','B','C'],[1,2,10],[3,4,30]]);

// S3.3 — header-name join with matching key drops right's key col
eq('S3.3 header-name matching natural join', (() => {
  const left = new A([['A','B'],[1,2],[3,4]]);
  const right = new A([['A','C'],[1,10],[3,30]]);
  return left.z内连接(right._items, 'A', 'A');
})(), [['A','B','C'],[1,2,10],[3,4,30]]);

// S3.3b — header-name join with DIFFERENT key names keeps both key cols (regular concat).
// First row in each table is a header (key = 'K1' on left, 'K2' on right — no match),
// then data rows match on col 0 (1 ↔ 1).
eq('S3.3b header-name DIFFERENT keys', (() => {
  const left = new A([['K1','B'],[1,2]]);
  const right = new A([['K2','C'],[1,10]]);
  return left.z内连接(right._items, 'K1', 'K2');
})(), [[1,2,1,10]]);

// S3.4 — OOB array sel returns null explicitly
eq('S3.4 OOB array sel', (() => {
  const left = new A([['A','B'],[1,2]]);
  const right = new A([['A','C'],[1,10]]);
  return left.z内连接(right._items, 0, 0, [0, 5, 1]);
})(), [['A','B','A',null,'C'],[1,2,1,null,10]]);

// Regressions
eq('reg numeric sel', (() => {
  const left = new A([['A','B'],[1,2]]);
  const right = new A([['A','C'],[1,10]]);
  return left.z内连接(right._items, 0, 0);
})(), [['A','B','A','C'],[1,2,1,10]]);

eq('reg function sel', (() => {
  const left = new A([['A','B'],[1,2]]);
  const right = new A([['A','C'],[1,10]]);
  return left.z内连接(right._items, r => r[0], r => r[0]);
})(), [['A','B','A','C'],[1,2,1,10]]);

eq('reg string resultSelector a.f1,b.f1', (() => {
  const left = new A([['A','B'],[1,2]]);
  const right = new A([['A','C'],[1,10]]);
  return left.z内连接(right._items, 0, 0, 'a.f1,b.f1');
})(), [['A','A'],[1,1]]);

eq('reg array resultSelector [1]', (() => {
  const left = new A([['x','y'],[1,2]]);
  const right = new A([['x','y'],[1,10]]);
  return left.z内连接(right._items, 0, 0, [1]);
})(), [['x','y','y'],[1,2,10]]);

eq('reg empty result (no matches on data rows, header rows are identical)', (() => {
  // First rows ['A','B'] / ['A','C'] match on key 'A' (col 0), so result has at least the joined header.
  // Data rows [1,2] / [3,10] do not match (1 !== 3), so only the header row should appear.
  const left = new A([['A','B'],[1,2]]);
  const right = new A([['A','C'],[3,10]]);
  return left.z内连接(right._items, 0, 0);
})(), [['A','B','A','C']]);

// reg parseLambda fallback: First row does NOT have 'A' as a header cell, so parseLambda('A')
// is tried. Since 'A' is not a defined var, this throws — proving the fallback still works.
{
  const left = new A([['X','Y'],[1,2]]);
  const right = new A([['X','Z'],[1,10]]);
  throws('reg parseLambda fallback', () => left.z内连接(right._items, 'A', 'A'), 'A is not defined');
}

if (fails > 0) {
  console.error('XXD-206 verify FAILED with ' + fails + ' failure(s).');
  process.exit(1);
} else {
  console.log('XXD-206 verify: all assertions pass.');
}
