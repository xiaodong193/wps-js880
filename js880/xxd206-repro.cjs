// XXD-206 repro: 4 boundary bugs in z内连接
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, 'JSA880.js'), 'utf8');
const vm = require('vm');
const ctx = {};
vm.createContext(ctx);
vm.runInContext(src, ctx);
const A = ctx.Array2D;

function show(name, fn) {
  try {
    const r = fn();
    console.log(name, '=>', JSON.stringify(r));
  } catch (e) {
    console.log(name, 'THREW:', e.message);
  }
}

// S3.3 — multi-col join returns duplicated B in result
show('S3.3 multi-col default concat', () => {
  const left = new A([['A', 'B'], [1, 2]]);
  const right = new A([['B', 'C'], [10, 20]]);
  return left.z内连接(right._items, 0, 0);
});

// S3.3b — same shape with header rows: user says 'B' duplicated in output
show('S3.3b multi-col with header', () => {
  const left = new A([['A', 'B'], [1, 2], [3, 4]]);
  const right = new A([['B', 'C'], [2, 20], [4, 40]]);
  return left.z内连接(right._items, 1, 0);
});

// S3.4 — right cols out of bounds (array resultSelector referencing idx 5 in 2-col table)
show('S3.4 out-of-bounds array sel', () => {
  const left = new A([['A', 'B'], [1, 2]]);
  const right = new A([['B', 'C'], [2, 20]]);
  return left.z内连接(right._items, 1, 0, [0, 5, 1]);
});

// S3.5 — null brr
show('S3.5 null brr', () => {
  const left = new A([['A', 'B'], [1, 2]]);
  return left.z内连接(null, 0, 0);
});

// S3.6 — string header selectors
show('S3.6 string header selectors', () => {
  const left = new A([['A', 'B'], [1, 2], [3, 4]]);
  const right = new A([['A', 'C'], [1, 10], [3, 30]]);
  return left.z内连接(right._items, 'A', 'A');
});

show('S3.6b string header + array resultSelector', () => {
  const left = new A([['A', 'B'], [1, 2], [3, 4]]);
  const right = new A([['A', 'C'], [1, 10], [3, 30]]);
  return left.z内连接(right._items, 'A', 'A', [1]);
});

// regression: numeric sel still works
show('regress numeric sel', () => {
  const left = new A([['A', 'B'], [1, 2]]);
  const right = new A([['A', 'C'], [1, 10]]);
  return left.z内连接(right._items, 0, 0);
});

// regression: function sel
show('regress function sel', () => {
  const left = new A([['A', 'B'], [1, 2]]);
  const right = new A([['A', 'C'], [1, 10]]);
  return left.z内连接(right._items, r => r[0], r => r[0]);
});

// regression: string 'a.f1,b.f2' resultSelector
show('regress string resultSelector a.f1,b.f1', () => {
  const left = new A([['A', 'B'], [1, 2]]);
  const right = new A([['A', 'C'], [1, 10]]);
  return left.z内连接(right._items, 0, 0, 'a.f1,b.f1');
});
