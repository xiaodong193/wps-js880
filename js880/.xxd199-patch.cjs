// XXD-199: z添加列 — name-as-string, header-row-aware, always-append-at-end
// F3.1 z添加列('Z', fn)        add at end, header='Z', data=fn(row)        OK
// F3.2 z添加列(99)             TypeError (number not a valid name)
// F3.3 z添加列('C', [10,20,30]) append at end, header='C', data=[10,20,30]
// F3.4 z添加列(null)           TypeError (null is not a valid name)
// F3.5 z添加列('Z')            append at end, header='Z', data all null
// S6.5 z添加列(['X','Y'], fn)  TypeError (array is not a valid name)
//
// Contract: z添加列(name, data|fn?) where name must be a string. data may be
// - function  → fn(row) per data row
// - array     → array[i] per data row (length must equal data row count)
// - undefined → all nulls
// Header row (row 0) is updated: the new last cell = name.

const fs = require('fs');
const path = require('path');
const target = '/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880/JSA880.js';
const src = fs.readFileSync(target, 'utf8');

const OLD = `Array2D.prototype.z添加列 = function(col, index) {
    // XXD-190: computed-column signature — z添加列(colName, fn)
    // header row gets colName appended, data rows get fn(row) appended.
    if (typeof col === "string" && typeof index === "function") {
        var computed = [];
        for (var r = 0; r < this._items.length; r++) {
            var row = this._items[r].slice();
            var v;
            try { v = index(row); } catch (e) { v = null; }
            row.push(v);
            computed.push(row);
        }
        if (computed.length > 0) {
            var header = computed[0].slice();
            header[header.length - 1] = col;
            computed[0] = header;
        }
        return this._new(computed);
    }
    var result = [];
    var colIndex = index !== undefined ? index : this.z列数();
    for (var i = 0; i < this._items.length; i++) {
        var newRow = this._items[i].slice();
        newRow.splice(colIndex, 0, col[i] !== undefined ? col[i] : null);
        result.push(newRow);
    }
    return this._new(result);
};`;

const NEW = `Array2D.prototype.z添加列 = function(name, data) {
    // XXD-199: z添加列(name, data|fn?) — name is the column header (string, required)
    // header row (row 0) gets \`name\` appended; data column is appended at the end.
    //   data === function  → fn(row) per data row
    //   data is array      → array[i] per data row (length must match data row count)
    //   data === undefined → all nulls
    if (typeof name !== 'string') {
        throw new TypeError('z添加列: name must be a string, got ' + (name === null ? 'null' : typeof name));
    }
    var rows = this._items;
    var dataCount = rows.length > 0 ? rows.length - 1 : 0;  // exclude header
    var result = [];
    for (var i = 0; i < rows.length; i++) {
        var newRow = rows[i].slice();
        var isHeader = (i === 0);
        var cell;
        if (isHeader) {
            cell = name;
        } else if (typeof data === 'function') {
            try { cell = data(rows[i]); } catch (e) { cell = null; }
        } else if (data === undefined || data === null) {
            cell = null;
        } else if (Array.isArray(data)) {
            cell = data[i - 1] !== undefined ? data[i - 1] : null;
        } else {
            cell = data;
        }
        newRow.push(cell);
        result.push(newRow);
    }
    return this._new(result);
};`;

if (!src.includes(OLD)) {
    console.error('XXD-199 PATCH: old block not found verbatim — aborting.');
    process.exit(2);
}
if (src.indexOf(OLD) !== src.lastIndexOf(OLD)) {
    console.error('XXD-199 PATCH: old block matched more than once — aborting.');
    process.exit(2);
}

// Atomic write: tmp file in same dir, then rename. Beats Synology+WPS+iCloud race.
const tmp = target + '.xxd199.tmp';
fs.writeFileSync(tmp, src.replace(OLD, NEW), 'utf8');
fs.renameSync(tmp, target);
console.log('XXD-199 PATCH: applied. file=' + target);

// Inline verify
const verifyCode = `
const A = require(${JSON.stringify(target)});
const Array2D = A.Array2D || A.default || A;
let pass = 0, fail = 0;
function check(name, got, expected) {
  const ok = JSON.stringify(got) === JSON.stringify(expected);
  console.log((ok ? 'PASS' : 'FAIL') + ' ' + name + ' got=' + JSON.stringify(got) + ' exp=' + JSON.stringify(expected));
  ok ? pass++ : fail++;
}

// F3.1 z添加列('Z', fn) — append at end, header='Z', data=fn(row)
{
  const t = Array2D([['A','B'],[1,2],[3,4]]);
  const r = t.z添加列('Z', function(row){ return row[0] + row[1]; });
  check('F3.1', r._items, [['A','B','Z'],[1,2,3],[3,4,7]]);
}

// F3.2 z添加列(99) — TypeError
{
  const t = Array2D([['A','B'],[1,2]]);
  try { t.z添加列(99); console.log('FAIL F3.2 — did not throw'); fail++; }
  catch (e) { check('F3.2', e instanceof TypeError, true); }
}

// F3.3 z添加列('C', [10,20,30]) — header='C', data appended at end
{
  const t = Array2D([['A','B'],[1,2],[3,4]]);
  const r = t.z添加列('C', [10,20,30]);
  check('F3.3', r._items, [['A','B','C'],[1,2,10],[3,4,20]]);
}

// F3.4 z添加列(null) — TypeError
{
  const t = Array2D([['A','B'],[1,2]]);
  try { t.z添加列(null); console.log('FAIL F3.4 — did not throw'); fail++; }
  catch (e) { check('F3.4', e instanceof TypeError, true); }
}

// F3.5 z添加列('Z') — header='Z', data all null
{
  const t = Array2D([['A','B'],[1,2],[3,4]]);
  const r = t.z添加列('Z');
  check('F3.5', r._items, [['A','B','Z'],[1,2,null],[3,4,null]]);
}

// S6.5 z添加列(['X','Y'], fn) — TypeError (array as name)
{
  const t = Array2D([['A','B'],[1,2]]);
  try { t.z添加列(['X','Y'], function(){}); console.log('FAIL S6.5 — did not throw'); fail++; }
  catch (e) { check('S6.5', e instanceof TypeError, true); }
}

// Edge: empty data row count (only header)
{
  const t = Array2D([['A','B']]);
  const r = t.z添加列('Z', function(){ return 99; });
  check('edge:header-only', r._items, [['A','B','Z']]);
}

console.log('---');
console.log('pass=' + pass + ' fail=' + fail);
process.exit(fail ? 1 : 0);
`;
require('fs').writeFileSync('/tmp/xxd199-verify.cjs', verifyCode, 'utf8');
const { execSync } = require('child_process');
try {
    const out = execSync('node /tmp/xxd199-verify.cjs', { encoding: 'utf8' });
    console.log(out);
} catch (e) {
    console.error('VERIFY FAILED:');
    console.error(e.stdout || '');
    console.error(e.stderr || '');
    process.exit(1);
}
