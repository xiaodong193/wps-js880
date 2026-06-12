// XXD-180/XXD-181 verify: SuperMap.z分组 and SuperMap.z分组统计 shape contract.
// Run from the js880 dir: node .xxd180-verify.cjs
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const TARGET = path.join('/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880', 'JSA880.js');

function assert(cond, msg) {
    if (!cond) {
        console.error('FAIL:', msg);
        process.exit(1);
    }
}

const text = fs.readFileSync(TARGET, 'utf8');
const ctx = { console, globalThis: {} };
ctx.global = ctx;
vm.createContext(ctx);
vm.runInContext(text, ctx);
const SM = ctx.SuperMap;

const arr = [['A','B'],[1,2],[1,2],[3,4]];

// 1. z分组 must return a dict (not array) keyed by group key, value = array of rows.
const g = SM.z分组(arr, 0);
assert(typeof g === 'object' && !Array.isArray(g),
       'z分组 must return a dict (object, not array)');
assert(Array.isArray(g['1']) && g['1'].length === 2,
       'z分组 result for key "1" should have 2 rows');
assert(Array.isArray(g['3']) && g['3'].length === 1,
       'z分组 result for key "3" should have 1 row');

// 2. z分组统计 must return a 2D array with header row.
const s = SM.z分组统计(arr, 0, 1);
assert(Array.isArray(s) && Array.isArray(s[0]),
       'z分组统计 must return a 2D array (array of arrays)');
assert(s[0][0] === 'A',
       'z分组统计 header first cell should be the group column header');
assert(s.length === 3,
       'z分组统计 should have 1 header + 2 group rows (keys "1" and "3")');
assert(s[1][0] === '1' && s[2][0] === '3',
       'z分组统计 group rows should contain keys 1 then 3');

// 3. JSDoc on both methods must mention the return shape.
assert(/@returns \{Object<string, Array<Array>>\}/.test(text),
       'z分组 JSDoc must declare dict @returns');
assert(/@returns \{Array<Array>\}/.test(text),
       'z分组统计 JSDoc must declare 2D-array @returns');

// 4. The XXD-180/XXD-181 markers must be present (fix landed and stuck).
assert(/XXD-180\/XXD-181 final fix start/.test(text),
       'start marker missing — patch did not land or was reverted');
assert(/XXD-180\/XXD-181 final fix end/.test(text),
       'end marker missing — patch did not land or was reverted');

// 5. v4.0.39 changelog block must reference this ticket.
assert(/更新日志 \(v4\.0\.39 — 2026-06-11\)/.test(text),
       'v4.0.39 changelog block missing');
assert(/XXD-180\/XXD-181/.test(text),
       'changelog must mention XXD-180/XXD-181');

console.log('OK: XXD-180/XXD-181 shape contract holds');
console.log('  z分组       → dict keys: 1, 3, A');
console.log('  z分组统计   → 2D array, header ["A"], rows 1 and 3');
console.log('  JSDoc, markers, and v4.0.39 changelog all present');
process.exit(0);
