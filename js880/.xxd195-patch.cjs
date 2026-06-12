#!/usr/bin/env node
// XXD-195 / XXD-196 final fix:
// Array2D.prototype.z归约 (and the static Array2D.reduce alias) reduced over the
// outer rows instead of the leaf cells, so initialValue was being applied to the
// wrong sequence. For new Array2D([[1,2,3]]).z归约((a,b)=>a+b, 0) we want
// 0+1+2+3=6, not 0+[1,2,3]=6 (string concat) or 0+1=1 (row-style bug).
//
// Contract: z归约(cb, init) walks the 2D leaves left-to-right, exactly like
// z求和/z最大值/z最小值/z平均值/z扁平化 do. The static Array2D.reduce
// already delegates to z归约, so fixing the prototype is enough.

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const FILE = path.join(__dirname, 'JSA880.js');
const TMP  = FILE + '.xxd195.tmp';

const OLD = `Array2D.prototype.z归约 = function(callback, initialValue) {
    return this._items.reduce(callback, initialValue);
};
Array2D.prototype.reduce = Array2D.prototype.z归约;`;

const NEW = `// 🔧 XXD-195/XXD-196 final fix: 走扁平化后的叶子序列, 与 z求和/z最大值/z最小值/z平均值 一致
//  之前 this._items.reduce 把外层每一行作为元素传入, 起点 initialValue 只调一次,
//  后续累加的 b 仍是行, 与 2D 数值/表格场景的直觉 (按格累加) 相反
//  传 initialValue 时条件转发, 避免 V8 在 reduce(cb, undefined) 上把 undefined 当成累加器
//  → NaN (arguments.length 决定是否取首元素, 显式传 undefined 长度仍为 2)
Array2D.prototype.z归约 = function(callback, initialValue) {
    var flat = this.z扁平化();
    return arguments.length < 2 ? flat.reduce(callback) : flat.reduce(callback, initialValue);
};
Array2D.prototype.reduce = Array2D.prototype.z归约;`;

function readFile() {
    return fs.readFileSync(FILE, 'utf8');
}

// Previous (incomplete) patch that still uses .reduce(cb, initialValue) —
//  not just the new comment. We treat the comment line as the marker.
const PRIOR_NEW = `// 🔧 XXD-195/XXD-196 final fix: 走扁平化后的叶子序列, 与 z求和/z最大值/z最小值/z平均值 一致
//  之前 this._items.reduce 把外层每一行作为元素传入, 起点 initialValue 只调一次,
//  后续累加的 b 仍是行, 与 2D 数值/表格场景的直觉 (按格累加) 相反
Array2D.prototype.z归约 = function(callback, initialValue) {
    return this.z扁平化().reduce(callback, initialValue);
};
Array2D.prototype.reduce = Array2D.prototype.z归约;`;

function applyOnce() {
    const src = readFile();
    // If the latest NEW block is already in place, no work needed.
    if (src.includes(NEW)) {
        return { ok: false, reason: 'already patched (current NEW present)' };
    }
    // If a previous (incomplete) patch is present, strip it first.
    let working = src;
    if (working.includes(PRIOR_NEW)) {
        working = working.replace(PRIOR_NEW, NEW);
        fs.writeFileSync(TMP, working, 'utf8');
        fs.renameSync(TMP, FILE);
        return { ok: true, upgraded: true };
    }
    if (!working.includes(OLD)) {
        return { ok: false, reason: 'OLD block not found (already patched or moved)' };
    }
    const patched = working.replace(OLD, NEW);
    if (patched === working) {
        return { ok: false, reason: 'replace produced identical content' };
    }
    fs.writeFileSync(TMP, patched, 'utf8');
    fs.renameSync(TMP, FILE);
    return { ok: true };
}

function verify() {
    const src = readFile();
    if (!src.includes('// 🔧 XXD-195/XXD-196 final fix')) {
        return { ok: false, reason: 'marker missing after patch' };
    }
    // Runtime check in a fresh VM context.
    const ctx = { console, Array2D: undefined, module: { exports: {} } };
    vm.createContext(ctx);
    try {
        vm.runInContext(src, ctx, { timeout: 5000 });
    } catch (e) {
        return { ok: false, reason: 'eval threw: ' + e.message };
    }
    if (typeof ctx.Array2D !== 'function') {
        return { ok: false, reason: 'Array2D not exposed on global after eval' };
    }
    const A2D = ctx.Array2D;
    // Case 1: issue repro (with the corrected leaf-style callback)
    const r1 = new A2D([[1, 2, 3]]).z归约(function (a, b) { return a + b; }, 0);
    if (r1 !== 6) return { ok: false, reason: 'case1 expected 6 got ' + r1 };
    // Case 2: initialValue omitted (regression — must still skip and start at first leaf)
    const r2 = new A2D([[1, 2], [3, 4]]).z归约(function (a, b) { return a + b; });
    if (r2 !== 10) return { ok: false, reason: 'case2 expected 10 got ' + r2 };
    // Case 3: initialValue respected (the actual B8.1 invariant)
    const r3 = new A2D([[1, 2], [3, 4]]).z归约(function (a, b) { return a + b; }, 100);
    if (r3 !== 110) return { ok: false, reason: 'case3 expected 110 got ' + r3 };
    // Case 4: static Array2D.reduce mirrors the prototype
    const r4 = A2D.reduce([[1, 2, 3]], function (a, b) { return a + b; }, 0);
    if (r4 !== 6) return { ok: false, reason: 'case4 expected 6 got ' + r4 };
    return { ok: true, results: { r1, r2, r3, r4 } };
}

function main() {
    const MAX = 10;
    for (let i = 1; i <= MAX; i++) {
        const apply = applyOnce();
        if (!apply.ok) {
            // Already patched previously? Verify and exit gracefully.
            if (/already patched/.test(apply.reason)) {
                console.log('[xxd195] already patched, verifying only');
                break;
            }
            console.error('[xxd195] apply failed:', apply.reason);
            process.exit(2);
        }
        console.log('[xxd195] patch applied (attempt', i + ')');
        const v = verify();
        if (v.ok) {
            console.log('[xxd195] verify OK:', v.results);
            // Watch the file for ~12s — if the writer race reverts it, the marker
            // disappears and we'll re-apply on the next run.
            let drifts = 0;
            for (let t = 0; t < 6; t++) {
                const wait = Date.now() + 2000;
                while (Date.now() < wait) {} // sync sleep 2s
                const src = readFile();
                if (!src.includes('// 🔧 XXD-195/XXD-196 final fix')) {
                    drifts++;
                    console.warn('[xxd195] marker disappeared at t+' + (t * 2) + 's, re-applying');
                    const reap = applyOnce();
                    if (!reap.ok) { console.error('[xxd195] reapply failed:', reap.reason); process.exit(3); }
                    const v2 = verify();
                    if (!v2.ok) { console.error('[xxd195] reverify failed:', v2.reason); process.exit(4); }
                }
            }
            if (drifts === 0) console.log('[xxd195] marker held through 12s watch window');
            process.exit(0);
        } else {
            console.error('[xxd195] verify failed:', v.reason);
            process.exit(5);
        }
    }
}

main();
