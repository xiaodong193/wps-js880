// XXD-184: z内连接 fixes (idempotent + verified)
//
// Two related fixes, both in Array2D.prototype.z内连接:
//
// Fix A — numeric key-selector (the reported bug):
//   `leftKeySelector ? parseLambda(...) : defaultFn` is falsy when the caller
//   passes 0 (or any non-negative integer column index). parseLambda never
//   runs; we fall back to JSON.stringify(row); the join key never matches;
//   the result is always []. Build a column-picker function when the
//   selector is a finite non-negative integer.
//
// Fix B — array-form resultSelector (the follow-up API question):
//   resultSelector=[1] is a useful "extend left with right's col 1" call.
//   Previously any non-function/non-string resultSelector fell through to
//   default concat — so passing an array silently produced a different
//   shape than callers expected. New branch: an array of finite non-
//   negative integers is interpreted as "result row = left row +
//   [right[idx1], right[idx2], ...]". Reaches the issue's stated expected
//   output `[['x', 1, 10]]` for `.z内连接(brr, 0, 0, [1])`.
//   Non-numeric arrays (or empty arrays) still fall through to the
//   default-concat branch — no behavior change for them.

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const TARGET = path.join(__dirname, 'JSA880.js');

// --- Fix A: numeric key-selector -----------------------------------------

const OLD_NUMERIC_KEY = `Array2D.prototype.z内连接 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    var leftFn = leftKeySelector ? parseLambda(leftKeySelector) : function(row) { return JSON.stringify(row); };
    var rightFn = rightKeySelector ? parseLambda(rightKeySelector) : function(row) { return JSON.stringify(row); };`;

const NEW_NUMERIC_KEY = `Array2D.prototype.z内连接 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    function pickKeyFn(sel) {
        if (typeof sel === 'function') return sel;
        if (typeof sel === 'number' && Number.isFinite(sel) && sel >= 0) {
            return function(row) { return row == null ? undefined : row[sel]; };
        }
        if (sel) return parseLambda(sel);
        return function(row) { return JSON.stringify(row); };
    }
    var leftFn = pickKeyFn(leftKeySelector);
    var rightFn = pickKeyFn(rightKeySelector);`;

// --- Fix B: array-form resultSelector ------------------------------------

// Anchor on the "pre-build rightMap for O(M+N)" comment that follows the
// resultSelector block — that comment is unique to z内连接 (the sibling
// z左连接 says "XXD-16: ..." and z一对多连接 goes straight to "var rightMap").
const OLD_ARRAY_FORM = `    } else {
        // 默认：直接拼接
        resFn = function(a, b) { return a.concat(b || []); };
    }

    // pre-build rightMap for O(M+N)`;

const NEW_ARRAY_FORM = `    } else if (Array.isArray(resultSelector) && resultSelector.length > 0 && resultSelector.every(function(x){ return typeof x === 'number' && Number.isFinite(x) && x >= 0; })) {
        // XXD-184 array-form resultSelector: 数组形式 = 扩展左行 + 追加右表指定列
        // [idx1, idx2, ...] → result row = leftRow + [right[idx1], right[idx2], ...]
        var arrIdx = resultSelector;
        resFn = function(leftRow, rightRow) {
            var out = leftRow.slice();
            for (var s = 0; s < arrIdx.length; s++) {
                var i = arrIdx[s];
                out.push(rightRow == null ? null : rightRow[i]);
            }
            return out;
        };
    } else {
        // 默认：直接拼接
        resFn = function(a, b) { return a.concat(b || []); };
    }

    // pre-build rightMap for O(M+N)`;

// --- Idempotency markers --------------------------------------------------

function isPatchedNumericKey(src) {
    // Code marker is the source of truth; the footer comment is just a hint.
    return src.includes('function pickKeyFn(sel)');
}
function isPatchedArrayForm(src) {
    return src.includes('var arrIdx = resultSelector;');
}
function isPatchedAll(src) {
    return isPatchedNumericKey(src) && isPatchedArrayForm(src);
}

// --- Apply each patch independently so any single missing fix is re-added

function applyNumericKey(src) {
    if (isPatchedNumericKey(src)) return { src, changed: false, reason: 'numeric-key-already-patched' };
    if (!src.includes(OLD_NUMERIC_KEY)) {
        // Maybe the file was reverted past the pickKeyFn block, but our
        // marker comment is gone too — the whole XXD-184 has rolled back.
        // Throw so the operator knows.
        throw new Error('xxd184 numeric-key: target block not found verbatim');
    }
    return { src: src.replace(OLD_NUMERIC_KEY, NEW_NUMERIC_KEY), changed: true };
}

function applyArrayForm(src) {
    if (isPatchedArrayForm(src)) return { src, changed: false, reason: 'array-form-already-patched' };
    if (!src.includes(OLD_ARRAY_FORM)) {
        throw new Error('xxd184 array-form: target block not found verbatim');
    }
    return { src: src.replace(OLD_ARRAY_FORM, NEW_ARRAY_FORM), changed: true };
}

function patch() {
    const orig = fs.readFileSync(TARGET, 'utf8');
    if (isPatchedAll(orig)) {
        return { src: orig, changed: false, reason: 'already-patched' };
    }
    const r1 = applyNumericKey(orig);
    const r2 = applyArrayForm(r1.src);
    const finalSrc = r2.src;
    // mark both patches in the file footer for grep-ability
    const footer = [
        '',
        '// 🔧 XXD-184 final fix: numeric key selector (column index) support',
        '// 🔧 XXD-184 array-form resultSelector: [idx1,idx2,...] = extend-left + right[idx]',
        ''
    ].join('\n');
    const tmp = TARGET + '.xxd184.tmp';
    fs.writeFileSync(tmp, finalSrc + footer, 'utf8');
    fs.renameSync(tmp, TARGET);
    const after = fs.readFileSync(TARGET, 'utf8');
    if (!isPatchedAll(after)) {
        throw new Error('xxd184: marker not found after rename');
    }
    return { src: after, changed: r1.changed || r2.changed };
}

function verify() {
    const src = fs.readFileSync(TARGET, 'utf8');
    const ctx = { module: { exports: {} }, exports: {}, require: require, console: console, process: process, Buffer: Buffer, setTimeout: setTimeout, clearTimeout: clearTimeout, setInterval: setInterval, clearInterval: clearInterval };
    ctx.global = ctx;
    ctx.globalThis = ctx;
    vm.createContext(ctx);
    vm.runInContext(src, ctx);

    function check(label, got, expect) {
        const g = JSON.stringify(got);
        const e = JSON.stringify(expect);
        if (g !== e) {
            throw new Error('xxd184 verify ' + label + ': got=' + g + ' expect=' + e);
        }
        console.log('xxd184 verify ' + label + ' OK:', g);
    }

    // Case 1: numeric key selectors, default concat (the reported bug).
    // Before fix: result is []. After fix: matched rows joined via default
    // concat → left row + right row (keys duplicated, per the docstring
    // "默认拼接" example).
    const a = new ctx.Array2D([['A', 'B'], ['x', 1], ['y', 2]]);
    const r1 = a.z内连接([['A', 'C'], ['x', 10]], 0, 0);
    check('#1 numeric keys default concat (the reported bug)', r1._items, [['A', 'B', 'A', 'C'], ['x', 1, 'x', 10]]);

    // Case 2: numeric keys with the string selector 'a.f1,b.f2'. f1=1-based
    // col 0, f2=col 1. Matched 'A' row → left col 0 = 'A', right col 1 = 'C'.
    const r2 = a.z内连接([['A', 'C'], ['x', 10]], 0, 0, 'a.f1,b.f2');
    check('#2 numeric keys + string resultSelector', r2._items, [['A', 'C'], ['x', 10]]);

    // Case 3: string selector 'f1' still works (regression).
    const r3 = a.z内连接([['A', 'C'], ['x', 10]], 'f1', 'f1');
    check('#3 string selector regression', r3._items, [['A', 'B', 'A', 'C'], ['x', 1, 'x', 10]]);

    // Case 4: function selector still works (regression).
    const r4 = a.z内连接([['A', 'C'], ['x', 10]], function(row) { return row[0]; }, function(row) { return row[0]; });
    check('#4 function selector regression', r4._items, [['A', 'B', 'A', 'C'], ['x', 1, 'x', 10]]);

    // Case 5: numeric key with single-col string resultSelector. 'a.f1'
    // picks col 0 of each matched row → 'A' and 'x'.
    const r5 = a.z内连接([['A', 'C'], ['x', 10]], 0, 0, 'a.f1');
    check('#5 numeric key + single-col string resultSelector', r5._items, [['A'], ['x']]);

    // Case 6 (Fix B — the array-form): [1] means "extend left with right's
    // col 1". With the user's data (row 0 as data), matched 'A' and 'x'
    // rows both get right's col 1 appended.
    const r6 = a.z内连接([['A', 'C'], ['x', 10]], 0, 0, [1]);
    check('#6 array-form [1] (Fix B)', r6._items, [['A', 'B', 'C'], ['x', 1, 10]]);

    // Case 7: user's stated expected output, reading row 0 as a header.
    // Drop the header rows; only 'x' row matches.
    const a7 = new ctx.Array2D([['x', 1], ['y', 2]]);
    const r7 = a7.z内连接([['x', 10]], 0, 0, [1]);
    check('#7 user\'s expected output (header-stripped)', r7._items, [['x', 1, 10]]);

    // Case 8: multi-index array [1, 0] → left + right[1] + right[0].
    const r8 = a.z内连接([['A', 'C'], ['x', 10]], 0, 0, [1, 0]);
    check('#8 array-form [1, 0] multi-index', r8._items, [['A', 'B', 'C', 'A'], ['x', 1, 10, 'x']]);

    // Case 9: non-numeric array must fall through to default concat
    // (regression: existing behavior preserved when array form is invalid).
    const r9 = a.z内连接([['A', 'C'], ['x', 10]], 0, 0, ['foo']);
    check('#9 non-numeric array falls through to default concat', r9._items, [['A', 'B', 'A', 'C'], ['x', 1, 'x', 10]]);

    // Case 10: empty array also falls through to default concat.
    const r10 = a.z内连接([['A', 'C'], ['x', 10]], 0, 0, []);
    check('#10 empty array falls through to default concat', r10._items, [['A', 'B', 'A', 'C'], ['x', 1, 'x', 10]]);
}

(function main() {
    let attempts = 0;
    let lastErr;
    while (attempts < 5) {
        attempts++;
        try {
            const r = patch();
            verify();
            console.log('xxd184 ' + (r.changed ? 'PATCHED' : 'NO-OP (already patched)') + ' on attempt', attempts);
            return;
        } catch (e) {
            lastErr = e;
            console.log('xxd184 attempt', attempts, 'fail:', e.message);
            const wait = 1500 + Math.floor(Math.random() * 1500);
            const end = Date.now() + wait;
            while (Date.now() < end) { /* spin */ }
        }
    }
    throw lastErr;
})();
