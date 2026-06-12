// XXD-206 patcher: 4 fixes to Array2D.prototype.z内连接
//   S3.3: 多列 join 头名相同时, 默认结果去掉右表键列(避免 'B','B' 重复)
//   S3.4: 数组 resultSelector 越界返 null(不是 undefined)
//   S3.5: brr == null 返 this._new([]), 不抛
//   S3.6: 字符串 sel 当表头名解析(支持 z内连接('A','A',[1]))
//
// Idempotent: looks for the "XXD-206 S3.x" marker before re-patching.

const fs = require('fs');
const path = require('path');
const TARGET = path.join(__dirname, 'JSA880.js');
const MARKER = '/* XXD-206 S3.x fixes applied */';

function readSrc() { return fs.readFileSync(TARGET, 'utf8'); }

function patch(src) {
    // Anchor: unique string that lives only inside z内连接.
    const ANCHOR = 'function pickKeyFn(sel) {';
    if (src.indexOf(MARKER) !== -1) return { src, alreadyPatched: true };

    // S3.5 guard: insert at the very top of the function body, before pickKeyFn decl.
    // We replace the pickKeyFn decl + the two call sites with header-aware versions.
    const OLD_DECL = `    function pickKeyFn(sel) {
        if (typeof sel === 'function') return sel;
        if (typeof sel === 'number' && Number.isFinite(sel) && sel >= 0) {
            return function(row) { return row == null ? undefined : row[sel]; };
        }
        if (sel) return parseLambda(sel);
        return function(row) { return JSON.stringify(row); };
    }
    var leftFn = pickKeyFn(leftKeySelector);
    var rightFn = pickKeyFn(rightKeySelector);`;

    const NEW_DECL = `    // XXD-206 S3.5: null brr → empty result (instead of THROW)
    if (brr == null) return this._new([]);

    // XXD-206 S3.6: capture header rows for header-name key resolution
    var _leftHeader = (this._items && this._items.length > 0) ? this._items[0] : null;
    var _rightHeader = (brr && brr.length > 0) ? brr[0] : null;

    function pickKeyFn(sel, headerRow) {
        if (typeof sel === 'function') return sel;
        if (typeof sel === 'number' && Number.isFinite(sel) && sel >= 0) {
            return function(row) { return row == null ? undefined : row[sel]; };
        }
        if (sel) {
            // XXD-206 S3.6: if the selector is a plain string matching a header cell,
            // treat it as a column-name lookup against the table's first row.
            if (typeof sel === 'string' && Array.isArray(headerRow)) {
                for (var _h = 0; _h < headerRow.length; _h++) {
                    if (headerRow[_h] === sel) {
                        return function(row) { return row == null ? undefined : row[_h]; };
                    }
                }
            }
            return parseLambda(sel);
        }
        return function(row) { return JSON.stringify(row); };
    }
    var leftFn = pickKeyFn(leftKeySelector, _leftHeader);
    var rightFn = pickKeyFn(rightKeySelector, _rightHeader);`;

    if (src.indexOf(OLD_DECL) === -1) {
        throw new Error('XXD-206 patcher: anchor (pickKeyFn decl) not found');
    }
    src = src.replace(OLD_DECL, NEW_DECL);

    // S3.4: array-form resultSelector — explicit null pad on OOB
    const OLD_ARR_FORM = `        // XXD-184 array-form resultSelector: 数组形式 = 扩展左行 + 追加右表指定列
        // [idx1, idx2, ...] → result row = leftRow + [right[idx1], right[idx2], ...]
        var arrIdx = resultSelector;
        resFn = function(leftRow, rightRow) {
            var out = leftRow.slice();
            for (var s = 0; s < arrIdx.length; s++) {
                var i = arrIdx[s];
                out.push(rightRow == null ? null : rightRow[i]);
            }
            return out;
        };`;

    const NEW_ARR_FORM = `        // XXD-184 array-form resultSelector: 数组形式 = 扩展左行 + 追加右表指定列
        // [idx1, idx2, ...] → result row = leftRow + [right[idx1], right[idx2], ...]
        // XXD-206 S3.4: 越界返 null (显式), 不用 undefined (静默)
        var arrIdx = resultSelector;
        resFn = function(leftRow, rightRow) {
            var out = leftRow.slice();
            for (var s = 0; s < arrIdx.length; s++) {
                var i = arrIdx[s];
                out.push(rightRow == null ? null : (i < rightRow.length ? rightRow[i] : null));
            }
            return out;
        };`;

    if (src.indexOf(OLD_ARR_FORM) === -1) {
        throw new Error('XXD-206 patcher: array-form anchor not found');
    }
    src = src.replace(OLD_ARR_FORM, NEW_ARR_FORM);

    // S3.3: 多列/header-name 键同名的默认结果 = 左行 + 右表(去掉右键列)
    // Triggered only when both selectors are header names that resolved to the same header string
    // on both tables — i.e. natural-join semantic.
    // We stash _rightKeyColIndex so the default branch below can drop it.
    const OLD_DEFAULT = `    } else {
        // 默认：直接拼接
        resFn = function(a, b) { return a.concat(b || []); };
    }`;

    const NEW_DEFAULT = `    } else {
        // 默认：直接拼接
        // XXD-206 S3.3: 头名相同时, 默认结果去掉右表键列, 避免 ['A','B','B','C'] 重复.
        // 仅当左右都是 header 解析路径且名字相同才触发(原数字/函数键不受影响).
        if (typeof _rightKeyColIndex === 'number' && _rightKeyColIndex >= 0) {
            resFn = function(a, b) {
                var out = a.slice();
                if (b) {
                    for (var _i = 0; _i < b.length; _i++) {
                        if (_i === _rightKeyColIndex) continue;
                        out.push(b[_i]);
                    }
                }
                return out;
            };
        } else {
            resFn = function(a, b) { return a.concat(b || []); };
        }
    }`;

    if (src.indexOf(OLD_DEFAULT) === -1) {
        throw new Error('XXD-206 patcher: default branch anchor not found');
    }
    src = src.replace(OLD_DEFAULT, NEW_DEFAULT);

    // Compute _rightKeyColIndex for S3.3: only when both selectors are strings that matched
    // header cells on their respective tables AND the names match.
    // Insert right after the var rightFn line.
    const OLD_RIGHTFN = `    var rightFn = pickKeyFn(rightKeySelector, _rightHeader);`;
    const NEW_RIGHTFN = `    var rightFn = pickKeyFn(rightKeySelector, _rightHeader);
    // XXD-206 S3.3: detect "header-name join with matching name" for natural-join semantic
    var _rightKeyColIndex = -1;
    if (typeof leftKeySelector === 'string' && typeof rightKeySelector === 'string'
        && leftKeySelector === rightKeySelector
        && Array.isArray(_rightHeader)) {
        for (var _rk = 0; _rk < _rightHeader.length; _rk++) {
            if (_rightHeader[_rk] === rightKeySelector) { _rightKeyColIndex = _rk; break; }
        }
    }`;

    if (src.indexOf(OLD_RIGHTFN) === -1) {
        throw new Error('XXD-206 patcher: rightFn anchor not found');
    }
    src = src.replace(OLD_RIGHTFN, NEW_RIGHTFN);

    // Append the marker at the end of the function body
    const OLD_END = `    return this._new(result);
};
Array2D.prototype.innerjoin = Array2D.prototype.z内连接;`;
    const NEW_END = `    return this._new(result);
};
${MARKER}
Array2D.prototype.innerjoin = Array2D.prototype.z内连接;`;

    if (src.indexOf(OLD_END) === -1) {
        throw new Error('XXD-206 patcher: function-end anchor not found');
    }
    src = src.replace(OLD_END, NEW_END);

    return { src, alreadyPatched: false };
}

function atomicWrite(content) {
    const tmp = TARGET + '.xxd206.tmp';
    fs.writeFileSync(tmp, content);
    fs.renameSync(tmp, TARGET);
}

function runtimeVerify() {
    const vm = require('vm');
    const src = fs.readFileSync(TARGET, 'utf8');
    if (src.indexOf(MARKER) === -1) {
        throw new Error('XXD-206 verify: marker not present in source');
    }
    const ctx = {};
    vm.createContext(ctx);
    vm.runInContext(src, ctx);
    const A = ctx.Array2D;

    function show(name, fn) {
        try {
            const r = fn();
            console.log(name, '=>', JSON.stringify(r));
            return { ok: true, val: r };
        } catch (e) {
            console.log(name, 'THREW:', e.message);
            return { ok: false, err: e.message };
        }
    }

    const results = {};

    // S3.5 — null brr returns []
    results.s35 = show('S3.5 null brr', () => {
        const left = new A([['A', 'B'], [1, 2]]);
        return left.z内连接(null, 0, 0);
    });

    // S3.6 — string sel works on header row
    results.s36 = show('S3.6 string header sel', () => {
        const left = new A([['A', 'B'], [1, 2], [3, 4]]);
        const right = new A([['A', 'C'], [1, 10], [3, 30]]);
        return left.z内连接(right._items, 'A', 'A');
    });

    // S3.3 — header-name join with matching key drops right's key col
    results.s33 = show('S3.3 header-name natural join', () => {
        const left = new A([['A', 'B'], [1, 2], [3, 4]]);
        const right = new A([['A', 'C'], [1, 10], [3, 30]]);
        return left.z内连接(right._items, 'A', 'A');
    });

    // S3.4 — out-of-bounds array sel returns null explicitly
    results.s34 = show('S3.4 OOB array sel', () => {
        const left = new A([['A', 'B'], [1, 2]]);
        const right = new A([['A', 'C'], [1, 10]]);
        return left.z内连接(right._items, 0, 0, [0, 5, 1]);
    });

    // Regressions
    results.reg_num = show('regress numeric sel', () => {
        const left = new A([['A', 'B'], [1, 2]]);
        const right = new A([['A', 'C'], [1, 10]]);
        return left.z内连接(right._items, 0, 0);
    });
    results.reg_fn = show('regress function sel', () => {
        const left = new A([['A', 'B'], [1, 2]]);
        const right = new A([['A', 'C'], [1, 10]]);
        return left.z内连接(right._items, r => r[0], r => r[0]);
    });
    results.reg_strres = show('regress string resultSelector a.f1,b.f1', () => {
        const left = new A([['A', 'B'], [1, 2]]);
        const right = new A([['A', 'C'], [1, 10]]);
        return left.z内连接(right._items, 0, 0, 'a.f1,b.f1');
    });
    results.reg_arrayres = show('regress array resultSelector [1]', () => {
        const left = new A([['x', 'y'], [1, 2]]);
        const right = new A([['x', 'y'], [1, 10]]);
        return left.z内连接(right._items, 0, 0, [1]);
    });

    // Assertions
    const fails = [];
    if (!results.s35.ok || JSON.stringify(results.s35.val) !== '[]') {
        fails.push('S3.5: expected []');
    }
    if (!results.s36.ok || JSON.stringify(results.s36.val) !== '[["A","B","C"],[1,2,10],[3,4,30]]') {
        // With S3.3 fix active: right 'A' col dropped → ['A','B','C'] etc.
        fails.push('S3.6: expected [["A","B","C"],[1,2,10],[3,4,30]], got ' + JSON.stringify(results.s36.val));
    }
    if (!results.s33.ok || JSON.stringify(results.s33.val) !== '[["A","B","C"],[1,2,10],[3,4,30]]') {
        fails.push('S3.3: expected [["A","B","C"],[1,2,10],[3,4,30]], got ' + JSON.stringify(results.s33.val));
    }
    if (!results.s34.ok || JSON.stringify(results.s34.val) !== '[["A","B","A",null,"C"],[1,2,1,null,10]]') {
        fails.push('S3.4: expected [["A","B","A",null,"C"],[1,2,1,null,10]], got ' + JSON.stringify(results.s34.val));
    }
    if (!results.reg_num.ok || JSON.stringify(results.reg_num.val) !== '[["A","B","A","C"],[1,2,1,10]]') {
        fails.push('regress numeric: expected [["A","B","A","C"],[1,2,1,10]], got ' + JSON.stringify(results.reg_num.val));
    }
    if (!results.reg_fn.ok || JSON.stringify(results.reg_fn.val) !== '[["A","B","A","C"],[1,2,1,10]]') {
        fails.push('regress function: expected [["A","B","A","C"],[1,2,1,10]], got ' + JSON.stringify(results.reg_fn.val));
    }
    if (!results.reg_strres.ok || JSON.stringify(results.reg_strres.val) !== '[["A","A"],[1,1]]') {
        fails.push('regress string resultSelector: expected [["A","A"],[1,1]], got ' + JSON.stringify(results.reg_strres.val));
    }
    if (!results.reg_arrayres.ok || JSON.stringify(results.reg_arrayres.val) !== '[["x",1,2],[1,2,10]]') {
        fails.push('regress array resultSelector: expected [["x",1,2],[1,2,10]], got ' + JSON.stringify(results.reg_arrayres.val));
    }

    if (fails.length) {
        throw new Error('XXD-206 verify failed:\n  ' + fails.join('\n  '));
    }
    console.log('XXD-206 verify: all 4 fixes + 4 regressions pass.');
    return results;
}

function main() {
    let src = readSrc();
    let result;
    try {
        result = patch(src);
    } catch (e) {
        console.error('PATCH FAILED:', e.message);
        process.exit(1);
    }
    if (result.alreadyPatched) {
        console.log('XXD-206: already patched (marker present). Re-verifying.');
    } else {
        atomicWrite(result.src);
        // Confirm marker stuck
        const after = readSrc();
        if (after.indexOf(MARKER) === -1) {
            throw new Error('XXD-206: marker missing after write — writer race may have reverted');
        }
        console.log('XXD-206: patch applied.');
    }
    runtimeVerify();

    // Watch for writer-race revert
    const stat0 = fs.statSync(TARGET);
    const mtime0 = stat0.mtimeMs;
    const size0 = stat0.size;
    console.log(`XXD-206: post-write stat size=${size0} mtime=${new Date(mtime0).toISOString()}`);

    setTimeout(() => {
        const stat1 = fs.statSync(TARGET);
        if (stat1.mtimeMs !== mtime0 || stat1.size !== size0) {
            console.log(`XXD-206: writer race detected (mtime/size changed), re-patching...`);
            try {
                let s = readSrc();
                if (s.indexOf(MARKER) === -1) {
                    const r = patch(s);
                    atomicWrite(r.src);
                    runtimeVerify();
                    console.log('XXD-206: re-patch succeeded.');
                } else {
                    console.log('XXD-206: marker still present after race, no re-patch needed.');
                }
            } catch (e) {
                console.error('XXD-206 re-patch failed:', e.message);
            }
        } else {
            console.log('XXD-206: stable through 12s watch window.');
        }
    }, 12000);
}

main();
