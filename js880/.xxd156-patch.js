#!/usr/bin/env node
// XXD-156/XXD-157 final-fix patcher — loops past Synology+WPS+iCloud writer race
const fs = require('fs');
const path = require('path');
const FILE = path.join(__dirname, 'JSA880.js');

const ANCHOR_OLD = '/**\n * 求和\n * @param {...Number} args - 数值';
const REPLACEMENT_HEAD = '/* eslint-disable */\n// 🔧 XXD-156/XXD-157 final fix: JSA.z求和/z最大值/z最小值/z平均值 支持嵌套数组参数';
const MARKER = 'XXD-156/XXD-157 final fix';

const NEW_BLOCK =
`// 🔧 XXD-156/XXD-157 final fix: JSA.z求和/z最大值/z最小值/z平均值 支持嵌套数组参数
// 复现: JSA.z最大值([[1,2,3]]) 返回 0 — Number([[1,2,3]]) === NaN 被吞成 0
// 修复: 用 _zFlatNums 递归扁平化所有 array 参数, 跳过 NaN, 保留 ...Number 旧用法
function _zFlatNums() {
    var out = [];
    function walk(v) {
        if (Array.isArray(v)) { for (var i = 0; i < v.length; i++) walk(v[i]); return; }
        if (v === null || v === undefined || v === '') return;
        var n = (typeof v === 'number') ? v : parseFloat(String(v).replace(/,/g, ''));
        if (!isNaN(n)) out.push(n);
    }
    for (var i = 0; i < arguments.length; i++) walk(arguments[i]);
    return out;
}

/**
 * 求和
 * @param {...(Number|Array)} args - 数值或(嵌套)数组
 * @returns {Number} 和
 */
JSA.z求和 = function() {
    var nums = _zFlatNums.apply(null, arguments);
    var s = 0;
    for (var i = 0; i < nums.length; i++) s += nums[i];
    return Math.round(s * 1e10) / 1e10;
};
JSA.sum = JSA.z求和;

/**
 * 最大值
 * @param {...(Number|Array)} args - 数值或(嵌套)数组
 * @returns {Number} 最大值，无有效数值时返回 0
 */
JSA.z最大值 = function() {
    var nums = _zFlatNums.apply(null, arguments);
    return nums.length === 0 ? 0 : Math.max.apply(null, nums);
};
JSA.max = JSA.z最大值;

/**
 * 最小值
 * @param {...(Number|Array)} args - 数值或(嵌套)数组
 * @returns {Number} 最小值，无有效数值时返回 0
 */
JSA.z最小值 = function() {
    var nums = _zFlatNums.apply(null, arguments);
    return nums.length === 0 ? 0 : Math.min.apply(null, nums);
};
JSA.min = JSA.z最小值;

/**
 * 平均值
 * @param {...(Number|Array)} args - 数值或(嵌套)数组
 * @returns {Number} 平均值，无有效数值时返回 0
 */
JSA.z平均值 = function() {
    var nums = _zFlatNums.apply(null, arguments);
    if (nums.length === 0) return 0;
    var s = 0; for (var i = 0; i < nums.length; i++) s += nums[i];
    return Math.round((s / nums.length) * 1e10) / 1e10;
};
JSA.average = JSA.z平均值;
`;

// Match the bug block: from JSDoc "求和" through "JSA.average = JSA.z平均值;"
const BUG_RE = /\/\*\*\s*\n \* 求和\s*\n \* @param \{\.\.\.Number\} args[\s\S]*?JSA\.average = JSA\.z平均值;/;

function tryPatch() {
    let src = fs.readFileSync(FILE, 'utf8');
    if (src.includes(MARKER)) return { status: 'already', size: src.length };
    if (!BUG_RE.test(src)) return { status: 'no-match', size: src.length };
    const patched = src.replace(BUG_RE, NEW_BLOCK.trim());
    if (patched === src) return { status: 'noop', size: src.length };
    // atomic: write tmp + rename
    const tmp = FILE + '.xxd156.tmp';
    fs.writeFileSync(tmp, patched, 'utf8');
    fs.renameSync(tmp, FILE);
    // verify it stuck
    const after = fs.readFileSync(FILE, 'utf8');
    return { status: after.includes(MARKER) ? 'ok' : 'reverted', size: after.length };
}

(async function main() {
    const deadline = Date.now() + 120000; // 2 min budget
    let attempt = 0;
    while (Date.now() < deadline) {
        attempt++;
        let r;
        try { r = tryPatch(); }
        catch (e) { r = { status: 'err: ' + e.message }; }
        console.log(`[attempt ${attempt}] status=${r.status} size=${r.size||'-'}`);
        if (r.status === 'ok' || r.status === 'already') {
            // confirm the runtime behavior
            try {
                const vm = require('vm');
                const code = fs.readFileSync(FILE, 'utf8');
                const ctx = vm.createContext({ console, Math, Array, String, Number, parseFloat, isNaN, Date });
                vm.runInContext(code, ctx, { timeout: 5000 });
                const max1 = vm.runInContext('JSA.z最大值([[1,2,3]])', ctx);
                const sum1 = vm.runInContext('JSA.z求和([[1,2,3]])', ctx);
                const max2 = vm.runInContext('JSA.z最大值([1,5,3])', ctx);
                const max3 = vm.runInContext('JSA.z最大值([[1,2],[3,4]])', ctx);
                const sum2 = vm.runInContext('JSA.z求和([[1,2,3],[4,5,6]])', ctx);
                console.log(`VERIFY z最大值([[1,2,3]])=${max1} expect 3`);
                console.log(`VERIFY z求和([[1,2,3]])=${sum1} expect 6`);
                console.log(`VERIFY z最大值([1,5,3])=${max2} expect 5`);
                console.log(`VERIFY z最大值([[1,2],[3,4]])=${max3} expect 4`);
                console.log(`VERIFY z求和([[1,2,3],[4,5,6]])=${sum2} expect 21`);
                const ok = (max1 === 3 && sum1 === 6 && max2 === 5 && max3 === 4 && sum2 === 21);
                if (ok) {
                    console.log('PATCH OK + RUNTIME OK');
                    // keep watching for 15s to confirm writer doesn't revert it
                    const t0 = Date.now();
                    while (Date.now() - t0 < 15000) {
                        await new Promise(r => setTimeout(r, 2000));
                        const cur = fs.readFileSync(FILE, 'utf8');
                        if (!cur.includes(MARKER)) {
                            console.log('REVERTED during watch — re-patching');
                            attempt = 0;
                            break;
                        }
                        console.log(`watch: still patched at ${((Date.now()-t0)/1000).toFixed(0)}s`);
                    }
                    const final = fs.readFileSync(FILE, 'utf8');
                    if (final.includes(MARKER)) {
                        console.log('FINAL STATUS: patched and stable');
                        process.exit(0);
                    }
                } else {
                    console.log('PATCH applied but runtime mismatch');
                    process.exit(2);
                }
            } catch (e) {
                console.log('verify err:', e.message);
            }
        }
        await new Promise(r => setTimeout(r, 500));
    }
    console.log('FINAL STATUS: timeout');
    process.exit(1);
})();
