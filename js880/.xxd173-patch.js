#!/usr/bin/env node
// XXD-173/XXD-174 final-fix patcher: IO.z路径拼接 (path-join alias)
// loops past Synology+WPS+iCloud writer race
const fs = require('fs');
const path = require('path');
const FILE = path.join(__dirname, 'JSA880.js');

const MARKER = 'XXD-173/XXD-174 final fix';
const ANCHOR = 'IO.lastDirectoty = IO.z上级文件夹;';

const NEW_BLOCK =
`IO.lastDirectoty = IO.z上级文件夹;

/* eslint-disable */
// 🔧 XXD-173/XXD-174 final fix: IO.z路径拼接 — 缺失的 path.join 中文别名
// 复现: IO.z路径拼接('/a','b','c') THROW — IO.z路径拼接 is not a function
// 期望: '/a/b/c' — 类似 Node path.join: 过滤 null/空, 用 '/' 连接, 折叠连续 '/'
/**
 * 路径拼接
 * @param {...String} parts - 路径片段
 * @returns {String} 拼接后的路径
 */
IO.z路径拼接 = function() {
    var p = Array.prototype.slice.call(arguments).filter(function(x) {
        return x != null && x !== '';
    }).join('/');
    return p.replace(/\\/+/g, '/');
};
IO.pathJoin = IO.z路径拼接;
IO.joinPath = IO.z路径拼接;
`;

function tryPatch() {
    let src = fs.readFileSync(FILE, 'utf8');
    if (src.includes(MARKER)) return { status: 'already', size: src.length };
    if (!src.includes(ANCHOR)) return { status: 'no-anchor', size: src.length };
    // Only patch if the function doesn't already exist (defensive — to avoid dup if writer overwrote a partial state)
    if (src.includes('IO.z路径拼接 = function')) return { status: 'already-fn', size: src.length };
    const patched = src.replace(ANCHOR, NEW_BLOCK.trim());
    if (patched === src) return { status: 'noop', size: src.length };
    const tmp = FILE + '.xxd173.tmp';
    fs.writeFileSync(tmp, patched, 'utf8');
    fs.renameSync(tmp, FILE);
    const after = fs.readFileSync(FILE, 'utf8');
    return { status: after.includes(MARKER) ? 'ok' : 'reverted', size: after.length };
}

(async function main() {
    const deadline = Date.now() + 120000;
    let attempt = 0;
    while (Date.now() < deadline) {
        attempt++;
        let r;
        try { r = tryPatch(); }
        catch (e) { r = { status: 'err: ' + e.message }; }
        console.log(`[attempt ${attempt}] status=${r.status} size=${r.size||'-'}`);
        if (r.status === 'ok' || r.status === 'already' || r.status === 'already-fn') {
            try {
                const vm = require('vm');
                const code = fs.readFileSync(FILE, 'utf8');
                const ctx = vm.createContext({ console, Math, Array, String, Number, Date });
                vm.runInContext(code, ctx, { timeout: 8000 });
                const t1 = vm.runInContext("IO.z路径拼接('/a','b','c')", ctx);
                const t2 = vm.runInContext("IO.z路径拼接('/a/','/b/','/c/')", ctx);
                const t3 = vm.runInContext("IO.z路径拼接('a','','b',null,'c')", ctx);
                const t4 = vm.runInContext("IO.pathJoin('x','y','z')", ctx);
                const t5 = vm.runInContext("IO.joinPath('p','q')", ctx);
                console.log(`VERIFY z路径拼接('/a','b','c')=${t1} expect /a/b/c`);
                console.log(`VERIFY z路径拼接('/a/','/b/','/c/')=${t2} expect /a/b/c/`);
                console.log(`VERIFY z路径拼接('a','','b',null,'c')=${t3} expect a/b/c`);
                console.log(`VERIFY pathJoin('x','y','z')=${t4} expect x/y/z`);
                console.log(`VERIFY joinPath('p','q')=${t5} expect p/q`);
                const ok = (t1 === '/a/b/c' && t2 === '/a/b/c/' && t3 === 'a/b/c' && t4 === 'x/y/z' && t5 === 'p/q');
                if (ok) {
                    console.log('PATCH OK + RUNTIME OK');
                    // keep watching for 15s to confirm writer doesn't revert it
                    const t0 = Date.now();
                    let stable = true;
                    while (Date.now() - t0 < 15000) {
                        await new Promise(r => setTimeout(r, 2000));
                        const cur = fs.readFileSync(FILE, 'utf8');
                        if (!cur.includes(MARKER) && !cur.includes('IO.z路径拼接 = function')) {
                            console.log('REVERTED during watch — re-patching');
                            stable = false;
                            attempt = 0;
                            break;
                        }
                        console.log(`watch: still patched at ${((Date.now()-t0)/1000).toFixed(0)}s`);
                    }
                    if (stable) {
                        const final = fs.readFileSync(FILE, 'utf8');
                        if (final.includes(MARKER)) {
                            console.log('FINAL STATUS: patched and stable');
                            process.exit(0);
                        }
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
