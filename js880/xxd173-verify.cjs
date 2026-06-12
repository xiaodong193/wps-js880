// XXD-173/XXD-174 regression check
// Run: node xxd173-verify.cjs  →  exits 0 on pass, 1 on fail
const fs = require('fs');
const vm = require('vm');
const FILE = 'JSA880.js';

function run(label, expr, expected) {
    const ctx = vm.createContext({ console, Math, Array, String, Number, Date });
    vm.runInContext(fs.readFileSync(FILE, 'utf8'), ctx, { timeout: 8000 });
    const got = vm.runInContext(expr, ctx);
    const ok = got === expected;
    console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}  ${expr} = ${JSON.stringify(got)}  (expected ${JSON.stringify(expected)})`);
    return ok;
}

const results = [
    run('basic',         "IO.z路径拼接('/a','b','c')",     '/a/b/c'),
    run('trailing-slash',"IO.z路径拼接('/a/','/b/','/c/')",'/a/b/c/'),
    run('filter-null',   "IO.z路径拼接('a','','b',null,'c')",'a/b/c'),
    run('alias-pathJoin',"IO.pathJoin('x','y','z')",        'x/y/z'),
    run('alias-joinPath',"IO.joinPath('p','q')",             'p/q'),
    run('marker-present',"typeof IO.z路径拼接 === 'function' ? 'yes' : 'no'", 'yes'),
];
process.exit(results.every(Boolean) ? 0 : 1);
