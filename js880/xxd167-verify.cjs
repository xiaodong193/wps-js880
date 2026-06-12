// XXD-167 regression verifier: StrUtils Chinese aliases must be present
// and behaviorally equivalent to their English counterparts.
// Run: node xxd167-verify.cjs
const fs = require('fs');
const vm = require('vm');

const TARGET = '/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880/JSA880.js';

const ALIASES = [
    ['z去空白',     'trim'],
    ['z去左空白',   'trimLeft'],
    ['z去右空白',   'trimRight'],
    ['z转大写',     'toUpperCase'],
    ['z转小写',     'toLowerCase'],
    ['z包含',       'contains'],
    ['z开始于',     'startsWith'],
    ['z结束于',     'endsWith'],
    ['z分割',       'split'],
    ['z连接',       'join'],
    ['z替换',       'replaceAll'],
    ['z替换全部',   'replaceAll'],
    ['z左填充',     'padLeft'],
    ['z右填充',     'padRight'],
    ['z重复',       'repeat'],
    ['z首字母大写', 'capitalize'],
    ['z驼峰命名',   'camelCase'],
    ['z下划线命名', 'snakeCase'],
    ['z是否为空',   'isEmpty'],
    ['z是否空白',   'isBlank'],
    ['z是否数字',   'isNumeric'],
    ['z是否整数',   'isInteger'],
    ['z是否字母',   'isAlpha'],
    ['z是否字母数字', 'isAlphanumeric'],
    ['z转义HTML',   'escapeHtml'],
    ['z反转义HTML', 'unescapeHtml'],
    ['z左取',       'left'],
    ['z右取',       'right'],
    ['z截取',       'substring'],
    ['z转数字',     'toNumber'],
    ['z去除前缀',   'removePrefix'],
    ['z去除后缀',   'removeSuffix'],
    ['z模板',       'template'],
    ['z计数',       'count'],
];

function main() {
    const src = fs.readFileSync(TARGET, 'utf8');
    if (!src.includes('XXD-167')) {
        console.error('FAIL: XXD-167 marker missing in ' + TARGET);
        process.exit(1);
    }
    const ctx = { console, module: { exports: {} }, require, setTimeout, clearTimeout, WPS: { LoadEvent: {} }, ActiveXObject: function () {} };
    vm.createContext(ctx);
    try { vm.runInContext(src, ctx, { filename: 'JSA880.js', timeout: 15000 }); } catch (_) {}
    const S = ctx.StrUtils;
    if (!S) { console.error('FAIL: StrUtils not defined'); process.exit(1); }
    let failed = 0;
    for (const [zh, en] of ALIASES) {
        if (typeof S[zh] !== 'function') {
            console.error(`FAIL: StrUtils.${zh} is ${typeof S[zh]}`);
            failed++;
        }
    }
    const checks = [
        ['z去空白',     ['  hi  '],                       'hi'],
        ['z转大写',     ['abc'],                          'ABC'],
        ['z转小写',     ['ABC'],                          'abc'],
        ['z是否为空',   [''],                             true],
        ['z是否空白',   ['   '],                          true],
        ['z是否数字',   ['123'],                          true],
        ['z替换',       ['a-b-c', '-', '+'],              'a+b+c'],
        ['z计数',       ['ababab', 'ab'],                  3],
        ['z首字母大写', ['hello'],                        'Hello'],
        ['z驼峰命名',   ['foo_bar_baz'],                   'fooBarBaz'],
        ['z下划线命名', ['FooBarBaz'],                     'foo_bar_baz'],
        ['z转义HTML',   ['<a>&"\'</a>'],                   '&lt;a&gt;&amp;&quot;&#39;&lt;/a&gt;'],
        ['z反转义HTML', ['&lt;a&gt;'],                     '<a>'],
        ['z去除前缀',   ['foobar', 'foo'],                 'bar'],
        ['z去除后缀',   ['foobar', 'bar'],                 'foo'],
        ['z转数字',     ['1,234.5'],                       1234.5],
    ];
    for (const [k, args, want] of checks) {
        const got = S[k].apply(S, args);
        if (got !== want) {
            console.error(`FAIL: ${k}(${JSON.stringify(args)}) = ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
            failed++;
        }
    }
    if (failed) { console.error(`XXD-167 verify FAILED: ${failed} issue(s)`); process.exit(1); }
    console.log(`XXD-167 verify OK: ${ALIASES.length} Chinese aliases present, ${checks.length} behavior checks pass.`);
}

main();
