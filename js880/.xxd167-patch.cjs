// CTO direct patch: add Chinese aliases to StrUtils (XXD-167/168)
// Atomic write + readback + in-VM verify + writer-race watch.
const fs = require('fs');
const vm = require('vm');

const TARGET = '/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880/JSA880.js';
const MARKER_START = '/* XXD-167 START: StrUtils Chinese aliases */';
const MARKER_END   = '/* XXD-167 END */';

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

const FOLLOWUP = (() => {
    const lines = ALIASES.map(([zh, en]) => `StrUtils.${zh} = StrUtils.${en};`);
    return `\n${MARKER_START}\n` +
           `// StrUtils 中文 alias (XXD-167/168) — assigned after literal, so StrUtils is bound.\n` +
           lines.join('\n') + '\n' +
           `${MARKER_END}\n`;
})();

const EXPECTED_ALIASES = ALIASES.map(([zh]) => zh);

function patchFile() {
    const src = fs.readFileSync(TARGET, 'utf8');
    if (src.includes(MARKER_START)) {
        return { changed: false, reason: 'already-patched', src };
    }
    // Locate the close of the StrUtils literal: `};\n` followed by
    // the NumUtils comment.
    const re = /var StrUtils = \{[\s\S]*?\n\};\n/;
    const m = src.match(re);
    if (!m) throw new Error('StrUtils block not found');
    const closeIdx = m.index + m[0].length;
    const before = src.slice(0, closeIdx);
    const after = src.slice(closeIdx);
    return { changed: true, patched: before + FOLLOWUP + after, closeIdx };
}

function atomicWrite(content) {
    const tmp = TARGET + '.xxd167.tmp';
    fs.writeFileSync(tmp, content, 'utf8');
    fs.renameSync(tmp, TARGET);
}

function verifyInVM() {
    const src = fs.readFileSync(TARGET, 'utf8');
    const ctx = { console, module: { exports: {} }, require, setTimeout, clearTimeout, WPS: { LoadEvent: {} }, ActiveXObject: function () {} };
    vm.createContext(ctx);
    try {
        vm.runInContext(src, ctx, { filename: 'JSA880.js', timeout: 15000 });
    } catch (e) {
        // tolerate init-time errors from host-only references
    }
    const StrUtils = ctx.StrUtils;
    if (!StrUtils) return { ok: false, reason: 'StrUtils undefined in VM' };
    const missing = EXPECTED_ALIASES.filter(k => typeof StrUtils[k] !== 'function');
    if (missing.length) return { ok: false, reason: 'missing:' + missing.join(',') };
    // Sample behavior checks.
    const checks = [
        ['z去空白',     ['  hi  '],             'hi'],
        ['z转大写',     ['abc'],                'ABC'],
        ['z转小写',     ['ABC'],                'abc'],
        ['z是否为空',   [''],                   true],
        ['z是否空白',   ['   '],                true],
        ['z是否数字',   ['123'],                true],
        ['z替换',       ['a-b-c', '-', '+'],    'a+b+c'],
        ['z计数',       ['ababab', 'ab'],        3],
        ['z首字母大写', ['hello'],              'Hello'],
        ['z驼峰命名',   ['foo_bar_baz'],         'fooBarBaz'],
        ['z下划线命名', ['FooBarBaz'],           'foo_bar_baz'],
        ['z转义HTML',   ['<a>&"\'</a>'],         '&lt;a&gt;&amp;&quot;&#39;&lt;/a&gt;'],
        ['z反转义HTML', ['&lt;a&gt;'],           '<a>'],
        ['z去除前缀',   ['foobar', 'foo'],       'bar'],
        ['z去除后缀',   ['foobar', 'bar'],       'foo'],
        ['z转数字',     ['1,234.5'],             1234.5],
    ];
    for (const [k, args, want] of checks) {
        const got = StrUtils[k].apply(StrUtils, args);
        if (got !== want) {
            return { ok: false, reason: `mismatch on ${k}(${JSON.stringify(args)}): got ${JSON.stringify(got)} want ${JSON.stringify(want)}` };
        }
    }
    return { ok: true, count: EXPECTED_ALIASES.length };
}

function stripStalePatch() {
    const src = fs.readFileSync(TARGET, 'utf8');
    const re = new RegExp(`\\n?${MARKER_START}[\\s\\S]*?${MARKER_END}\\n`, 'm');
    if (re.test(src)) {
        const stripped = src.replace(re, '\n');
        atomicWrite(stripped);
        return true;
    }
    return false;
}

function run() {
    for (let attempt = 1; attempt <= 8; attempt++) {
        const before = fs.statSync(TARGET);
        const r = patchFile();
        if (r.changed) {
            atomicWrite(r.patched);
            const after = fs.readFileSync(TARGET, 'utf8');
            if (!after.includes(MARKER_START)) {
                console.log(`attempt ${attempt}: write did not stick, retrying`);
                continue;
            }
            console.log(`attempt ${attempt}: patch applied (${before.size} -> ${fs.statSync(TARGET).size})`);
        } else {
            console.log(`attempt ${attempt}: ${r.reason}`);
        }
        const v = verifyInVM();
        if (!v.ok) {
            console.log(`verify failed: ${v.reason}`);
            if (!r.changed && r.reason === 'already-patched') {
                if (stripStalePatch()) {
                    console.log('stripped stale patch, retrying');
                    continue;
                }
            }
            continue;
        }
        console.log(`verify ok: ${v.count} Chinese aliases live and correct`);
        return { ok: true, attempt };
    }
    return { ok: false };
}

const result = run();
if (result.ok) {
    const startSize = fs.statSync(TARGET).size;
    const startMtime = fs.statSync(TARGET).mtimeMs;
    setTimeout(() => {
        const a = fs.statSync(TARGET);
        const stable = a.size === startSize && Math.abs(a.mtimeMs - startMtime) < 5000;
        console.log(`watch: size ${a.size} (was ${startSize}), mtime ${new Date(a.mtimeMs).toISOString()} stable=${stable}`);
        const v = verifyInVM();
        console.log(`re-verify after watch: ok=${v.ok} ${v.reason || ''}`);
        process.exit(v.ok ? 0 : 1);
    }, 15000);
} else {
    process.exit(1);
}
