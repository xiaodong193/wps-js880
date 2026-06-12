// XXD-180 / XXD-181: SuperMap.z分组 vs SuperMap.z分组统计 return-shape documentation
// Atomic-write loop patcher that beats the Synology+WPS+iCloud writer race.
// Marks a stable start (XXD-180 final fix start) and end marker after both edits.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const TARGET = path.join('/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880', 'JSA880.js');
const TMP    = TARGET + '.xxd180.tmp';

const START_MARKER  = '// 🔧 XXD-180/XXD-181 final fix start: SuperMap.z分组 vs z分组统计 形状区分文档化';
const END_MARKER    = '// 🔧 XXD-180/XXD-181 final fix end';
const Z_GROUP_LINE  = 'SuperMap.z分组 = function(arr, sel) { return new Array2D(arr).z分组(sel); };';
const Z_STATS_BEFORE = '/**\n * SuperMap.z分组统计 - 按列分组并统计\n */\n(?:\n// 🔧 XXD-180/XXD-181 final fix end\n)?SuperMap\\.z分组统计 = function\\(arr, groupCol, statsConfig\\) \\{';

// Replacement blocks.
const Z_GROUP_REPLACEMENT =
    '// 🔧 XXD-154 final fix: SuperMap 全套数据操作 alias (= Array2D)\n' +
    'SuperMap.z求和 = function(arr, sel) { return new Array2D(arr).z求和(sel); };\n' +
    'SuperMap.z计数 = function(arr, sel) { return new Array2D(arr).z计数(sel); };\n' +
    'SuperMap.z最大值 = function(arr, sel) { return new Array2D(arr).z最大值(sel); };\n' +
    'SuperMap.z最小值 = function(arr, sel) { return new Array2D(arr).z最小值(sel); };\n' +
    'SuperMap.z平均 = function(arr, sel) { return new Array2D(arr).z平均值(sel); };\n' +
    'SuperMap.z去重 = function(arr, sel) { return new Array2D(arr).z去重(sel); };\n' +
    '/**\n' +
    ' * SuperMap.z分组 - 纯分组 (无聚合)\n' +
    ' *\n' +
    ' * @param {Array2D|Array} arr  - 源二维数组 (含 header 行)\n' +
    ' * @param {Function|string|number|Array} sel - key 选择器,同 Array2D.prototype.z分组\n' +
    ' * @returns {Object<string, Array<Array>>} 字典: { groupKey: [row, row, ...] }\n' +
    ' *\n' +
    ' * 形状约定 (与 z分组统计 区分):\n' +
    ' *   - z分组:      返回 字典 { key: rows[] }    (纯分组,无聚合)\n' +
    ' *   - z分组统计:  返回 二维数组 [header, ...]   (分组+聚合,每组 1 行)\n' +
    ' * 调用方按需选择:\n' +
    ' *   - 需要 key→行集合  → 用 z分组\n' +
    ' *   - 需要每组 1 行汇总 → 用 z分组统计\n' +
    ' */\n' +
    'SuperMap.z分组 = function(arr, sel) { return new Array2D(arr).z分组(sel); };';

const Z_STATS_REPLACEMENT =
    '/**\n' +
    ' * SuperMap.z分组统计 - 按列分组并统计 (分组 + 聚合, 每组 1 行)\n' +
    ' *\n' +
    ' * @param {Array2D|Array} arr  - 源二维数组 (含 header 行)\n' +
    ' * @param {string|number}  groupCol    - 分组列 (f1 风格或 0 基索引)\n' +
    ' * @param {Object<string,string>} statsConfig - 聚合配置 { colName: aggType }\n' +
    ' *                                              aggType ∈ sum|avg|count|min|max\n' +
    ' * @returns {Array<Array>} 二维数组,首行是表头 [groupHeader, "colA_sum", "colB_avg", ...]\n' +
    ' *                          后续每行 [groupKey, aggValA, aggValB, ...]\n' +
    ' *\n' +
    ' * 形状约定 (与 z分组 区分):\n' +
    ' *   - z分组:      返回 字典 { key: rows[] }    (纯分组,无聚合)\n' +
    ' *   - z分组统计:  返回 二维数组 [header, ...]   (分组+聚合,每组 1 行)\n' +
    ' * 调用方按需选择:\n' +
    ' *   - 需要 key→行集合  → 用 z分组\n' +
    ' *   - 需要每组 1 行汇总 → 用 z分组统计\n' +
    ' */\n' +
    'SuperMap.z分组统计 = function(arr, groupCol, statsConfig) {';

function buildReplacements() {
    // v4.0.39 changelog block, inserted right after the line
    //   "* 更新日志 (v4.0.38 — 2026-06-11)"
    // and before the next "------------------------------------------------------------------------"
    // separator that follows the v4.0.38 content.
    const v38 = [
        ' * 更新日志 (v4.0.38 — 2026-06-11)',
    ];
    // We anchor on the v4.0.38 header line + the existing "------------------------------------------------------------------------" line
    // that comes right before the v4.0.37 block, and insert the v4.0.39 block in between.
    return [
        {
            id: 'z_group_jsdoc',
            match: new RegExp(
                '(/\\*\\*\\n' +
                ' \\* SuperMap\\.z选择列 - 选择指定的列\\n' +
                ' \\*/\\n)' +
                '// 🔧 XXD-154 final fix: SuperMap 全套数据操作 alias \\(= Array2D\\)\\n' +
                'SuperMap\\.z求和 = function\\(arr, sel\\) \\{ return new Array2D\\(arr\\)\\.z求和\\(sel\\); \\};\\n' +
                'SuperMap\\.z计数 = function\\(arr, sel\\) \\{ return new Array2D\\(arr\\)\\.z计数\\(sel\\); \\};\\n' +
                'SuperMap\\.z最大值 = function\\(arr, sel\\) \\{ return new Array2D\\(arr\\)\\.z最大值\\(sel\\); \\};\\n' +
                'SuperMap\\.z最小值 = function\\(arr, sel\\) \\{ return new Array2D\\(arr\\)\\.z最小值\\(sel\\); \\};\\n' +
                'SuperMap\\.z平均 = function\\(arr, sel\\) \\{ return new Array2D\\(arr\\)\\.z平均值\\(sel\\); \\};\\n' +
                'SuperMap\\.z去重 = function\\(arr, sel\\) \\{ return new Array2D\\(arr\\)\\.z去重\\(sel\\); \\};\\n' +
                'SuperMap\\.z分组 = function\\(arr, sel\\) \\{ return new Array2D\\(arr\\)\\.z分组\\(sel\\); \\};'
            ),
            build: () =>
                '/**\n' +
                ' * SuperMap.z选择列 - 选择指定的列\n' +
                ' */\n' +
                '// 🔧 XXD-154 final fix: SuperMap 全套数据操作 alias (= Array2D)\n' +
                'SuperMap.z求和 = function(arr, sel) { return new Array2D(arr).z求和(sel); };\n' +
                'SuperMap.z计数 = function(arr, sel) { return new Array2D(arr).z计数(sel); };\n' +
                'SuperMap.z最大值 = function(arr, sel) { return new Array2D(arr).z最大值(sel); };\n' +
                'SuperMap.z最小值 = function(arr, sel) { return new Array2D(arr).z最小值(sel); };\n' +
                'SuperMap.z平均 = function(arr, sel) { return new Array2D(arr).z平均值(sel); };\n' +
                'SuperMap.z去重 = function(arr, sel) { return new Array2D(arr).z去重(sel); };\n' +
                START_MARKER + '\n' +
                '/**\n' +
                ' * SuperMap.z分组 - 纯分组 (无聚合)\n' +
                ' *\n' +
                ' * @param {Array2D|Array} arr  - 源二维数组 (含 header 行)\n' +
                ' * @param {Function|string|number|Array} sel - key 选择器,同 Array2D.prototype.z分组\n' +
                ' * @returns {Object<string, Array<Array>>} 字典: { groupKey: [row, row, ...] }\n' +
                ' *\n' +
                ' * 形状约定 (与 z分组统计 区分):\n' +
                ' *   - z分组:      返回 字典 { key: rows[] }    (纯分组,无聚合)\n' +
                ' *   - z分组统计:  返回 二维数组 [header, ...]   (分组+聚合,每组 1 行)\n' +
                ' * 调用方按需选择:\n' +
                ' *   - 需要 key→行集合  → 用 z分组\n' +
                ' *   - 需要每组 1 行汇总 → 用 z分组统计\n' +
                ' */\n' +
                'SuperMap.z分组 = function(arr, sel) { return new Array2D(arr).z分组(sel); };'
        },
        {
            id: 'z_stats_jsdoc',
            match: /\/\*\*\n \* SuperMap\.z分组统计 - 按列分组并统计\n \*\/\n(?:\n\/\/ 🔧 XXD-180\/XXD-181 final fix end\n)?SuperMap\.z分组统计 = function\(arr, groupCol, statsConfig\) \{/,
            build: () =>
                '/**\n' +
                ' * SuperMap.z分组统计 - 按列分组并统计 (分组 + 聚合, 每组 1 行)\n' +
                ' *\n' +
                ' * @param {Array2D|Array} arr  - 源二维数组 (含 header 行)\n' +
                ' * @param {string|number}  groupCol    - 分组列 (f1 风格或 0 基索引)\n' +
                ' * @param {Object<string,string>} statsConfig - 聚合配置 { colName: aggType }\n' +
                ' *                                              aggType ∈ sum|avg|count|min|max\n' +
                ' * @returns {Array<Array>} 二维数组,首行是表头 [groupHeader, "colA_sum", "colB_avg", ...]\n' +
                ' *                          后续每行 [groupKey, aggValA, aggValB, ...]\n' +
                ' *\n' +
                ' * 形状约定 (与 z分组 区分):\n' +
                ' *   - z分组:      返回 字典 { key: rows[] }    (纯分组,无聚合)\n' +
                ' *   - z分组统计:  返回 二维数组 [header, ...]   (分组+聚合,每组 1 行)\n' +
                ' * 调用方按需选择:\n' +
                ' *   - 需要 key→行集合  → 用 z分组\n' +
                ' *   - 需要每组 1 行汇总 → 用 z分组统计\n' +
                ' */\n' +
                'SuperMap.z分组统计 = function(arr, groupCol, statsConfig) {'
        },
        {
            id: 'changelog_v4039',
            // Insert v4.0.39 changelog block right before the "v4.0.38" header line.
            // We anchor on the "v4.0.38" line and prepend the new block.
            match: /(\* 更新日志 \(v4\.0\.38 — 2026-06-11\))/,
            build: () =>
                ' * 更新日志 (v4.0.39 — 2026-06-11)\n' +
                ' * --------------------------------------------------------------------------\n' +
                ' * 1. [文档] XXD-180/XXD-181: SuperMap.z分组 / z分组统计 形状约定澄清\n' +
                ' *    - 现象: 两者同名 z分组 但返回形状不同,易被误以为是 bug\n' +
                ' *      · SuperMap.z分组(arr, sel)                  → 字典 { key: rows[] }\n' +
                ' *      · SuperMap.z分组统计(arr, groupCol, cfg)    → 二维数组 [header, ...rows]\n' +
                ' *    - 判定: 不是 bug,是不同操作的不同自然表示\n' +
                ' *      · z分组      是纯分组 (无聚合),字典 key→行集合 最自然\n' +
                ' *      · z分组统计  是分组+聚合,每组 1 行,二维表格最自然\n' +
                ' *    - 修法: 在两个方法上方补 JSDoc 明确写出 @returns 形状 + 形状对照说明\n' +
                ' *    - 行为: 无运行时代码变更;两种形状保持向后兼容\n' +
                ' *\n' +
                ' * 更新日志 (v4.0.38 — 2026-06-11)'
        }
    ];
}

function applyOnce(repls, markerStart, markerEnd) {
    let text = fs.readFileSync(TARGET, 'utf8');
    // Always strip any pre-existing v4.0.39 changelog block(s) — the changelog_v4039
    // replacement below will insert exactly one canonical copy. Header line may have
    // 1 or 2 leading spaces due to historical patch quirks. Non-greedy match: stop at
    // the first ' *\n' closing line before the next changelog header.
    const blockRe = / {1,2}\* 更新日志 \(v4\.0\.39 — 2026-06-11\)\n \* -{50,}\n \* 1\. \[文档\] XXD-180\/XXD-181:[\s\S]*?\n \*\n/g;
    text = text.replace(blockRe, '');
    // First, idempotency guard: if the new z分组统计 JSDoc is missing but the
    // marker+start block is present, the writer race reverted only the second
    // JSDoc. Re-insert it without disturbing the first one.
    if (text.includes(markerStart) && !text.includes('SuperMap.z分组统计 - 按列分组并统计 (分组 + 聚合')) {
        const statsBefore1 = '/**\n * SuperMap.z分组统计 - 按列分组并统计\n */\nSuperMap.z分组统计 = function(arr, groupCol, statsConfig) {';
        const statsBefore2 = '/**\n * SuperMap.z分组统计 - 按列分组并统计\n */\n// 🔧 XXD-180/XXD-181 final fix end\nSuperMap.z分组统计 = function(arr, groupCol, statsConfig) {';
        const newStats = '/**\n' +
            ' * SuperMap.z分组统计 - 按列分组并统计 (分组 + 聚合, 每组 1 行)\n' +
            ' *\n' +
            ' * @param {Array2D|Array} arr  - 源二维数组 (含 header 行)\n' +
            ' * @param {string|number}  groupCol    - 分组列 (f1 风格或 0 基索引)\n' +
            ' * @param {Object<string,string>} statsConfig - 聚合配置 { colName: aggType }\n' +
            ' *                                              aggType ∈ sum|avg|count|min|max\n' +
            ' * @returns {Array<Array>} 二维数组,首行是表头 [groupHeader, "colA_sum", "colB_avg", ...]\n' +
            ' *                          后续每行 [groupKey, aggValA, aggValB, ...]\n' +
            ' *\n' +
            ' * 形状约定 (与 z分组 区分):\n' +
            ' *   - z分组:      返回 字典 { key: rows[] }    (纯分组,无聚合)\n' +
            ' *   - z分组统计:  返回 二维数组 [header, ...]   (分组+聚合,每组 1 行)\n' +
            ' * 调用方按需选择:\n' +
            ' *   - 需要 key→行集合  → 用 z分组\n' +
            ' *   - 需要每组 1 行汇总 → 用 z分组统计\n' +
            ' */\n' +
            'SuperMap.z分组统计 = function(arr, groupCol, statsConfig) {';
        if (text.includes(statsBefore2)) {
            text = text.replace(statsBefore2, newStats);
        } else if (text.includes(statsBefore1)) {
            text = text.replace(statsBefore1, newStats);
        }
    }
    // Apply the rest of the replacements only if the pattern is present.
    for (const r of repls) {
        const m = text.match(r.match);
        if (!m) {
            // Idempotent skip: edits may already be in place after a partial revert.
            if (r.id === 'z_group_jsdoc' && text.includes('SuperMap.z分组 - 纯分组 (无聚合')) continue;
            if (r.id === 'z_stats_jsdoc' && text.includes('SuperMap.z分组统计 - 按列分组并统计 (分组 + 聚合')) continue;
            if (r.id === 'changelog_v4039') {
                // Detect duplicates and dedupe by removing all v4.0.39 blocks then
                // re-applying the canonical one.
                if (text.includes('更新日志 (v4.0.39 — 2026-06-11)')) {
                    const dupCount = (text.match(/更新日志 \(v4\.0\.39 — 2026-06-11\)/g) || []).length;
                    if (dupCount > 1) {
                        // Remove ALL v4.0.39 blocks (greedy, multi-line), then continue the
                        // loop body — `continue` will skip this iteration; we'll re-apply
                        // canonical block on next pass via the `r.match` below.
                        // The simpler path: just rewrite the file with the canonical block
                        // inserted at the right anchor. Here, we just continue without
                        // re-applying, since the next loop iteration's match anchor
                        // `v4.0.38` is still present.
                    }
                    continue;
                }
            }
            throw new Error(`[xxd180] pattern miss for id=${r.id}; file may have shifted.`);
        }
        text = text.replace(r.match, r.build());
    }
    // Wrap the patched block with start/end markers (idempotent).
    if (!text.includes(markerStart)) {
        text = text.replace(
            START_MARKER + '\n/**\n * SuperMap.z分组 - 纯分组 (无聚合)',
            markerStart + '\n/**\n * SuperMap.z分组 - 纯分组 (无聚合)'
        );
    }
    if (!text.includes(markerEnd)) {
        text = text.replace(
            'SuperMap.z分组统计 = function(arr, groupCol, statsConfig) {',
            markerEnd + '\nSuperMap.z分组统计 = function(arr, groupCol, statsConfig) {'
        );
    }
    fs.writeFileSync(TMP, text, 'utf8');
    fs.renameSync(TMP, TARGET);
    return text;
}

function verifyShape() {
    const code = fs.readFileSync(TARGET, 'utf8');
    const ctx = { console, globalThis: {} };
    ctx.global = ctx;
    vm.createContext(ctx);
    vm.runInContext(code, ctx);
    const SM = ctx.SuperMap;
    const arr = [['A','B'],[1,2],[1,2],[3,4]];
    const g = SM.z分组(arr, 0);
    const s = SM.z分组统计(arr, 0, 1);
    const okDict = (typeof g === 'object' && !Array.isArray(g)
                    && Array.isArray(g['1']) && Array.isArray(g['3']));
    const okArr  = (Array.isArray(s) && Array.isArray(s[0]) && s[0][0] === 'A'
                    && s.length === 3 && s[1][0] === '1' && s[2][0] === '3');
    return { okDict, okArr, g, s };
}

function main() {
    const repls = buildReplacements();
    let attempt = 0;
    let lastErr = null;
    while (attempt < 8) {
        attempt++;
        try {
            applyOnce(repls, START_MARKER, END_MARKER);
        } catch (e) {
            lastErr = e;
            console.error(`[xxd180] apply attempt ${attempt} failed: ${e.message}`);
            // File may have been overwritten by writer race — refetch pattern and retry.
            continue;
        }
        // Verify markers stuck.
        const text = fs.readFileSync(TARGET, 'utf8');
        if (!text.includes(START_MARKER) || !text.includes(END_MARKER)) {
            console.error(`[xxd180] attempt ${attempt}: markers missing after write`);
            continue;
        }
        // Runtime verify.
        const v = verifyShape();
        if (!v.okDict || !v.okArr) {
            console.error(`[xxd180] attempt ${attempt}: runtime verify FAILED`,
                          JSON.stringify(v));
            continue;
        }
        // Watch the file for 10s — if writer race reverts, re-patch.
        const size = fs.statSync(TARGET).size;
        let raceHit = false;
        for (let s = 0; s < 10; s++) {
            Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 1000);
            const now = fs.statSync(TARGET);
            if (now.size !== size || !fs.readFileSync(TARGET, 'utf8').includes(START_MARKER)) {
                raceHit = true;
                break;
            }
        }
        if (raceHit) {
            console.error(`[xxd180] attempt ${attempt}: writer race reverted, re-patching`);
            continue;
        }
        console.log(`[xxd180] OK on attempt ${attempt}`);
        console.log(`[xxd180] z分组  → dict keys: ${Object.keys(v.g).join(',')}`);
        console.log(`[xxd180] z分组统计 → 2D array: ${JSON.stringify(v.s)}`);
        return 0;
    }
    console.error(`[xxd180] FAILED after ${attempt} attempts; last err: ${lastErr && lastErr.message}`);
    return 1;
}

process.exit(main());
