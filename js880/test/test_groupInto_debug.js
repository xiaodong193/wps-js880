/**
 * 复现 groupInto "结果没分组" 问题 - Node.js 测试
 * 
 * 运行: node js880/test/test_groupInto_debug.js
 */

// ========== 从 JSA880.js 提取核心代码 ==========

// --- Lambda 模式 ---
const LAMBDA_PATTERNS = {
    ARROW_FUNCTION: /=>/,
    INDEX_SELECTOR: /\$(\d+)/g,
    COLUMN_SELECTOR: /^f\d+/,
    MULTI_COLUMN: /^f\d+(\s*,\s*f\d+)+$/
};
const LAMBDA_SAFE_PATTERN = /^[a-zA-Z0-9_\s\(\)\[\]\{\}\,\;\:\+\-\*\/\%\&\|\^\~\<\>\=\!\?\.\$\@\"\'\\\u4e00-\u9fa5]+$/;
const LAMBDA_FORBIDDEN_KEYWORDS = ['__proto__', 'constructor', 'prototype', 'Function', 'eval', 'import', 'require', 'module', 'exports', 'global', 'window', 'document'];
const _lambdaCache = Object.create(null);

function parseLambda(expr) {
    if (typeof expr === 'function') return expr;
    if (typeof expr !== 'string') return null;
    if (expr.length > 500) { console.warn('parseLambda: 表达式过长'); return null; }
    if (!LAMBDA_PATTERNS.ARROW_FUNCTION.test(expr) && !LAMBDA_SAFE_PATTERN.test(expr)) {
        console.warn('parseLambda: 表达式包含不允许的字符: ' + expr);
        return null;
    }
    for (let i = 0; i < LAMBDA_FORBIDDEN_KEYWORDS.length; i++) {
        if (expr.indexOf(LAMBDA_FORBIDDEN_KEYWORDS[i]) >= 0) {
            console.warn('parseLambda: 表达式包含禁止关键词: ' + LAMBDA_FORBIDDEN_KEYWORDS[i]);
            return null;
        }
    }
    if (_lambdaCache[expr]) return _lambdaCache[expr];
    let fn;
    try {
        if (LAMBDA_PATTERNS.ARROW_FUNCTION.test(expr)) {
            fn = eval('(' + expr + ')');
        } else if (expr.includes('$')) {
            const indexMatch = expr.match(LAMBDA_PATTERNS.INDEX_SELECTOR);
            if (indexMatch) {
                const indices = indexMatch.map(m => parseInt(m.substring(1)));
                const maxIndex = Math.max(...indices);
                if (maxIndex > 1000000) { console.warn('Lambda索引超出限制:', maxIndex); return null; }
                fn = new Function('_', 'return ' + expr.replace(LAMBDA_PATTERNS.INDEX_SELECTOR, '_[$1]'));
            }
        } else if (LAMBDA_PATTERNS.MULTI_COLUMN.test(expr)) {
            const cols = expr.split(/\s*,\s*/).map(c => '_[' + (parseInt(c.substring(1)) - 1) + ']').join(',');
            fn = new Function('_', 'return [' + cols + ']');
        } else if (LAMBDA_PATTERNS.COLUMN_SELECTOR.test(expr)) {
            fn = new Function('_', 'return ' + expr.replace(/f(\d+)/g, '_[$1-1]'));
        } else {
            fn = new Function('_', 'return ' + expr);
        }
    } catch (e) {
        console.warn('Lambda解析失败:', expr, e);
        return null;
    }
    _lambdaCache[expr] = fn;
    return fn;
}

// --- Array2D 简化版 ---
function Array2D(data) {
    var items = [];
    if (data === null || data === undefined) {
        items = [];
    } else if (data instanceof Array2D) {
        items = data._items;
    } else if (Array.isArray(data)) {
        items = data;
    } else {
        items = [[data]];
    }
    Array.prototype.push.apply(this, items);
    Object.defineProperty(this, '_original', { value: data, writable: true, enumerable: false, configurable: true });
    Object.defineProperty(this, '_items', {
        get: function() { return Array.prototype.slice.call(this); },
        set: function(value) { Array.prototype.splice.call(this, 0, this.length); Array.prototype.push.apply(this, value); },
        enumerable: false, configurable: true
    });
}
Array2D.prototype = Object.create(Array.prototype);
Array2D.prototype.constructor = Array2D;
Object.defineProperty(Array2D.prototype, 'toJSON', {
    value: function() { return this._items; },
    enumerable: false, configurable: true, writable: true
});

// --- groupInto ---
Array2D.groupInto = function(arr, keySelector, valueSelector, separator) {
    console.log('=== groupInto called ===');
    console.log('arr type:', typeof arr, 'isArray:', Array.isArray(arr), 'instanceof Array2D:', arr instanceof Array2D);
    
    // Handle Array2D instances
    if (arr && arr._items && !Array.isArray(arr)) {
        console.log('  → extracting _items (not an array)');
        arr = arr._items;
    } else if (arr instanceof Array2D) {
        console.log('  → arr is Array2D (which IS an array due to inheritance), NOT extracting _items');
        console.log('  → arr._items exists:', !!arr._items);
        console.log('  → !Array.isArray(arr):', !Array.isArray(arr));
    }
    
    if (!arr || !Array.isArray(arr)) {
        console.log('  → not an array, returning empty');
        return new Array2D([]);
    }
    
    separator = separator || '@^@';
    console.log('separator:', separator);
    
    var keyFn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    console.log('keySelector:', keySelector, '→ keyFn:', typeof keyFn, keyFn ? keyFn.toString() : 'null');
    
    if (!keyFn) {
        console.log('  → keyFn is null, returning empty!');
        return new Array2D([]);
    }

    function _resolveCol(col) {
        if (col === undefined || col === null) return -1;
        if (typeof col === 'number') return col - 1;
        var str = String(col).replace(/^["']|["']$/g, '').replace(/^f/i, '');
        var idx = parseInt(str, 10);
        return isNaN(idx) ? -1 : idx - 1;
    }

    function createAggHelper(rows) {
        // Simplified version for test
        return {
            _rows: rows,
            count: function() { return rows.length; },
            sum: function(col) {
                var idx = _resolveCol(col);
                var s = 0;
                for (var i = 0; i < rows.length; i++) {
                    var v = idx >= 0 ? (Array.isArray(rows[i]) ? rows[i][idx] : rows[i]) : rows[i];
                    var n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
                    if (!isNaN(n)) s += n;
                }
                return s;
            },
            average: function(col) {
                var idx = _resolveCol(col);
                var s = 0, c = 0;
                for (var i = 0; i < rows.length; i++) {
                    var v = idx >= 0 ? (Array.isArray(rows[i]) ? rows[i][idx] : rows[i]) : rows[i];
                    var n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
                    if (!isNaN(n)) { s += n; c++; }
                }
                return c > 0 ? s / c : 0;
            },
            textjoin: function(col, sep) {
                sep = sep !== undefined ? sep : ',';
                var idx = _resolveCol(col);
                if (idx < 0) return '';
                var vals = [];
                for (var i = 0; i < rows.length; i++) {
                    var row = rows[i];
                    var v = Array.isArray(row) && idx < row.length ? row[idx] : '';
                    if (v !== null && v !== undefined && String(v) !== '') vals.push(v);
                }
                return vals.join(sep);
            }
        };
    }

    function parseAggString(str) {
        var parts = [];
        var depth = 0, cur = '';
        for (var i = 0; i < str.length; i++) {
            var c = str[i];
            if (c === '(') depth++;
            else if (c === ')') depth--;
            if (c === ',' && depth === 0) { parts.push(cur.trim()); cur = ''; }
            else { cur += c; }
        }
        if (cur.trim()) parts.push(cur.trim());
        var defs = [];
        for (var p = 0; p < parts.length; p++) {
            var m = parts[p].match(/(sum|count|average|avg|max|min|textjoin|平方和)\s*\(\s*([^)]*)\s*\)/i);
            if (m) {
                var fn = m[1].toLowerCase();
                var argsStr = m[2].trim();
                var args = [];
                if (argsStr) {
                    var inQ = false, curA = '';
                    for (var j = 0; j < argsStr.length; j++) {
                        var ch = argsStr[j];
                        if (ch === '"' || ch === "'") { inQ = !inQ; }
                        else if (ch === ',' && !inQ) { args.push(curA.trim().replace(/^["']|["']$/g, '')); curA = ''; }
                        else { curA += ch; }
                    }
                    if (curA.trim()) args.push(curA.trim().replace(/^["']|["']$/g, ''));
                }
                defs.push({ func: fn, args: args });
            }
        }
        return defs;
    }

    var valueFn;
    if (typeof valueSelector === 'string') {
        var defs = parseAggString(valueSelector);
        console.log('parseAggString result:', JSON.stringify(defs));
        if (defs.length > 0) {
            valueFn = function(rows) {
                var helper = createAggHelper(rows);
                var results = [];
                for (var i = 0; i < defs.length; i++) {
                    var d = defs[i];
                    switch (d.func) {
                        case 'sum': results.push(helper.sum(d.args[0])); break;
                        case 'count': results.push(helper.count()); break;
                        case 'average': case 'avg': results.push(helper.average(d.args[0])); break;
                        case 'textjoin': results.push(helper.textjoin(d.args[0], d.args[1])); break;
                        default: results.push(null);
                    }
                }
                return results.length === 1 ? results[0] : results;
            };
        } else {
            valueFn = parseLambda(valueSelector);
        }
    } else if (typeof valueSelector === 'function') {
        valueFn = function(rows) { var helper = createAggHelper(rows); return valueSelector(helper); };
    } else {
        valueFn = valueSelector;
    }
    
    console.log('valueFn:', typeof valueFn);
    if (!valueFn) { console.log('  → valueFn is null, returning empty!'); return new Array2D([]); }

    // ===== 执行分组 =====
    var groups = Object.create(null);
    console.log('Starting grouping, arr.length:', arr.length);
    for (var i = 0; i < arr.length; i++) {
        var key = keyFn(arr[i], i);
        var keyStr = Array.isArray(key) ? key.join(separator) : String(key);
        console.log('  row', i, '→ key:', JSON.stringify(key), 'keyStr:', keyStr);
        if (!groups[keyStr]) {
            groups[keyStr] = { key: key, rows: [] };
        }
        groups[keyStr].rows.push(arr[i]);
    }
    
    console.log('Number of groups:', Object.keys(groups).length);
    
    // ===== 汇总结果 =====
    var result = [];
    for (var key in groups) {
        var group = groups[key];
        var agg = valueFn(group.rows);
        var row;
        console.log('  group keyStr:', key, 'rows:', group.rows.length, 'agg:', JSON.stringify(agg));
        if (Array.isArray(group.key)) {
            row = Array.isArray(agg) ? group.key.concat(agg) : group.key.concat([agg]);
        } else if (group.key !== null && group.key !== undefined) {
            row = Array.isArray(agg) ? [group.key].concat(agg) : [group.key, agg];
        }
        console.log('  → result row:', JSON.stringify(row));
        result.push(row);
    }
    
    console.log('Final result:', JSON.stringify(result));
    return new Array2D(result);
};

// ========== 测试 ==========
console.log('========================================');
console.log('测试1: 基本 groupInto 调用');
console.log('========================================');

var testData = [
    ['产品A', '北京', 10, 100],
    ['产品A', '上海', 20, 200],
    ['产品B', '北京', 30, 300],
    ['产品B', '上海', 40, 400],
    ['产品A', '北京', 15, 150]
];

console.log('\n--- 测试1a: 字符串key + 字符串聚合 ---');
var rs1 = Array2D.groupInto(testData, 'f2,f3', 'count(),sum("f4"),average("f5"),textjoin("f4","+")');
console.log('Result type:', rs1 instanceof Array2D ? 'Array2D' : typeof rs1);
console.log('Result length:', rs1.length);
console.log('Result:', JSON.stringify(rs1));
console.log('Expected: 4 groups (产品A-北京, 产品A-上海, 产品B-北京, 产品B-上海)\n');

console.log('--- 测试1b: 模拟 WPS Array2D 输入 (包装在 Array2D 中) ---');
var arr2d = new Array2D(testData);
console.log('arr2d is Array2D:', arr2d instanceof Array2D);
console.log('Array.isArray(arr2d):', Array.isArray(arr2d));
var rs1b = Array2D.groupInto(arr2d, 'f2,f3', 'count(),sum("f4"),average("f5")');
console.log('Result:', JSON.stringify(rs1b));
console.log('');

// 测试关键问题：parseLambda('f2,f3') 返回什么？
console.log('========================================');
console.log('测试2: 诊断 parseLambda("f2,f3")');
console.log('========================================');
console.log('MULTI_COLUMN pattern:', LAMBDA_PATTERNS.MULTI_COLUMN);
console.log('"f2,f3" matches MULTI_COLUMN:', LAMBDA_PATTERNS.MULTI_COLUMN.test('f2,f3'));
console.log('"f2,f3" matches SAFE_PATTERN:', LAMBDA_SAFE_PATTERN.test('f2,f3'));
console.log('"f2,f3" matches COLUMN_SELECTOR:', LAMBDA_PATTERNS.COLUMN_SELECTOR.test('f2,f3'));
console.log('"f2,f3" includes "$":', 'f2,f3'.includes('$'));
console.log('"f2,f3" matches ARROW_FUNCTION:', LAMBDA_PATTERNS.ARROW_FUNCTION.test('f2,f3'));

var keyFn = parseLambda('f2,f3');
console.log('parseLambda("f2,f3") result:', keyFn);
if (keyFn) {
    console.log('keyFn.toString():', keyFn.toString());
    var testRow = ['产品A', '北京', 10, 100];
    var keyResult = keyFn(testRow, 0);
    console.log('keyFn(["产品A","北京",10,100]):', JSON.stringify(keyResult));
}

// 测试空数组
console.log('\n========================================');
console.log('测试3: 空数组保护');
console.log('========================================');
var rs3 = Array2D.groupInto([], 'f2,f3', 'count(),sum("f4")');
console.log('Empty result type:', rs3 instanceof Array2D ? 'Array2D' : typeof rs3);
console.log('Empty result length:', rs3.length);

// 测试单字段分组
console.log('\n========================================');
console.log('测试4: 单字段分组 f1');
console.log('========================================');
var rs4 = Array2D.groupInto(testData, 'f1', 'count(),sum("f3")');
console.log('Result:', JSON.stringify(rs4));

console.log('\n========================================');
console.log('测试完成');
console.log('========================================');