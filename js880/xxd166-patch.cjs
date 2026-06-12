#!/usr/bin/env node
// XXD-166: Atomic patch script for NumUtils Chinese aliases
// Uses the CTO direct-patch pattern: read → splice → atomic rename in a loop
// to beat the Synology+WPS+iCloud writer race.

const fs = require('fs');
const path = require('path');

const TARGET = '/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880/JSA880.js';
const MARKER = '// XXD-166-PATCH-MARKER';

const NEW_NUMUTILS = `var NumUtils = {
    round: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return 0;
        if (decimals === undefined || decimals === null) decimals = 0;
        var factor = Math.pow(10, decimals);
        return Math.round(num * factor) / factor;
    },
    ceil: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return 0;
        if (decimals === undefined || decimals === null) decimals = 0;
        var factor = Math.pow(10, decimals);
        return Math.ceil(num * factor) / factor;
    },
    floor: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return 0;
        if (decimals === undefined || decimals === null) decimals = 0;
        var factor = Math.pow(10, decimals);
        return Math.floor(num * factor) / factor;
    },
    abs: function(num) { return typeof num === "number" && !isNaN(num) ? Math.abs(num) : 0; },
    sign: function(num) { if (typeof num !== "number" || isNaN(num)) return 0; return num > 0 ? 1 : (num < 0 ? -1 : 0); },
    formatNumber: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return "0";
        if (decimals === undefined || decimals === null) decimals = 0;
        var fixed = this.round(num, decimals).toString();
        var parts = fixed.split(".");
        var intPart = parts[0];
        var result = "";
        var count = 0;
        for (var i = intPart.length - 1; i >= 0; i--) {
            if (count > 0 && count % 3 === 0 && intPart.charAt(i) !== "-") result = "," + result;
            result = intPart.charAt(i) + result; count++;
        }
        if (parts.length > 1) {
            var decStr = parts[1];
            while (decStr.length < decimals) decStr += "0";
            result += "." + decStr;
        } else if (decimals > 0) {
            result += "." + this._repeatChar("0", decimals);
        }
        return result;
    },
    formatCurrency: function(num, currencySymbol, decimals) {
        if (typeof num !== "number" || isNaN(num)) return "0.00";
        currencySymbol = currencySymbol || "¥";
        if (decimals === undefined || decimals === null) decimals = 2;
        return currencySymbol + this.formatNumber(num, decimals);
    },
    formatPercent: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return "0%";
        if (decimals === undefined || decimals === null) decimals = 0;
        var percent = num * 100;
        var formatted = this.round(percent, decimals).toString();
        if (decimals > 0 && formatted.indexOf(".") === -1) formatted += "." + this._repeatChar("0", decimals);
        return formatted + "%";
    },
    parse: function(str, defaultValue) {
        if (str === null || str === undefined) return defaultValue !== undefined ? defaultValue : 0;
        if (typeof str === "number") return isNaN(str) ? (defaultValue !== undefined ? defaultValue : 0) : str;
        var numStr = String(str).replace(/[^0-9.\\-]/g, "");
        var num = parseFloat(numStr);
        return isNaN(num) ? (defaultValue !== undefined ? defaultValue : 0) : num;
    },
    clamp: function(num, min, max) {
        if (typeof num !== "number" || isNaN(num)) num = 0;
        if (num < min) return min;
        if (num > max) return max;
        return num;
    },
    inRange: function(num, min, max) { return typeof num === "number" && !isNaN(num) && num >= min && num <= max; },
    sum: function(nums) {
        var args = Array.prototype.slice.call(arguments);
        var total = 0;
        for (var i = 0; i < args.length; i++) {
            if (typeof args[i] === "number" && !isNaN(args[i])) total += args[i];
            else if (args[i] instanceof Array) {
                for (var j = 0; j < args[i].length; j++) {
                    if (typeof args[i][j] === "number" && !isNaN(args[i][j])) total += args[i][j];
                }
            }
        }
        return total;
    },
    avg: function(nums) {
        var args = Array.prototype.slice.call(arguments);
        var values = [];
        for (var i = 0; i < args.length; i++) {
            if (typeof args[i] === "number" && !isNaN(args[i])) values.push(args[i]);
            else if (args[i] instanceof Array) {
                for (var j = 0; j < args[i].length; j++) {
                    if (typeof args[i][j] === "number" && !isNaN(args[i][j])) values.push(args[i][j]);
                }
            }
        }
        return values.length === 0 ? 0 : this.sum(values) / values.length;
    },
    max: function(nums) {
        var args = Array.prototype.slice.call(arguments);
        var values = [];
        for (var i = 0; i < args.length; i++) {
            if (typeof args[i] === "number" && !isNaN(args[i])) values.push(args[i]);
            else if (args[i] instanceof Array) {
                for (var j = 0; j < args[i].length; j++) {
                    if (typeof args[i][j] === "number" && !isNaN(args[i][j])) values.push(args[i][j]);
                }
            }
        }
        return values.length === 0 ? 0 : Math.max.apply(Math, values);
    },
    min: function(nums) {
        var args = Array.prototype.slice.call(arguments);
        var values = [];
        for (var i = 0; i < args.length; i++) {
            if (typeof args[i] === "number" && !isNaN(args[i])) values.push(args[i]);
            else if (args[i] instanceof Array) {
                for (var j = 0; j < args[i].length; j++) {
                    if (typeof args[i][j] === "number" && !isNaN(args[i][j])) values.push(args[i][j]);
                }
            }
        }
        return values.length === 0 ? 0 : Math.min.apply(Math, values);
    },
    randomInt: function(min, max) {
        if (max === undefined) { max = min; min = 0; }
        return Math.floor(Math.random() * (max - min + 1)) + min;
    },
    random: function(min, max) {
        if (max === undefined) { max = min; min = 0; }
        return Math.random() * (max - min) + min;
    },
    randomId: function(length) {
        if (typeof length !== "number" || length < 1) length = 8;
        var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        var result = "";
        for (var i = 0; i < length; i++) result += chars.charAt(Math.floor(Math.random() * chars.length));
        return result;
    },
    repeat: function(value, count) {
        if (typeof count !== "number" || count < 0) return [];
        var result = [];
        for (var i = 0; i < count; i++) result.push(value);
        return result;
    },
    _repeatChar: function(ch, count) {
        var result = "";
        for (var i = 0; i < count; i++) result += ch;
        return result;
    },
    isNumber: function(val) { return typeof val === "number" && !isNaN(val); },
    isInteger: function(val) { return typeof val === "number" && !isNaN(val) && val === Math.floor(val); },
    isEven: function(num) { return typeof num === "number" && !isNaN(num) && num % 2 === 0; },
    isOdd: function(num) { return typeof num === "number" && !isNaN(num) && num % 2 !== 0; },
    isPositive: function(num) { return typeof num === "number" && !isNaN(num) && num > 0; },
    isNegative: function(num) { return typeof num === "number" && !isNaN(num) && num < 0; },
    // === XXD-166: 中文 alias (大包装, 与 Array2D 风格一致) ===
    z取整: function(num, decimals) { return this.round(num, decimals); },
    z向上取整: function(num, decimals) { return this.ceil(num, decimals); },
    z向下取整: function(num, decimals) { return this.floor(num, decimals); },
    z四舍五入: function(num, decimals) { return this.round(num, decimals); },
    z绝对值: function(num) { return this.abs(num); },
    z幂: function(base, exp) { return Math.pow(base, exp); },
    z对数: function(n) { return Math.log(n); },
    z正弦: function(n) { return Math.sin(n); },
    z余弦: function(n) { return Math.cos(n); },
    z正切: function(n) { return Math.tan(n); },
    z取小数: function(num, decimals) { return this.round(num, decimals); },
    z百分比: function(num, decimals) { return this.formatPercent(num, decimals); },
    z随机数: function(min, max) { return this.random(min, max); },
    z随机整数: function(min, max) { return this.randomInt(min, max); },
    z随机ID: function(length) { return this.randomId(length); },
    z是数字: function(val) { return this.isNumber(val); },
    z是整数: function(val) { return this.isInteger(val); },
    z是偶数: function(num) { return this.isEven(num); },
    z是奇数: function(num) { return this.isOdd(num); },
    z是正数: function(num) { return this.isPositive(num); },
    z是负数: function(num) { return this.isNegative(num); },
    z是NaN: function(val) { return isNaN(val); },
    z裁剪: function(num, min, max) { return this.clamp(num, min, max); },
    z在范围内: function(num, min, max) { return this.inRange(num, min, max); },
    z格式化货币: function(num, currencySymbol, decimals) { return this.formatCurrency(num, currencySymbol, decimals); },
    z格式化数字: function(num, decimals) { return this.formatNumber(num, decimals); },
    z格式化百分比: function(num, decimals) { return this.formatPercent(num, decimals); },
    z求和: function() { return this.sum.apply(this, arguments); },
    z平均: function() { return this.avg.apply(this, arguments); },
    z最大值: function() { return this.max.apply(this, arguments); },
    z最小值: function() { return this.min.apply(this, arguments); },
    z重复: function(value, count) { return this.repeat(value, count); },
    z解析: function(str, defaultValue) { return this.parse(str, defaultValue); },
    z符号: function(num) { return this.sign(num); }
};`;

function findNumUtilsBlock(src) {
    const startMarker = 'var NumUtils = {';
    const endMarker = '};';
    const startIdx = src.indexOf(startMarker);
    if (startIdx === -1) return null;
    let depth = 0;
    let inString = null;
    let escape = false;
    for (let i = startIdx; i < src.length; i++) {
        const ch = src[i];
        if (escape) { escape = false; continue; }
        if (inString) {
            if (ch === '\\\\') { escape = true; continue; }
            if (ch === inString) inString = null;
            continue;
        }
        if (ch === '"' || ch === "'") { inString = ch; continue; }
        if (ch === '/' && src[i+1] === '/') {
            while (i < src.length && src[i] !== '\\n') i++;
            continue;
        }
        if (ch === '{') depth++;
        else if (ch === '}') {
            depth--;
            if (depth === 0) {
                return { start: startIdx, end: i + 2 };
            }
        }
    }
    return null;
}

function applyPatch() {
    const original = fs.readFileSync(TARGET, 'utf8');
    const range = findNumUtilsBlock(original);
    if (!range) throw new Error('NumUtils block not found');

    const head = original.substring(0, range.start);
    const tail = original.substring(range.end);
    const patched = head + NEW_NUMUTILS + tail;

    const tmp = TARGET + '.xxd166.tmp';
    fs.writeFileSync(tmp, patched, 'utf8');
    fs.renameSync(tmp, TARGET);
    return patched.length;
}

function verify() {
    const { execSync } = require('child_process');
    const script = `var fs=require('fs');eval(fs.readFileSync(${JSON.stringify(TARGET)},'utf8'));var n=NumUtils;var keys=['z取整','z向上取整','z向下取整','z四舍五入','z绝对值','z幂','z对数','z正弦','z余弦','z正切','z取小数','z百分比','z随机数','z随机整数','z随机ID','z是数字','z是整数','z是偶数','z是奇数','z是正数','z是负数','z是NaN','z裁剪','z在范围内','z格式化货币','z格式化数字','z格式化百分比','z求和','z平均','z最大值','z最小值','z重复','z解析','z符号'];var missing=keys.filter(function(k){return typeof n[k]!=='function';});console.log('total_aliases='+keys.length,'missing='+missing.length, missing.length?'missing='+missing.join(','):'');console.log('z取整(1.5)='+n.z取整(1.5));console.log('z绝对值(-3)='+n.z绝对值(-3));console.log('z向上取整(1.2)='+n.z向上取整(1.2));console.log('z求和(1,2,3)='+n.z求和(1,2,3));console.log('z格式化货币(1234.5)='+n.z格式化货币(1234.5));console.log('z是数字(123)='+n.z是数字(123));console.log('z随机ID='+n.z随机ID());`;
    const out = execSync('node -e ' + JSON.stringify(script), { encoding: 'utf8' });
    return out;
}

// Atomic write loop: Synology+WPS+iCloud overwrite the file every few seconds.
// Re-read, re-patch, re-write until the marker is present and the file stays
// stable for one verification round.
let attempts = 0;
const MAX = 8;
let lastSize = 0;
let stable = 0;
while (attempts < MAX) {
    attempts++;
    let src;
    try { src = fs.readFileSync(TARGET, 'utf8'); } catch (e) { continue; }

    if (src.indexOf(MARKER) !== -1 && src.indexOf('z取整:') !== -1 && src.indexOf('z格式化货币:') !== -1) {
        // already patched
        if (src.length === lastSize) stable++;
        else { stable = 0; lastSize = src.length; }
        if (stable >= 1) {
            console.log('Patch already present and stable (attempt ' + attempts + ')');
            break;
        }
        continue;
    }

    // not patched yet — apply
    try {
        const size = applyPatch();
        console.log('Patch applied (attempt ' + attempts + '), size=' + size);
    } catch (e) {
        console.error('Patch failed attempt ' + attempts + ': ' + e.message);
        // wait briefly for the racing writer to settle
        const wait = 200 + Math.floor(Math.random() * 300);
        const until = Date.now() + wait;
        while (Date.now() < until) {}
        continue;
    }
}

console.log('--- VERIFY ---');
try {
    const result = verify();
    console.log(result);
} catch (e) {
    console.error('Verify failed: ' + e.message);
    process.exit(2);
}
