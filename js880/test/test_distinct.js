/**
 * 测试 Array2D.distinct 功能
 */

// 模拟 Array2D 类
class Array2D {
    constructor(arr) {
        this._items = Array.isArray(arr) ? arr : [];
    }
    _new(arr) { return new Array2D(arr); }
    val() { return this._items; }

    // 核心去重方法
    z去重(colSelector, resultSelector) {
        var seen = Object.create(null);
        var result = [];

        var keyFn;
        var isFunctionMode = false;
        var colIndexes = [];

        if (colSelector === undefined) {
            keyFn = function(row) { return JSON.stringify(row); };
        } else if (typeof colSelector === 'function') {
            keyFn = colSelector;
            isFunctionMode = true;
        } else if (typeof colSelector === 'number') {
            keyFn = function(row) { return row[colSelector]; };
        } else if (typeof colSelector === 'string') {
            if (colSelector.includes(',')) {
                var parts = colSelector.split(',');
                for (var p = 0; p < parts.length; p++) {
                    var part = parts[p].trim();
                    if (part.toLowerCase().startsWith('f')) {
                        colIndexes.push(parseInt(part.substring(1)) - 1);
                    } else if (part.includes('-')) {
                        var range = part.split('-');
                        var start = parseInt(range[0].toLowerCase().replace('f', ''));
                        var end = parseInt(range[1].toLowerCase().replace('f', ''));
                        for (var r = start; r <= end; r++) {
                            colIndexes.push(r - 1);
                        }
                    } else {
                        colIndexes.push(parseInt(part) - 1);
                    }
                }
                keyFn = function(row) {
                    var keyParts = [];
                    for (var i = 0; i < colIndexes.length; i++) {
                        keyParts.push(row[colIndexes[i]]);
                    }
                    return JSON.stringify(keyParts);
                };
            } else if (colSelector.toLowerCase().startsWith('f')) {
                // 单列 f模式
                var colIdx = parseInt(colSelector.substring(1)) - 1;
                colIndexes = [colIdx]; // 保存单列索引
                keyFn = function(row) { return row[colIdx]; };
            } else {
                keyFn = function(row) { return JSON.stringify(row); };
            }
        } else if (Array.isArray(colSelector)) {
            keyFn = function(row) {
                var keyParts = [];
                for (var i = 0; i < colSelector.length; i++) {
                    keyParts.push(row[colSelector[i]]);
                }
                return JSON.stringify(keyParts);
            };
        } else {
            keyFn = function(row) { return JSON.stringify(row); };
        }

        // resultSelector 处理
        var outputFn;
        if (resultSelector === undefined) {
            if (isFunctionMode && typeof colSelector === 'function') {
                outputFn = function(row) { return [colSelector(row)]; };
            } else {
                outputFn = function(row) {
                    var out = [];
                    for (var i = 0; i < colIndexes.length; i++) {
                        out.push(row[colIndexes[i]]);
                    }
                    return out;
                };
            }
        } else if (resultSelector === '') {
            outputFn = function(row) { return row.slice(); };
        } else if (typeof resultSelector === 'string') {
            if (resultSelector.includes(',')) {
                var outIndexes = [];
                var outParts = resultSelector.split(',');
                for (var j = 0; j < outParts.length; j++) {
                    var outPart = outParts[j].trim();
                    if (outPart.toLowerCase().startsWith('f')) {
                        outIndexes.push(parseInt(outPart.substring(1)) - 1);
                    } else {
                        outIndexes.push(parseInt(outPart) - 1);
                    }
                }
                outputFn = function(row) {
                    var out = [];
                    for (var k = 0; k < outIndexes.length; k++) {
                        out.push(row[outIndexes[k]]);
                    }
                    return out;
                };
            } else if (resultSelector.toLowerCase().startsWith('f')) {
                var outIdx = parseInt(resultSelector.substring(1)) - 1;
                outputFn = function(row) { return [row[outIdx]]; };
            } else {
                outputFn = function(row) { return row.slice(); };
            }
        } else if (Array.isArray(resultSelector)) {
            outputFn = function(row) {
                var out = [];
                for (var m = 0; m < resultSelector.length; m++) {
                    out.push(row[resultSelector[m]]);
                }
                return out;
            };
        } else {
            outputFn = function(row) { return row.slice(); };
        }

        for (var i = 0; i < this._items.length; i++) {
            var row = this._items[i];
            var key = keyFn(row);
            var keyStr = typeof key === 'string' ? key : JSON.stringify(key);
            if (!seen[keyStr]) {
                seen[keyStr] = true;
                result.push(outputFn(row));
            }
        }
        return this._new(result);
    }
}

Array2D.distinct = function(arr, keySelector, resultSelector) {
    return new Array2D(arr).z去重(keySelector, resultSelector).val();
};

// 测试数据
var testData = [
    ['A001', '产品A', 100, 50],
    ['A001', '产品A', 200, 30],
    ['A002', '产品B', 150, 40],
    ['A002', '产品B', 150, 20],
    ['A003', '产品C', 180, 60]
];

console.log('=== 测试数据 ===');
console.log(JSON.stringify(testData, null, 2));

// 测试1: 按f1去重，默认只输出f1列
console.log('\n=== 测试1: arr.z去重("f1") ===');
console.log('预期: 只输出第1列 (产品编号)');
var result1 = new Array2D(testData).z去重('f1').val();
console.log(JSON.stringify(result1, null, 2));

// 测试2: 按f1,f2去重，输出所有列
console.log('\n=== 测试2: arr.z去重("f1,f2", "") ===');
console.log('预期: 按1、2列去重，输出所有列');
var result2 = new Array2D(testData).z去重('f1,f2', '').val();
console.log(JSON.stringify(result2, null, 2));

// 测试3: 按f1去重，选择输出f1,f3列
console.log('\n=== 测试3: arr.z去重("f1", "f1,f3") ===');
console.log('预期: 按第1列去重，输出第1、3列');
var result3 = new Array2D(testData).z去重('f1', 'f1,f3').val();
console.log(JSON.stringify(result3, null, 2));

// 测试4: 回调函数模式
console.log('\n=== 测试4: arr.z去重(x => x[0]) ===');
console.log('预期: 按第1列去重，输出单列');
var result4 = new Array2D(testData).z去重(x => x[0]).val();
console.log(JSON.stringify(result4, null, 2));

// 测试5: 静态方法调用
console.log('\n=== 测试5: Array2D.distinct(arr, "f1") ===');
var result5 = Array2D.distinct(testData, 'f1');
console.log(JSON.stringify(result5, null, 2));

// 测试6: 静态方法带resultSelector
console.log('\n=== 测试6: Array2D.distinct(arr, "f1", "f1,f3") ===');
var result6 = Array2D.distinct(testData, 'f1', 'f1,f3');
console.log(JSON.stringify(result6, null, 2));

console.log('\n=== 所有测试完成 ===');