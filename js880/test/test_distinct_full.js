/**
 * 测试 Array2D.distinct 所有用法
 * 基于课程材料 3.4 的完整测试
 */

function Array2D(arr) {
    if (!(this instanceof Array2D)) return new Array2D(arr);
    this._items = Array.isArray(arr) ? arr : [];
}
Array2D.prototype._new = function(arr) { return new Array2D(arr); };
Array2D.prototype.val = function() { return this._items; };

// 模拟 asString 和 logjson
function asString(val) {
    return val === null || val === undefined ? '' : String(val);
}
function logjson(obj) {
    console.log(JSON.stringify(obj, null, 2));
}

// ==================== z去重 实现 ====================
Array2D.prototype.z去重 = function(colSelector, resultSelector) {
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
            var colIdx = parseInt(colSelector.substring(1)) - 1;
            colIndexes = [colIdx];
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

    var outputFn;
    if (resultSelector === undefined) {
        if (isFunctionMode && typeof colSelector === 'function') {
            outputFn = function(row, key) {
                if (Array.isArray(key)) return key;
                return [key];
            };
        } else if (!colIndexes || colIndexes.length === 0) {
            outputFn = function(row) {
                if (!row) return [];
                return Array.isArray(row) ? row.slice() : [row];
            };
        } else {
            outputFn = function(row) {
                if (!row) return [];
                var out = [];
                for (var i = 0; i < colIndexes.length; i++) {
                    out.push(row[colIndexes[i]]);
                }
                return out;
            };
        }
    } else if (resultSelector === '') {
        outputFn = function(row) {
            if (!row) return [];
            return Array.isArray(row) ? row.slice() : [row];
        };
    } else if (typeof resultSelector === 'string') {
        if (resultSelector.includes(',')) {
            var outIndexes = [];
            var outParts = resultSelector.split(',');
            for (var j = 0; j < outParts.length; j++) {
                var outPart = outParts[j].trim();
                if (outPart.toLowerCase().startsWith('f')) {
                    if (outPart.includes('-')) {
                        // 处理范围 f3-f5
                        var range = outPart.split('-');
                        var start = parseInt(range[0].toLowerCase().replace('f', ''));
                        var end = parseInt(range[1].toLowerCase().replace('f', ''));
                        for (var r = start; r <= end; r++) {
                            outIndexes.push(r - 1);
                        }
                    } else {
                        outIndexes.push(parseInt(outPart.substring(1)) - 1);
                    }
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
            if (resultSelector.includes('-')) {
                // 处理范围 f3-f5
                var range = resultSelector.split('-');
                var start = parseInt(range[0].toLowerCase().replace('f', ''));
                var end = parseInt(range[1].toLowerCase().replace('f', ''));
                var rangeIndexes = [];
                for (var r = start; r <= end; r++) {
                    rangeIndexes.push(r - 1);
                }
                outputFn = function(row) {
                    var out = [];
                    for (var i = 0; i < rangeIndexes.length; i++) {
                        out.push(row[rangeIndexes[i]]);
                    }
                    return out;
                };
            } else {
                var outIdx = parseInt(resultSelector.substring(1)) - 1;
                outputFn = function(row) { return [row[outIdx]]; };
            }
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
            result.push(outputFn(row, key));
        }
    }
    return this._new(result);
};
Array2D.prototype.distinct = Array2D.prototype.z去重;
Array2D.distinct = function(arr, colSelector, resultSelector) {
    return new Array2D(arr).z去重(colSelector, resultSelector).val();
};

// ==================== z映射 实现 ====================
Array2D.prototype.z映射 = function(mapper) {
    var result = this._items.map(function(row, index) {
        var proxy = Array.isArray(row) ? row.slice() : [row];
        for (var c = 0; c < proxy.length; c++) {
            proxy['f' + (c + 1)] = proxy[c];
        }
        return mapper(proxy, index);
    });
    return this._new(result);
};
Array2D.prototype.map = Array2D.prototype.z映射;

// ==================== 测试数据 ====================
var testData = [
    ['A001', '产品A', 100, 50],
    ['A001', '产品A', 200, 30],
    ['A002', '产品B', 150, 40],
    ['A002', '产品B', 150, 20],
    ['A003', '产品C', 180, 60]
];

var testDataWithHeader = [
    ['序号', '产品编号', '品名', '净含量', '售卖规格'],
    ['1', 'H0188', '轮胎护理剂', '500ML', '24瓶/箱'],
    ['2', 'H0571', 'PIXY海贼王', '只', '12只/箱'],
    ['3', 'H0423', '镇邪英雄--柠檬', '只', '24只/箱'],
    ['4', 'HN073', '洗车香波', '1L', '12瓶/箱'],
    ['5', 'H0188', '双效洗车香波', '750ML', '12瓶/箱'],
    ['6', 'H0593', '研磨剂STEP1', '250G', '6瓶/箱'],
    ['7', 'H0594', '研磨剂STEP2', '250G', '6瓶/箱'],
    ['8', 'H0595', '抛光剂STEP3', '250G', '6瓶/箱'],
    ['9', 'H0188', '轮胎护理剂', '250G', '6瓶/箱'],
    ['10', 'H0597', '密封釉STEP5', '250G', '6瓶/箱'],
    ['11', 'H0188', '轮胎护理剂', '500ML', '24瓶/箱']
];

var specData = [
    ['24瓶/箱'],
    ['12只/箱'],
    ['24只/箱'],
    ['12瓶/箱'],
    ['12瓶/箱'],
    ['6瓶/箱'],
    ['6瓶/箱'],
    ['6瓶/箱'],
    ['6瓶/箱'],
    ['6瓶/箱']
];

var abTestData = [
    ['A,B'],
    ['B,A'],
    ['C,D'],
    ['D,C'],
    ['A,B,C,D'],
    ['B,A,C,D'],
    ['A,C,D'],
    ['A,B'],
    ['B,A,D,C'],
    ['A,C,D'],
    ['A,B'],
    ['A,B,C,D'],
    ['B,A,C,D'],
    ['A,C,D']
];

// ==================== 测试用例 ====================
var passed = 0;
var failed = 0;

function test(name, actual, expected) {
    var actualStr = JSON.stringify(actual);
    var expectedStr = JSON.stringify(expected);
    if (actualStr === expectedStr) {
        console.log(`✅ ${name}`);
        passed++;
    } else {
        console.log(`❌ ${name}`);
        console.log(`   Expected: ${expectedStr}`);
        console.log(`   Actual:   ${actualStr}`);
        failed++;
    }
}

console.log('='.repeat(60));
console.log('3.4 Array2D.distinct 完整测试');
console.log('='.repeat(60));

// 测试1: 整行去重（不传参数）
console.log('\n--- 测试1: 整行去重 ---');
var r1 = Array2D.distinct(testData);
test('整行去重', r1.length, 5);

// 测试2: 数字索引去重
console.log('\n--- 测试2: 数字索引去重 ---');
var r2 = Array2D.distinct(testData, 0);
test('按第0列去重', r2, [['A001', '产品A', 100, 50], ['A002', '产品B', 150, 40], ['A003', '产品C', 180, 60]]);

// 测试3: f1 单列去重
console.log('\n--- 测试3: f1 单列去重 ---');
var r3 = Array2D.distinct(testData, 'f1');
test('f1单列去重', r3, [['A001'], ['A002'], ['A003']]);

// 测试4: f1,f2 多列组合去重
console.log('\n--- 测试4: f1,f2 多列组合去重 ---');
var r4 = Array2D.distinct(testData, 'f1,f2', '');
test('f1,f2多列去重', r4, [
    ['A001', '产品A', 100, 50],
    ['A002', '产品B', 150, 40],
    ['A003', '产品C', 180, 60]
]);

// 测试5: 函数模式 x=>x.f1
// 注意：函数模式需要使用 z映射 或传入带proxy的数据
console.log('\n--- 测试5: 函数模式 x=>x.f1 (使用z映射) ---');
var data5 = new Array2D(testData).z映射(x => x.f1).val();
var r5 = Array2D.distinct(data5);
test('函数模式x=>x.f1', r5, [['A001'], ['A002'], ['A003']]);

// 测试6: 函数模式 x=>[x.f1,x.f2] 多列
console.log('\n--- 测试6: 函数模式 x=>[x.f1,x.f2] (使用z映射) ---');
var data6 = new Array2D(testData).z映射(x => [x.f1, x.f2]).val();
var r6 = Array2D.distinct(data6);
test('函数模式多列', r6, [
    ['A001', '产品A'],
    ['A002', '产品B'],
    ['A003', '产品C']
]);

// 测试7: 函数模式 + resultSelector (使用z映射处理后去重)
// 先用 z映射 处理，然后用数字索引去重
console.log('\n--- 测试7: 函数模式 + resultSelector ---');
// z映射 处理后返回 [[key, value], ...] 形式
var data7 = new Array2D(testData).z映射(x => [x.f1, x.f3]).val();
// 按第0列去重，输出所有列
var r7 = Array2D.distinct(data7, 0, '');
test('函数+resultSelector', r7, [
    ['A001', 100],
    ['A002', 150],
    ['A003', 180]
]);

// 测试8: f1 + resultSelector 输出指定列
console.log('\n--- 测试8: f1 + resultSelector ---');
var r8 = Array2D.distinct(testData, 'f1', 'f1,f3');
test('f1+resultSelector', r8, [
    ['A001', 100],
    ['A002', 150],
    ['A003', 180]
]);

// 测试9: 按产品编号去重，输出品名和售卖规格（跳过表头）
console.log('\n--- 测试9: 按产品编号去重，输出品名和售卖规格 ---');
var r9 = new Array2D(testDataWithHeader.slice(1)).z去重('f2', 'f3,f5').val();
test('多列输出', r9, [
    ['轮胎护理剂', '24瓶/箱'],
    ['PIXY海贼王', '12只/箱'],
    ['镇邪英雄--柠檬', '24只/箱'],
    ['洗车香波', '12瓶/箱'],
    ['研磨剂STEP1', '6瓶/箱'],
    ['研磨剂STEP2', '6瓶/箱'],
    ['抛光剂STEP3', '6瓶/箱'],
    ['密封釉STEP5', '6瓶/箱']
]);

// 测试10: 单列预处理去重（删除数字）
console.log('\n--- 测试10: 单列预处理去重（删除数字）---');
var specCleaned = new Array2D(specData).z映射(x => asString(x.f1).replace(/\d+/g, '')).val();
var r10 = Array2D.distinct(specCleaned);
test('删除数字后去重', r10, [['瓶/箱'], ['只/箱']]);

// 测试11: 预处理后去重（删除数字后去重）
console.log('\n--- 测试11: 预处理后去重 ---');
var data11 = new Array2D(specData).z映射(x => asString(x.f1).replace(/\d+/g, '')).val();
var r11 = Array2D.distinct(data11);
test('回调函数预处理', r11, [['瓶/箱'], ['只/箱']]);

// 测试12: 组合条件去重 (AB=BA) - 使用z映射处理
console.log('\n--- 测试12: 组合条件去重 (AB=BA) ---');
var data12 = new Array2D(abTestData).z映射(x => {
    var parts = asString(x.f1).split(',').sort();
    return parts.join(',');
}).val();
var r12 = Array2D.distinct(data12);
test('AB=BA去重', r12, [['A,B'], ['C,D'], ['A,B,C,D'], ['A,C,D']]);

// 测试13: 静态方法调用
console.log('\n--- 测试13: 静态方法调用 ---');
var r13 = Array2D.distinct(testData, 'f1');
test('静态方法', r13, [['A001'], ['A002'], ['A003']]);

// 测试14: 实例方法链式调用 - 需要指定 '' 输出所有列
console.log('\n--- 测试14: 实例方法链式调用 ---');
var r14 = new Array2D(testData).z去重('f1', '').val();
test('实例方法', r14, [['A001', '产品A', 100, 50], ['A002', '产品B', 150, 40], ['A003', '产品C', 180, 60]]);

// 测试15: resultSelector 空字符串输出所有列
console.log('\n--- 测试15: resultSelector 空字符串 ---');
var r15 = Array2D.distinct(testData, 'f1', '');
test('空字符串输出所有列', r15, [
    ['A001', '产品A', 100, 50],
    ['A002', '产品B', 150, 40],
    ['A003', '产品C', 180, 60]
]);

// 测试16: 数组形式列选择器
console.log('\n--- 测试16: 数组形式列选择器 ---');
var r16 = Array2D.distinct(testData, [0, 1], '');
test('数组形式', r16, [
    ['A001', '产品A', 100, 50],
    ['A002', '产品B', 150, 40],
    ['A003', '产品C', 180, 60]
]);

// 测试17: 范围选择 f3-f4（按第3列去重，输出第3、4列）
console.log('\n--- 测试17: 范围选择 f3-f4 ---');
// f3列的值: [100, 200, 150, 150, 180] -> 4个不同值，所以保留4条记录
var r17 = Array2D.distinct(testData, 'f3', 'f3-f4');
test('范围选择', r17, [[100, 50], [200, 30], [150, 40], [180, 60]]);

// ==================== 测试结果 ====================
console.log('\n' + '='.repeat(60));
console.log(`测试结果: ${passed} 通过, ${failed} 失败`);
console.log('='.repeat(60));

if (failed > 0) {
    process.exit(1);
}