/**
 * 测试 deleteRows 函数模式
 */

// 模拟 Array2D 类
function Array2D(data) {
    if (!(this instanceof Array2D)) return new Array2D(data);
    this._items = data ? (data._items ? data._items : data) : [];
}
Array2D.prototype._new = function(arr) { return new Array2D(arr); };
Array2D.prototype.val = function() { return this._items; };
Array2D.prototype.toRange = function(addr) {
    console.log('输出到', addr, ':', JSON.stringify(this._items));
    return this;
};

Array2D.prototype.z批量删除行 = function(rows) {
    var rowIndexes = [];
    console.log('[DEBUG] z批量删除行被调用, rows类型:', typeof rows);

    if (typeof rows === 'function') {
        var data = this._items;
        console.log('[DEBUG] 函数模式, 数据行数:', data.length);
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            if (Array.isArray(row)) {
                var proxy = row.slice();
                for (var c = 0; c < proxy.length; c++) {
                    proxy['f' + (c + 1)] = proxy[c];
                }
                console.log('[DEBUG] 行', i, 'proxy.f3:', proxy.f3);
                if (rows(proxy, i)) {
                    console.log('[DEBUG] 行', i, '匹配条件');
                    rowIndexes.push(i);
                }
            }
        }
        console.log('[DEBUG] 匹配的行索引:', rowIndexes);
    } else if (typeof rows === 'string') {
        // ...
    } else if (Array.isArray(rows)) {
        rowIndexes = rows;
    }

    rowIndexes.sort(function(a, b) { return b - a; });
    var result = this._items.slice();
    for (var r = 0; r < rowIndexes.length; r++) {
        if (rowIndexes[r] >= 0 && rowIndexes[r] < result.length) {
            result.splice(rowIndexes[r], 1);
        }
    }
    return this._new(result);
};
Array2D.prototype.deleteRows = Array2D.prototype.z批量删除行;

Array2D.deleteRows = function(arr, rows) {
    return new Array2D(arr).z批量删除行(rows).val();
};

// 测试数据 - 模拟Excel读取的数据（包含表头）
var arr = [
    ["ID", "产品", "国家", "数量", "价格", "年", "月", "备注"],
    [1, "Product1", "中国", 19, 1, 2023, 10, "a1"],
    [2, "Product2", "德国", 19, 5, 2023, 4, "a2"],
    [3, "Product2", "英国", 19, 5, 2022, 6, "a3"],
    [4, "Product2", "美国", 15, 5, 2024, 5, "a4"],
    [5, "Product1", "中国", 11, 1, 2024, 11, "a5"],
    [6, "Product2", "德国", 18, 5, 2023, 2, "a6"],
    [7, "Product2", "英国", 11, 5, 2023, 6, "a7"],
    [8, "Product2", "美国", 11, 5, 2023, 6, "a8"]
];

console.log('=== 测试 deleteRows ===');
console.log('原始数据行数:', arr.length);

// 测试1: 使用实例方法
console.log('\n--- 测试1: 实例方法 arr.z批量删除行(r=>r.f3=="美国") ---');
var result1 = new Array2D(arr).z批量删除行(function(r) {
    console.log('回调检查: r.f3 =', r.f3);
    return r.f3 === "美国";
});
console.log('删除美国后的数据:', JSON.stringify(result1.val(), null, 2));

// 测试2: 使用静态方法
console.log('\n--- 测试2: 静态方法 Array2D.deleteRows(arr, r=>r.f3=="美国") ---');
var result2 = Array2D.deleteRows(arr, r => r.f3 === "美国");
console.log('删除美国后的数据:', JSON.stringify(result2, null, 2));

// 测试3: 完整调用链
console.log('\n--- 测试3: 完整调用链 ---');
var rs = Array2D.deleteRows(arr, r => r.f3 === "美国");
console.log('rs =', JSON.stringify(rs, null, 2));

console.log('\n=== 测试完成 ===');