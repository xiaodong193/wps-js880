// 测试当前 jsaLambda 的能力
// 用 vm 把 JSA880 加载到全局作用域
const fs = require('fs');
const vm = require('vm');

// 模拟 WPS Range
class MockRange {
    constructor(addr, value2) {
        this._addr = addr;
        this._value2 = value2;
    }
    get Address() { return this._addr; }
    get Value2() { return this._value2; }
    get Value() { return this._value2; }
    get Count() { return Array.isArray(this._value2) ? this._value2.length : 1; }
}

const mockData = {
    'A2:L23': [
        ['水表', 'DN50', '只', 5, 10, 50, 'A', 100, 200, 300, 400, '猪八戒'],
        ['开孔器', 'Φ35', '个', 3, 15, 45, 'B', 50, 100, 150, 200, '沙和尚'],
        ['绳卡', 'Φ16', '盒', 2, 8, 16, 'C', 30, 60, 90, 120, '孙悟空'],
    ],
    'A1:H40': [
        ['产品', '型号', '国家', '价格', '数量', '年', '月', '日'],
        ['Product1', '大号', '中国', 100, 5, 2024, 1, 15],
        ['Product1', '中号', 'usa', 80, 3, 2024, 2, 20],
        ['Product2', '小号', '英国', 50, 2, 2024, 3, 10],
        ['Product1', '大号', '中国', 60, 4, 2024, 4, 5],
    ],
    'A2:B7': [
        ['1', 'a'], ['2', 'b'], ['3', 'c'], ['4', 'd'], ['5', 'e'], ['6', 'f']
    ],
    'D2:D4': [['x'], ['x'], ['y']],
    'A1:C17': [
        ['A', 'B', 'C'],
        ['A1', 'B1', 10],
        ['A1', 'B2', 20],
        ['A2', 'B1', 30],
        ['A2', 'B2', 40],
    ],
    'A2:C17': [
        ['A1', 'B1', 10],
        ['A1', 'B2', 20],
        ['A2', 'B1', 30],
        ['A2', 'B2', 40],
    ],
};

// 设置 globalThis
globalThis.Application = {
    ActiveSheet: { Range: function() { return new MockRange('A1', 'val'); } }
};
globalThis.Range = function(addr) {
    return new MockRange('$' + addr.replace(/!/g, '!$'), mockData[addr] || [['mock']]);
};
globalThis.Console = { log: console.log };
globalThis.console = console;

// 把 console.warn 静音避免 jsaLambda 失败时打一堆
const origWarn = console.warn;
console.warn = function() {};

// 加载 JSA880 - 用 vm.runInThisContext 让 JSA/$$/Array2D 进入全局
const JSA880Code = fs.readFileSync('/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880/JSA880.js', 'utf-8');
vm.runInThisContext(JSA880Code);

console.warn = origWarn;

console.log('JSA typeof:', typeof JSA);
console.log('JSA.jsaLambda typeof:', typeof JSA.jsaLambda);
console.log('JSA.k typeof:', typeof JSA.k);
console.log('$$ typeof:', typeof $$);
console.log('$$.superPivot typeof:', typeof ($$ && $$.superPivot));
console.log('Array2D.superPivot typeof:', typeof Array2D.superPivot);
console.log('Array2D.distinct typeof:', typeof Array2D.distinct);
console.log('Array2D.insertCols typeof:', typeof Array2D.insertCols);
console.log('Array2D.leftjoin typeof:', typeof Array2D.leftjoin);

console.log('\n--- 测试 1: JSA.getIndexs ---');
console.log(JSA.k('JSA.getIndexs', 1, 5, 1));

console.log('\n--- 测试 2: 字符串直接传 $$.superPivot ---');
try {
    var r = JSA.k('$$.superPivot', Range('A1:H40').Value2, 'f3,f2', 'f6', 'sum("f4*f5"),textjoin("f4+\'-\'+f5","+")', 1);
    console.log('ok:', JSON.stringify(r).slice(0, 200));
} catch (e) {
    console.log('FAIL:', e.message);
}

console.log('\n--- 测试 3: 直接 superPivot ---');
try {
    var r2 = Array2D.superPivot(Range('A1:H40').Value2, ['f3,f2'], ['f6'], ['sum("f4*f5"),textjoin("f4+\'-\'+f5","+")'], 1);
    console.log('ok:', JSON.stringify(r2).slice(0, 200));
} catch (e) {
    console.log('FAIL:', e.message);
}

console.log('\n--- 测试 4: $$.superPivot 在 node 端 ---');
try {
    var r3 = $$.superPivot(Range('A1:H40').Value2, ['f3,f2'], ['f6'], ['sum("f4*f5"),textjoin("f4+\'-\'+f5","+")'], 1);
    console.log('ok:', JSON.stringify(r3).slice(0, 200));
} catch (e) {
    console.log('FAIL:', e.message);
}
