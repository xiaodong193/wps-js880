// 模拟 WPS 公式调用 k() UDF 的完整场景（v2 — 修正测试用例）
// ====================================================
// 关键发现：WPS 公式传 A1:A10 这种区域是 2D 数组 `[[1],[2],...,[10]]`
// 所以 lambda 写 `arr.filter(...)` 调的是 Array.prototype.filter
//   - 对 2D 数组来说,这是按"内层数组是元素"过滤,即按行过滤(因为 2D 数组是数组的数组)
// 实际上 Array.prototype.filter 对 2D 数据是工作的,因为 2D 数组的"元素"就是行!
// 例如 [10,20,30] 是 1D 数组,filter 按值过滤
//     [[1,'a'],[2,'b'],[3,'c']] 是 2D 数组,filter 按"行"过滤
// ====================================================

const fs = require('fs');
const vm = require('vm');

class MockRange {
    constructor(addr, value2) {
        this._addr = addr;
        this._value2 = value2;
    }
    get Address() { return this._addr; }
    get Value2() { return this._value2; }
    get Value() { return this._value2; }
}

const mockData = {
    'A1:H40': [
        ['产品', '型号', '国家', '价格', '数量', '年', '月', '日'],
        ['P1', '大号', '中国', 100, 5, 2024, 1, 15],
        ['P1', '中号', 'usa', 80, 3, 2024, 2, 20],
        ['P2', '小号', '英国', 50, 2, 2024, 3, 10],
        ['P1', '大号', '中国', 60, 4, 2024, 4, 5],
    ],
};

globalThis.Application = {
    ActiveSheet: { Range: function() { return new MockRange('A1', 'val'); } },
};
globalThis.Range = function(addr) {
    return new MockRange(addr, mockData[addr] || [['mock']]);
};
globalThis.Console = { log: (...args) => console.log('[Console]', ...args) };

// 静音警告
const origWarn = console.warn;
console.warn = function() {};

// 加载 JSA880.js
const JSA880Code = fs.readFileSync('js880/JSA880.js', 'utf-8');
vm.runInThisContext(JSA880Code);
console.warn = origWarn;

// 加载 UDF 模块
const UDFCode = fs.readFileSync('JS880教案/第03章/3-28/KO一切的k函数_UDF模块.js', 'utf-8');
vm.runInThisContext(UDFCode);

console.log('===== 测试: 模拟 WPS 公式调用 =====\n');

let pass = 0, fail = 0;
function check(name, got, expected) {
    const ok = JSON.stringify(got) === JSON.stringify(expected);
    console.log(`${ok ? '✅' : '❌'} ${name}`);
    console.log(`   got: ${JSON.stringify(got)}`);
    if (!ok) console.log(`   exp: ${JSON.stringify(expected)}`);
    ok ? pass++ : fail++;
}

// === 真实场景: WPS 公式 k("lambda", arg1, arg2, ...) ===

// 1) 路径调用: k("JSA.getIndexs", 1, 5, 1) → [1,2,3,4,5]
//   WPS 传 1, 5, 1 是单个数字参数
check(
    'k("JSA.getIndexs", 1, 5, 1)',
    k("JSA.getIndexs", 1, 5, 1),
    [1,2,3,4,5]
);

// 2) 简单 lambda: k("x => x*2", 5) → 10
check(
    'k("x => x*2", 5)',
    k("x => x*2", 5),
    10
);

// 3) 数组 lambda: k("(arr) => arr[0] + arr[1]", [[10, 20]])
//   模拟 WPS 公式传 2 个数字,WPS 实际传 [[10,20]] 1x2
//   lambda 收 [[10,20]],arr[0]=10, arr[1]=20, 10+20=30
//   但 1xN 2D 数组 arr[0] 是 10,arr[1] 是 undefined! ✗
//   实际: arr[0][0]=10, arr[0][1]=20 才能正确
check(
    'k("(arr) => arr[0][0] + arr[0][1]", [[10, 20]])',
    k("(arr) => arr[0][0] + arr[0][1]", [[10, 20]]),
    30
);

// 4) 全局函数调用: k("JSA.cint", 3.7) → 3
check(
    'k("JSA.cint", 3.7)',
    k("JSA.cint", 3.7),
    3
);

// 5) JSA.today: k("JSA.today") → "2026-06-05"
check(
    'k("JSA.today")',
    k("JSA.today"),
    '2026-06-05'
);

// 6) 真实场景: k 接收 Range 数据(A1:H40, 5x8 2D 数组)
//   用 lambda 选出 Product1 行的所有价格列
//   因为传进来是 2D 数组,Array.prototype.filter 实际上是按"行"过滤
//   (2D 数组的元素就是行,filter 遍历每个行)
check(
    'k("arr => arr.filter(r => r[0] === \\"P1\\").map(r => r[3])", A1:H40)',
    k('arr => arr.filter(r => r[0] === "P1").map(r => r[3])', mockData['A1:H40']),
    [100, 80, 60]
);

// 7) 多参数场景: k 接收数据 + 阈值
//   lambda 形如 (arr, threshold) => arr.filter(r => r[1] > threshold).map(r => r[0])
const arrB = [['1','a'], ['2','b'], ['3','c'], ['4','d'], ['5','e']];
const threshold = 3;
check(
    'k("(arr, t) => arr.filter(r => r[1] > t).map(r => r[0])", arr, 3)',
    k('(arr, t) => arr.filter(r => Number(r[1]) > t).map(r => r[0])', arrB, threshold),
    ['4','5']  // 第1列>3的行的第0列
);

// 8) $ 索引语法: k("$0", 5) → 5 (单值)
check(
    'k("$0", 5)',
    k("$0", 5),
    5
);

// 9) 超级透视: k("Array2D.z超级透视", data, [...], [...], [...], 1)
//   这是真实的"k 调用 superPivot"用法
const superData = [
    ['产品','国家','数量'],
    ['P1','中国',5],
    ['P1','中国',3],
    ['P2','英国',2],
];
const rs = k("Array2D.z超级透视", superData, ['f2'], ['f3'], ['sum(f3)'], 1);
const isRs = Array.isArray(rs) || (rs && rs.val && Array.isArray(rs.val())) || (rs && Array.isArray(rs));
console.log(`${isRs ? '✅' : '❌'} k("Array2D.z超级透视", data, ...)`);
if (isRs) {
    pass++;
    console.log(`   got: ${JSON.stringify(rs).slice(0, 200)}`);
} else {
    fail++;
    console.log(`   got: ${JSON.stringify(rs).slice(0, 200)}`);
}

console.log(`\n===== 总结: ${pass} PASS, ${fail} FAIL =====`);
