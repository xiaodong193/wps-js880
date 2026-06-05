/**
 * 测试 groupInto 修复（模拟 WPS JSA 环境）
 * 不能直接运行，需在 WPS 中运行或阅读代码确认逻辑
 */

// 需要 平方和 方法
var NODE_TEST = false;

// ========== 测试数据 ==========
var testData = [
    ['产品A', '北京', 10, 100],
    ['产品A', '上海', 20, 200],
    ['产品B', '北京', 30, 300],
    ['产品B', '上海', 40, 400],
    ['产品A', '北京', 15, 150]
];

// ========== 测试1：字符串多聚合函数 ==========
console.log('=== 测试1: 字符串多聚合函数 ===');
var rs1 = Array2D.groupInto(testData, 'f1,f2', 'count(),sum("f3"),average("f4"),textjoin("f3","+")');
console.log('结果类型:', rs1 instanceof Array2D ? 'Array2D ✓' : typeof rs1);
console.log('结果内容:', JSON.stringify(rs1));
// 预期结果（产品A,北京）：count=2, sum=10+15=25, avg=(100+150)/2=125, textjoin="10+15"
// 预期结果（产品A,上海）：count=1, sum=20, avg=200, textjoin="20"
// ...

// ========== 测试2：函数回调形式 ==========
console.log('\n=== 测试2: 函数回调形式 ===');
var rs2 = Array2D.groupInto(testData, 
    function(row) { return [row[0], row[1]]; },  // key: 产品+国家
    function(g) { 
        return [g.count(), g.sum(3), g.平方和(3), g.textjoin(4, '+')]; 
    }
);
console.log('结果类型:', rs2 instanceof Array2D ? 'Array2D ✓' : typeof rs2);
console.log('结果内容:', JSON.stringify(rs2));

// ========== 测试3：链式调用 toRange ==========
console.log('\n=== 测试3: 链式调用 toRange 支持 ===');
var rs3 = Array2D.groupInto(testData, 'f1', 'count(),sum("f3"),average("f4")');
console.log('有 toRange 方法:', typeof rs3.toRange === 'function' ? '✓' : '✗');
// 在 WPS 中调用：rs3.toRange("K3");

// ========== 测试4：平方和 ==========
console.log('\n=== 测试4: 平方和 ===');
var rs4 = Array2D.groupInto(testData, 'f1,f2', '平方和(3),平方和(4)');
console.log('结果:', JSON.stringify(rs4));
// 产品A,北京 rows: [10,100], [15,150]
// 平方和(3) = 10²+15² = 100+225 = 325
// 平方和(4) = 100²+150² = 10000+22500 = 32500

// ========== 测试5：单函数兼容 ==========
console.log('\n=== 测试5: 单函数兼容 ===');
var rs5 = Array2D.groupInto(testData, 'f1', 'sum("f3")');
console.log('结果:', JSON.stringify(rs5));
// 产品A: sum = 10+20+15 = 45
// 产品B: sum = 30+40 = 70

// ========== 测试6：空数组保护 ==========
console.log('\n=== 测试6: 空数组保护 ===');
var rs6 = Array2D.groupInto([], 'f1', 'count()');
console.log('空数组结果类型:', rs6 instanceof Array2D ? 'Array2D ✓' : typeof rs6);
console.log('空数组结果长度:', rs6.length);

// ========== 测试7：用户原始测试用例 ==========
console.log('\n=== 测试7: 用户原始测试用例 ===');
var arr = testData;
// 模拟用户代码：Array2D.groupInto(arr,'f2,f3','count(),sum("f4"),average("f5"),textjoin("f4","+")').toRange("k3");
var rs7 = Array2D.groupInto(arr, 'f2,f3', 'count(),sum("f3"),average("f4"),textjoin("f3","+")');
console.log('结果类型:', rs7 instanceof Array2D ? 'Array2D ✓' : typeof rs7);
console.log('结果:', JSON.stringify(rs7));

// 模拟用户代码：x=>[x.f2,x.f3,x.f6], g=>[g.count(), g.sum('f4'), g.平方和(4), g.textjoin("f5","+")]
var rs8 = Array2D.groupInto(arr, 
    function(row) { return [row[0], row[1]]; },
    function(g) { return [g.count(), g.sum(3), g.平方和(4), g.textjoin(4, '+')]; }
);
console.log('函数回调结果类型:', rs8 instanceof Array2D ? 'Array2D ✓' : typeof rs8);
console.log('函数回调结果:', JSON.stringify(rs8));

console.log('\n========== 测试完成 ==========');
