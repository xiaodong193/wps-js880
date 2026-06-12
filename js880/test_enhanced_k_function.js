// ═══════════════════════════════════════════════════════════════════════
// 增强 k 函数测试套件
// 复制到 WPS JSA 工作簿中运行测试
// ═══════════════════════════════════════════════════════════════════════

/**
 * 主测试函数 - 在 WPS 中运行测试
 * 使用: 在 WPS 中选择此函数并运行 (Ctrl+Shift+F10)
 */
function test_k_function_enhancements() {
    var results = [];
    var allPassed = true;

    // 测试 1: 检查 $$ 全局对象
    results.push(test1_global_alias());

    // 测试 2: 简单数组 filter
    results.push(test2_simple_filter());

    // 测试 3: 链式 filter + map
    results.push(test3_chained_operations());

    // 测试 4: 表达式缓存
    results.push(test4_expression_caching());

    // 测试 5: 错误处理
    results.push(test5_error_handling());

    // 输出结果
    console.clear();
    console.log('╔════════════════════════════════════════════════════════════════╗');
    console.log('║ 增强 k 函数测试结果                                            ║');
    console.log('╚════════════════════════════════════════════════════════════════╝');

    for (var i = 0; i < results.length; i++) {
        var r = results[i];
        var icon = r.passed ? '✅' : '❌';
        console.log(icon + ' ' + r.name);
        if (r.message) {
            console.log('   ' + r.message);
        }
        if (!r.passed) allPassed = false;
    }

    console.log('');
    console.log(allPassed ? '✅ 所有测试通过！' : '❌ 部分测试失败，请检查实现');

    MsgBox(allPassed ? '✅ 增强 k 函数测试通过！' : '❌ 测试失败，详见日志');
}

/**
 * 测试 1: $$ 全局别名
 */
function test1_global_alias() {
    try {
        var passed = typeof $$ !== 'undefined' && $$ === Array2D;
        return {
            name: '全局 $$ 别名定义',
            passed: passed,
            message: passed
                ? '$$ 已正确映射到 Array2D'
                : '$$ 未定义或未正确映射'
        };
    } catch (e) {
        return {
            name: '全局 $$ 别名定义',
            passed: false,
            message: '异常: ' + e.message
        };
    }
}

/**
 * 测试 2: 简单数组 filter
 */
function test2_simple_filter() {
    try {
        var testData = [[1, 'a'], [2, 'b'], [3, 'c']];

        // 测试 filter 功能
        var result = JSA.jsaLambda(
            'data => data.filter((x, i) => i > 0)',
            testData
        );

        var passed = result && result.length === 2 &&
                     result[0][0] === 2 && result[1][0] === 3;

        return {
            name: '简单数组 filter 操作',
            passed: passed,
            message: passed
                ? '过滤结果正确: [[2,"b"],[3,"c"]]'
                : '结果不符预期: ' + JSON.stringify(result)
        };
    } catch (e) {
        return {
            name: '简单数组 filter 操作',
            passed: false,
            message: '异常: ' + e.message
        };
    }
}

/**
 * 测试 3: 链式 filter + map
 */
function test3_chained_operations() {
    try {
        var testData = [[1, 'a'], [2, 'b'], [3, 'c']];

        // 测试链式操作
        var result = JSA.jsaLambda(
            'data => data.filter((x, i) => i > 0).map(x => [x[0]*2, x[1].toUpperCase()])',
            testData
        );

        var passed = result && result.length === 2 &&
                     result[0][0] === 4 && result[0][1] === 'B' &&
                     result[1][0] === 6 && result[1][1] === 'C';

        return {
            name: '链式 filter + map 操作',
            passed: passed,
            message: passed
                ? '链式结果正确: [[4,"B"],[6,"C"]]'
                : '结果不符预期: ' + JSON.stringify(result)
        };
    } catch (e) {
        return {
            name: '链式 filter + map 操作',
            passed: false,
            message: '异常: ' + e.message
        };
    }
}

/**
 * 测试 4: 表达式缓存
 */
function test4_expression_caching() {
    try {
        var expr = 'data => data.filter((x, i) => i > 0)';
        var testData = [[1, 'a'], [2, 'b']];

        // 第一次调用
        var result1 = JSA.jsaLambda(expr, testData);

        // 第二次调用（应该使用缓存）
        var result2 = JSA.jsaLambda(expr, testData);

        // 检查结果一致性
        var passed = result1 && result2 &&
                     result1.length === result2.length &&
                     result1[0][0] === result2[0][0];

        return {
            name: '表达式缓存机制',
            passed: passed,
            message: passed
                ? '多次调用结果一致'
                : '缓存可能无效'
        };
    } catch (e) {
        return {
            name: '表达式缓存机制',
            passed: false,
            message: '异常: ' + e.message
        };
    }
}

/**
 * 测试 5: 错误处理
 */
function test5_error_handling() {
    try {
        // 测试 k 函数错误处理
        var errorResult = k('invalid => invalid..invalid', [1, 2, 3]);

        var passed = typeof errorResult === 'string' &&
                     errorResult.indexOf('#K_ERR') === 0;

        return {
            name: '错误处理机制',
            passed: passed,
            message: passed
                ? '错误被正确捕获并返回错误信息'
                : '错误处理可能不正确'
        };
    } catch (e) {
        return {
            name: '错误处理机制',
            passed: false,
            message: '异常: ' + e.message
        };
    }
}

// ═══════════════════════════════════════════════════════════════════════
// 单元格公式示例测试
// ═══════════════════════════════════════════════════════════════════════

/**
 * 在 Excel/WPS 中复制以下公式到单元格，验证功能
 *
 * 注意: 需要先在 A1:D5 建立测试数据
 */

/*
├─ 测试表 - 推荐数据结构
├─ A列: ID (1, 2, 3, 4, 5)
├─ B列: 产品 (Product1, Product1, Product2, Product2, Product1)
├─ C列: 销售额 (1000, 1500, 800, 2000, 1200)
├─ D列: 数量 (10, 15, 8, 20, 12)

F1: 标题 = "测试项目"
F2: 结果 = "输出"
G1: 说明 = "公式说明"
G2: 说明文本

F3: 标题 = "1. 简单 filter"
F4: 公式 = =k("data=>data.filter((x,i)=>i>0)",A1:D5)
G3: 说明 = "跳过第一行"
G4: 说明 = ""

F6: 标题 = "2. Filter + Map"
F7: 公式 = =k("data=>data.filter((x,i)=>i>0).map(x=>[x[0]*2,x[1],x[2],x[3]])",A1:D5)
G6: 说明 = "ID 列翻倍"
G7: 说明 = ""

F9: 标题 = "3. 条件筛选"
F10: 公式 = =k("data=>data.filter((x,i)=>i==0||x[1]=='Product1')",A1:D5)
G9: 说明 = "只显示 Product1"
G10: 说明 = ""

F12: 标题 = "4. 复杂链式"
F13: 公式 = =k("data=>data.filter((x,i)=>i==0||x[1]=='Product1').map(x=>[x[0],x[1],x[2]*1.1])",A1:D5)
G12: 说明 = "Product1 销售额增加 10%"
G13: 说明 = ""

F15: 标题 = "5. 数值聚合"
F16: 公式 = =k("data=>data.reduce((sum,x,i)=>i>0?sum+x[2]:sum,0)",A1:D5)
G15: 说明 = "计算总销售额"
G16: 说明 = ""

*/

// ═══════════════════════════════════════════════════════════════════════
// 实际生产环境中的复杂示例
// ═══════════════════════════════════════════════════════════════════════

/**
 * 示例 A: 透视表筛选 + 条件过滤
 *
 * 场景: 你有源数据在 A1:H40，包含产品、地区、销售额等信息
 *       需要生成透视表并只显示 Product1 的数据
 *
 * 公式:
 * =k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",
 *    A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")
 *
 * 说明:
 * - superPivot(): 生成按 f3 行、f2 列的透视表，统计 f4 的计数、求和、连接
 * - .filter(): 保留第一行（标题）和 f2 列为 Product1 的行
 * - 返回结果自动在下面单元格溢出
 */

/**
 * 示例 B: 分组求和 + 排序
 *
 * 场景: 销售数据需要按产品分组求和，然后排序
 *
 * 公式:
 * =k("data=>data.filter((x,i)=>i>0).reduce((acc,x)=>{var key=x[1];var existing=acc.find(e=>e[0]==key);if(existing){existing[2]+=x[2]}else{acc.push([x[0],key,x[2]])}return acc},[]).sort((a,b)=>b[2]-a[2])",
 *    A1:D100)
 *
 * 说明:
 * - filter: 跳过标题行
 * - reduce: 按产品名分组累加销售额
 * - sort: 按销售额降序排列
 */

/**
 * 示例 C: 数据验证 + 转换
 *
 * 场景: 导入的数据需要验证和清理
 *
 * 公式:
 * =k("data=>data.filter((x,i)=>i==0||x[2]>0).map(x=>({id:x[0],name:x[1],amount:x[2],valid:x[2]>0?'Y':'N'}))",
 *    A1:D1000)
 *
 * 说明:
 * - filter: 只保留销售额 > 0 的记录
 * - map: 转换为对象格式并添加有效性标志
 */

console.log('✅ 增强 k 函数测试套件已加载');
console.log('运行 test_k_function_enhancements() 执行测试');
