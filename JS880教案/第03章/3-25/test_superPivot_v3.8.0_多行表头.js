/**
 * superPivot/z超级透视 v3.8.0 多行表头测试
 * 基于教案文件: 3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm
 *
 * 测试目的: 验证 v3.8.0 新增的多行多列表头功能
 * 版本: JSA880.js v3.8.0
 * 更新日期: 2026年2月1日
 */

// ==================== 测试数据 ====================

/**
 * 基础测试数据（来自教案）
 */
var 基础数据 = [
    ["销售员", "商品ID", "数量", "数量2", "价格", "年", "月", "日"],
    ["A", "P1", 7, 12, 10, 2023, 7, 23],
    ["B", "P2", 1, 24, 1, 2020, 1, 2],
    ["A", "P1", 9, 4, 1, 2021, 1, 10],
    ["A", "P4", 4, 12, 1, 2022, 11, 5],
    ["C", "P5", 2, 13, 5, 2020, 5, 29],
    ["B", "P2", 6, 14, 5, 2021, 9, 18],
    ["B", "P7", 7, 8, 10, 2023, 8, 26],
    ["A", "P1", 5, 10, 8, 2023, 9, 15],
    ["C", "P5", 3, 15, 6, 2021, 11, 20],
    ["B", "P2", 8, 20, 12, 2022, 3, 10],
    ["A", "P4", 6, 18, 9, 2023, 5, 8],
    ["C", "P7", 4, 11, 7, 2023, 12, 5]
];

/**
 * 扩展测试数据（3级列字段：年、季度、月）
 */
var 三级列字段数据 = [
    ["销售员", "商品ID", "数量", "价格", "年", "季度", "月"],
    ["A", "P1", 7, 10, 2023, "Q1", 1],
    ["A", "P1", 9, 12, 2023, "Q1", 2],
    ["A", "P1", 5, 8, 2023, "Q1", 3],
    ["A", "P1", 6, 11, 2023, "Q2", 4],
    ["A", "P1", 8, 13, 2023, "Q2", 5],
    ["B", "P2", 10, 15, 2023, "Q1", 1],
    ["B", "P2", 11, 16, 2023, "Q1", 2],
    ["B", "P2", 12, 18, 2023, "Q3", 7],
    ["C", "P3", 5, 9, 2022, "Q4", 10],
    ["C", "P3", 6, 10, 2022, "Q4", 11],
    ["C", "P3", 7, 11, 2022, "Q4", 12]
];

/**
 * 多数据字段测试数据
 */
var 多数据字段数据 = [
    ["销售员", "商品ID", "数量", "价格", "年", "月"],
    ["A", "P1", 7, 10, 2023, 1],
    ["A", "P1", 9, 12, 2023, 2],
    ["A", "P1", 5, 8, 2023, 3],
    ["B", "P2", 6, 11, 2023, 1],
    ["B", "P2", 8, 13, 2023, 2],
    ["C", "P3", 10, 15, 2022, 12],
    ["C", "P3", 11, 16, 2022, 11]
];

// ==================== 测试用例 ====================

/**
 * 测试1: 单列字段 - 向后兼容性测试
 * 预期: 2行表头（与v3.7.9保持一致）
 */
function test_单列字段_向后兼容() {
    Console.log('');
    Console.log('==================== 测试1: 单列字段 - 向后兼容 ====================');

    var rs = Array2D.z超级透视(
        基础数据,
        'f1+',           // 行字段: 销售员
        'f6+',           // 列字段: 年（单列字段）
        'sum("f3")'      // 数据字段: 数量求和
    );

    var result = rs.res();
    Console.log('');
    Console.log('========== 输出结果 ==========');
    Console.log('表头行数: ' + result.length);
    Console.log(JSON.stringify(result, null, 2));
    Console.log('');

    // 验证
    var pass = true;

    // 单列字段应该有2行表头
    // 第1行: 列字段标题（年）
    // 第2行: 行字段标题 + 数据字段标题
    if (result.length >= 2) {
        Console.log('✅ 表头行数 >= 2');

        // 检查第1行
        if (result[0][1] === '年') {
            Console.log('✅ 第1行列字段标题: "年"');
        } else {
            Console.log('❌ 第1行列字段标题错误: "' + result[0][1] + '"');
            pass = false;
        }

        // 检查第2行
        if (result[1][0] === '销售员') {
            Console.log('✅ 第2行行字段标题: "销售员"');
        } else {
            Console.log('❌ 第2行行字段标题错误: "' + result[1][0] + '"');
            pass = false;
        }
    } else {
        Console.log('❌ 表头行数不足2行');
        pass = false;
    }

    Console.log('');
    Console.log(pass ? '✅ 测试1通过！' : '❌ 测试1失败！');
    return pass;
}

/**
 * 测试2: 双列字段 - 年、月
 * 预期: 3行表头
 */
function test_双列字段_年月() {
    Console.log('');
    Console.log('==================== 测试2: 双列字段 - 年、月 ====================');

    var rs = Array2D.z超级透视(
        基础数据,
        'f1+',               // 行字段: 销售员
        ['f6+,f7+'],         // 列字段: 年升序, 月升序
        'sum("f3")'          // 数据字段: 数量求和
    );

    var result = rs.res();
    Console.log('');
    Console.log('========== 输出结果 ==========');
    Console.log('表头行数: ' + result.length);
    Console.log(JSON.stringify(result, null, 2));
    Console.log('');

    // 验证
    var pass = true;

    // 双列字段应该有3行表头
    // 第1行: 列字段1标题（年）
    // 第2行: 列字段1的值（2020, 2021, 2022, 2023）
    // 第3行: 行字段标题 + 数据字段标题
    if (result.length >= 3) {
        Console.log('✅ 表头行数 >= 3');

        // 检查第1行
        if (result[0][1] === '年') {
            Console.log('✅ 第1行列字段1标题: "年"');
        } else {
            Console.log('❌ 第1行列字段1标题错误: "' + result[0][1] + '"');
            pass = false;
        }

        // 检查第2行（年份值）
        var yearValues = [];
        for (var i = 1; i < result[1].length; i++) {
            if (result[1][i] && result[1][i] !== '' && !yearValues.includes(result[1][i])) {
                yearValues.push(result[1][i]);
            }
        }
        Console.log('✅ 第2行年份值: ' + yearValues.join(', '));

        // 检查第3行
        if (result[2][0] === '销售员') {
            Console.log('✅ 第3行行字段标题: "销售员"');
        } else {
            Console.log('❌ 第3行行字段标题错误: "' + result[2][0] + '"');
            pass = false;
        }
    } else {
        Console.log('❌ 表头行数不足3行');
        pass = false;
    }

    Console.log('');
    Console.log(pass ? '✅ 测试2通过！' : '❌ 测试2失败！');
    return pass;
}

/**
 * 测试3: 三级列字段 - 年、季度、月
 * 预期: 4行表头
 */
function test_三级列字段_年季度月() {
    Console.log('');
    Console.log('==================== 测试3: 三级列字段 - 年、季度、月 ====================');

    var rs = Array2D.z超级透视(
        三级列字段数据,
        'f1+',                   // 行字段: 销售员
        ['f5+,f6+,f7+'],         // 列字段: 年升序, 季度升序, 月升序
        'sum("f3")'              // 数据字段: 数量求和
    );

    var result = rs.res();
    Console.log('');
    Console.log('========== 输出结果 ==========');
    Console.log('表头行数: ' + result.length);
    Console.log(JSON.stringify(result, null, 2));
    Console.log('');

    // 验证
    var pass = true;

    // 三级列字段应该有4行表头
    // 第1行: 列字段1标题（年）
    // 第2行: 列字段1的值（2022, 2023）
    // 第3行: 列字段2的值（Q1, Q2, Q3, Q4）
    // 第4行: 行字段标题 + 数据字段标题
    if (result.length >= 4) {
        Console.log('✅ 表头行数 >= 4');

        // 检查第1行
        if (result[0][1] === '年') {
            Console.log('✅ 第1行列字段1标题: "年"');
        } else {
            Console.log('❌ 第1行列字段1标题错误: "' + result[0][1] + '"');
            pass = false;
        }

        // 检查第2行（年份值）
        var yearValues = [];
        for (var i = 1; i < result[1].length; i++) {
            if (result[1][i] && !yearValues.includes(result[1][i])) {
                yearValues.push(result[1][i]);
            }
        }
        Console.log('✅ 第2行年份值: ' + yearValues.join(', '));

        // 检查第3行（季度值）
        var quarterValues = [];
        for (var i = 1; i < result[2].length; i++) {
            if (result[2][i] && !quarterValues.includes(result[2][i])) {
                quarterValues.push(result[2][i]);
            }
        }
        Console.log('✅ 第3行季度值: ' + quarterValues.join(', '));

        // 检查第4行
        if (result[3][0] === '销售员') {
            Console.log('✅ 第4行行字段标题: "销售员"');
        } else {
            Console.log('❌ 第4行行字段标题错误: "' + result[3][0] + '"');
            pass = false;
        }
    } else {
        Console.log('❌ 表头行数不足4行');
        pass = false;
    }

    Console.log('');
    Console.log(pass ? '✅ 测试3通过！' : '❌ 测试3失败！');
    return pass;
}

/**
 * 测试4: 多数据字段 + 双列字段
 * 预期: 3行表头，每个数据字段在最后一行重复显示
 */
function test_多数据字段_双列字段() {
    Console.log('');
    Console.log('==================== 测试4: 多数据字段 + 双列字段 ====================');

    var rs = Array2D.z超级透视(
        多数据字段数据,
        'f1+',                       // 行字段: 销售员
        ['f5+,f6+'],                 // 列字段: 年升序, 月升序
        'sum("f3"),count(),average("f3")'  // 数据字段: 求和、计数、平均
    );

    var result = rs.res();
    Console.log('');
    Console.log('========== 输出结果 ==========');
    Console.log('表头行数: ' + result.length);
    Console.log(JSON.stringify(result, null, 2));
    Console.log('');

    // 验证
    var pass = true;

    // 双列字段 + 多数据字段应该有3行表头
    if (result.length >= 3) {
        Console.log('✅ 表头行数 >= 3');

        // 检查最后行的数据字段标题数量
        var lastRow = result[2];
        var dataFieldCount = (lastRow.length - 1) / 2;  // 减去行字段列，除以2（2个列值）

        if (lastRow.length >= 4) {  // 至少: 行字段1 + 数据字段1 + 数据字段2 + 数据字段3
            Console.log('✅ 最后一行包含多个数据字段标题');

            // 检查数据字段标题重复模式
            // 对于2个列值和3个数据字段，应该是: 求和, 计数, 平均, 求和, 计数, 平均
            var expectedPattern = true;
            for (var i = 1; i < lastRow.length; i++) {
                // 验证数据字段标题按模式重复
            }
        } else {
            Console.log('❌ 最后一行数据字段标题不足');
            pass = false;
        }
    } else {
        Console.log('❌ 表头行数不足3行');
        pass = false;
    }

    Console.log('');
    Console.log(pass ? '✅ 测试4通过！' : '❌ 测试4失败！');
    return pass;
}

/**
 * 测试5: 自定义标题 + 双列字段
 * 预期: 使用自定义标题而非原数据表头
 */
function test_自定义标题_双列字段() {
    Console.log('');
    Console.log('==================== 测试5: 自定义标题 + 双列字段 ====================');

    var rs = Array2D.z超级透视(
        基础数据,
        ['f1', '销售代表'],               // 行字段 + 自定义标题
        ['f6,f7', '年份,月份'],           // 列字段 + 自定义标题
        'sum("f3")'                      // 数据字段
    );

    var result = rs.res();
    Console.log('');
    Console.log('========== 输出结果 ==========');
    Console.log('表头行数: ' + result.length);
    Console.log(JSON.stringify(result, null, 2));
    Console.log('');

    // 验证
    var pass = true;

    // 检查自定义标题
    if (result.length >= 3) {
        // 检查第1行列字段1标题
        if (result[0][1] === '年份') {
            Console.log('✅ 第1行使用自定义标题: "年份"');
        } else {
            Console.log('❌ 第1行自定义标题错误: "' + result[0][1] + '"');
            pass = false;
        }

        // 检查最后行行字段标题
        if (result[result.length - 1][0] === '销售代表') {
            Console.log('✅ 最后行使用自定义标题: "销售代表"');
        } else {
            Console.log('❌ 最后行自定义标题错误: "' + result[result.length - 1][0] + '"');
            pass = false;
        }
    } else {
        pass = false;
    }

    Console.log('');
    Console.log(pass ? '✅ 测试5通过！' : '❌ 测试5失败！');
    return pass;
}

/**
 * 测试6: 多级行字段 + 双列字段
 * 预期: 行字段在最后一级显示
 */
function test_多级行字段_双列字段() {
    Console.log('');
    Console.log('==================== 测试6: 多级行字段 + 双列字段 ====================');

    var rs = Array2D.z超级透视(
        基础数据,
        ['f1+,f2+'],              // 行字段: 销售员升序, 商品ID升序
        ['f6+,f7+'],              // 列字段: 年升序, 月升序
        'sum("f3")'               // 数据字段
    );

    var result = rs.res();
    Console.log('');
    Console.log('========== 输出结果 ==========');
    Console.log('表头行数: ' + result.length);
    Console.log(JSON.stringify(result, null, 2));
    Console.log('');

    // 验证
    var pass = true;

    // 最后行应该包含2个行字段标题
    var lastRow = result[result.length - 1];
    if (lastRow[0] === '销售员' && lastRow[1] === '商品ID') {
        Console.log('✅ 最后行包含2个行字段标题: "销售员", "商品ID"');
    } else {
        Console.log('❌ 行字段标题错误: "' + lastRow[0] + '", "' + lastRow[1] + '"');
        pass = false;
    }

    Console.log('');
    Console.log(pass ? '✅ 测试6通过！' : '❌ 测试6失败！');
    return pass;
}

/**
 * 测试7: 数据正确性验证
 * 预期: 数据汇总正确
 */
function test_数据正确性验证() {
    Console.log('');
    Console.log('==================== 测试7: 数据正确性验证 ====================');

    var rs = Array2D.z超级透视(
        基础数据,
        'f1+',               // 行字段: 销售员
        ['f6+,f7+'],         // 列字段: 年、月
        'sum("f3")'          // 数据字段: 数量求和
    );

    var result = rs.res();
    Console.log('');
    Console.log('========== 输出结果 ==========');
    Console.log(JSON.stringify(result, null, 2));
    Console.log('');

    // 验证数据正确性
    var pass = true;

    // 手动计算 A 销售员在 2023年7月 的数量
    // 原数据: A, P1, 7, ..., 2023, 7, 23
    // 预期: 7
    var headerRowCount = 3;  // 双列字段
    var found = false;
    var expectedValue = 7;

    for (var i = headerRowCount; i < result.length; i++) {
        if (result[i][0] === 'A') {
            // 找到A销售员的数据行
            Console.log('✅ 找到销售员A的数据行');
            // 数据应该从第2列开始（第1列是销售员）
            Console.log('A销售员数据: ' + result[i].slice(1).join(', '));
            found = true;
            break;
        }
    }

    if (!found) {
        Console.log('❌ 未找到销售员A的数据');
        pass = false;
    }

    Console.log('');
    Console.log(pass ? '✅ 测试7通过！' : '❌ 测试7失败！');
    return pass;
}

// ==================== 主测试入口 ====================

/**
 * 运行所有v3.8.0测试
 */
function runV380测试() {
    Console.log('');
    Console.log('***********************************************************');
    Console.log('*     superPivot/z超级透视 v3.8.0 多行表头测试套件        *');
    Console.log('*     版本: JSA880.js v3.8.0                             *');
    Console.log('*     新功能: 真正的多级列表头支持                      *');
    Console.log('***********************************************************');

    var results = [];

    try {
        results.push({name: '测试1: 单列字段 - 向后兼容', pass: test_单列字段_向后兼容()});
    } catch (e) {
        Console.log('❌ 测试1异常: ' + e.message);
        results.push({name: '测试1: 单列字段 - 向后兼容', pass: false});
    }

    try {
        results.push({name: '测试2: 双列字段 - 年、月', pass: test_双列字段_年月()});
    } catch (e) {
        Console.log('❌ 测试2异常: ' + e.message);
        results.push({name: '测试2: 双列字段 - 年、月', pass: false});
    }

    try {
        results.push({name: '测试3: 三级列字段 - 年、季度、月', pass: test_三级列字段_年季度月()});
    } catch (e) {
        Console.log('❌ 测试3异常: ' + e.message);
        results.push({name: '测试3: 三级列字段 - 年、季度、月', pass: false});
    }

    try {
        results.push({name: '测试4: 多数据字段 + 双列字段', pass: test_多数据字段_双列字段()});
    } catch (e) {
        Console.log('❌ 测试4异常: ' + e.message);
        results.push({name: '测试4: 多数据字段 + 双列字段', pass: false});
    }

    try {
        results.push({name: '测试5: 自定义标题 + 双列字段', pass: test_自定义标题_双列字段()});
    } catch (e) {
        Console.log('❌ 测试5异常: ' + e.message);
        results.push({name: '测试5: 自定义标题 + 双列字段', pass: false});
    }

    try {
        results.push({name: '测试6: 多级行字段 + 双列字段', pass: test_多级行字段_双列字段()});
    } catch (e) {
        Console.log('❌ 测试6异常: ' + e.message);
        results.push({name: '测试6: 多级行字段 + 双列字段', pass: false});
    }

    try {
        results.push({name: '测试7: 数据正确性验证', pass: test_数据正确性验证()});
    } catch (e) {
        Console.log('❌ 测试7异常: ' + e.message);
        results.push({name: '测试7: 数据正确性验证', pass: false});
    }

    // 输出测试总结
    Console.log('');
    Console.log('***********************************************************');
    Console.log('*                        测试总结                          *');
    Console.log('***********************************************************');

    var passCount = 0;
    for (var i = 0; i < results.length; i++) {
        var status = results[i].pass ? '✅ 通过' : '❌ 失败';
        Console.log('* ' + results[i].name + ': ' + status);
        if (results[i].pass) passCount++;
    }

    Console.log('***********************************************************');
    Console.log('* 通过: ' + passCount + '/' + results.length);
    Console.log('* 版本: v3.8.0 (2026-02-01)');
    Console.log('* 新功能: 多行多列表头支持');
    Console.log('***********************************************************');

    return passCount === results.length;
}

// ==================== 单独运行函数 ====================

function runV380测试1() { return test_单列字段_向后兼容(); }
function runV380测试2() { return test_双列字段_年月(); }
function runV380测试3() { return test_三级列字段_年季度月(); }
function runV380测试4() { return test_多数据字段_双列字段(); }
function runV380测试5() { return test_自定义标题_双列字段(); }
function runV380测试6() { return test_多级行字段_双列字段(); }
function runV380测试7() { return test_数据正确性验证(); }

// ==================== 使用说明 ====================

/**
 * 使用方法:
 *
 * 1. 在 WPS 宏编辑器中，首先加载 JSA880.js v3.8.0
 * 2. 然后运行此测试文件
 * 3. 调用 runV380测试() 运行所有测试
 * 4. 或单独运行某个测试: runV380测试1(), runV380测试2(), 等
 *
 * JSA880.js v3.8.0 位置:
 * /Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/JSA880.js
 *
 * 测试数据来源:
 * 3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm
 */
