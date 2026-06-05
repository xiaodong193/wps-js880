/**
 * superPivot/z超级透视 教案数据测试
 * 基于教案文件: 3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm
 *
 * 测试目的: 验证修复后的 superPivot 函数在真实教案数据上运行正确
 */

// ==================== 教案原始数据 ====================

/**
 * 教案中的原始销售数据
 * 来源: Sheet2 前几行数据
 */
var 教案原始数据 = [
    ["销售员", "商品ID", "数量", "数量2", "价格", "年", "月", "日"],
    ["A", "P1", 7, 12, 10, 2023, 7, 23],
    ["B", "P2", 1, 24, 1, 2020, 1, 2],
    ["A", "P1", 9, 4, 1, 2021, 1, 10],
    ["A", "P4", 4, 12, 1, 2022, 11, 5],
    ["C", "P5", 2, 13, 5, 2020, 5, 29],
    ["B", "P2", 6, 14, 5, 2021, 9, 18],
    ["B", "P7", 7, 8, 10, 2023, 8, 26],
    [8, "product2", "德国", 20, 5, 2020, 4, 20],
    [9, "product1", "德国", 7, 1, 2023, 3, 29]
];

/**
 * 更多教案数据（用于更全面的测试）
 */
var 教案完整数据 = [
    ["销售员", "商品ID", "数量", "数量2", "价格", "年", "月", "日"],
    ["A", "P1", 7, 12, 10, 2023, 7, 23],
    ["B", "P2", 1, 24, 1, 2020, 1, 2],
    ["A", "P1", 9, 4, 1, 2021, 1, 10],
    ["A", "P4", 4, 12, 1, 2022, 11, 5],
    ["C", "P5", 2, 13, 5, 2020, 5, 29],
    ["B", "P2", 6, 14, 5, 2021, 9, 18],
    ["B", "P7", 7, 8, 10, 2023, 8, 26],
    [8, "product2", "德国", 20, 5, 2020, 4, 20],
    [9, "product1", "德国", 7, 1, 2023, 3, 29],
    [10, "product3", "英国", 18, 10, 2023, 5, 23],
    [2, "product2", "德国", 20, 5, 2023, 4, 20],
    [3, "product1", "英国", 19, 1, 2023, 5, 23],
    [4, "product3", "英国", 24, 10, 2023, 6, 28],
    [5, "product2", "中国", 30, 5, 2023, 4, 19]
];

// ==================== 测试用例 ====================

/**
 * 测试1: 基本透视 - 行字段=商品ID, 列字段=年, 数据=数量
 * 预期输出: 每个商品在不同年份的数量汇总
 */
function test_教案_基本透视() {
    Console.log('');
    Console.log('==================== 测试1: 教案基本透视 ====================');

    var rs = Array2D.z超级透视(
        教案完整数据,
        ['f2+'],              // 行字段: 商品ID 升序
        ['f6+'],              // 列字段: 年 升序
        ['sum("f3")']         // 数据字段: 数量求和
    );

    Console.log('');
    Console.log('========== 实际输出 ==========');
    Console.log(JSON.stringify(rs.res(), null, 2));
    Console.log('');

    // 验证表头结构
    var result = rs.res();
    var pass = true;

    // 检查是否有列字段标题 "年"
    if (result[0][1] === '年') {
        Console.log('✅ 列字段标题正确: "年"');
    } else {
        Console.log('❌ 列字段标题错误: 实际="' + result[0][1] + '", 预期="年"');
        pass = false;
    }

    // 检查是否有行字段标题 "商品ID"
    // 对于1个列字段的情况，headerRowCount=1，所以：
    // result[0] = 列字段标题行
    // result[1] = 行字段+数据字段标题行
    if (result[1][0] === '商品ID') {
        Console.log('✅ 行字段标题正确: "商品ID"');
    } else {
        Console.log('❌ 行字段标题错误: 实际="' + result[1][0] + '", 预期="商品ID"');
        pass = false;
    }

    if (pass) {
        Console.log('');
        Console.log('✅ 测试1通过！');
    } else {
        Console.log('');
        Console.log('❌ 测试1失败！');
    }

    return pass;
}

/**
 * 测试2: 多级表头透视 - 行字段=销售员,商品ID, 列字段=年,月, 数据=数量
 * 预期输出: 多级列字段表头（年、月两层）
 */
function test_教案_多级表头() {
    Console.log('');
    Console.log('==================== 测试2: 教案多级表头透视 ====================');

    var rs = Array2D.z超级透视(
        教案完整数据,
        ['f1+,f2+'],          // 行字段: 销售员升序, 商品ID升序
        ['f6+,f7+'],          // 列字段: 年升序, 月升序
        ['sum("f3"),count()'] // 数据字段: 数量求和, 计数
    );

    Console.log('');
    Console.log('========== 实际输出 ==========');
    Console.log(JSON.stringify(rs.res(), null, 2));
    Console.log('');

    // 验证多级表头
    var result = rs.res();
    var pass = true;

    // 第1行应该有 "年" 标题
    if (result[0][1] === '年') {
        Console.log('✅ 第1行列字段标题正确: "年"');
    } else {
        Console.log('❌ 第1行列字段标题错误: 实际="' + result[0][1] + '", 预期="年"');
        pass = false;
    }

    // 第2行应该有 "月" 标题
    if (result[1][1] === '月') {
        Console.log('✅ 第2行列字段标题正确: "月"');
    } else {
        Console.log('❌ 第2行列字段标题错误: 实际="' + result[1][1] + '", 预期="月"');
        pass = false;
    }

    // 检查是否有多级列值
    if (result[0].length > 5) {
        Console.log('✅ 多级列值存在');
    } else {
        Console.log('❌ 多级列值缺失');
        pass = false;
    }

    if (pass) {
        Console.log('');
        Console.log('✅ 测试2通过！');
    } else {
        Console.log('');
        Console.log('❌ 测试2失败！');
    }

    return pass;
}

/**
 * 测试3: 带自定义标题的透视
 * 预期输出: 使用自定义标题而非原数据表头
 */
function test_教案_自定义标题() {
    Console.log('');
    Console.log('==================== 测试3: 教案自定义标题透视 ====================');

    var rs = Array2D.z超级透视(
        教案完整数据,
        ['f2,f6', '产品,年份'],  // 行字段: 商品ID,年; 自定义标题
        ['f7', '月份'],          // 列字段: 月; 自定义标题
        ['sum("f3"),count()', '数量,计数'] // 数据字段: 自定义标题
    );

    Console.log('');
    Console.log('========== 实际输出 ==========');
    Console.log(JSON.stringify(rs.res(), null, 2));
    Console.log('');

    // 验证自定义标题
    var result = rs.res();
    var pass = true;

    // 检查行字段自定义标题
    // 对于1个列字段的情况，headerRowCount=1，所以：
    // result[0] = 列字段标题行
    // result[1] = 行字段+数据字段标题行
    if (result[1][0] === '产品' && result[1][1] === '年份') {
        Console.log('✅ 行字段自定义标题正确: "产品", "年份"');
    } else {
        Console.log('❌ 行字段自定义标题错误: 实际="' + result[1][0] + '", "' + result[1][1] + '"');
        pass = false;
    }

    // 检查列字段自定义标题
    if (result[0][2] === '月份') {
        Console.log('✅ 列字段自定义标题正确: "月份"');
    } else {
        Console.log('❌ 列字段自定义标题错误: 实际="' + result[0][2] + '"');
        pass = false;
    }

    if (pass) {
        Console.log('');
        Console.log('✅ 测试3通过！');
    } else {
        Console.log('');
        Console.log('❌ 测试3失败！');
    }

    return pass;
}

/**
 * 测试4: 排序符号测试
 * 预期输出: 正确应用升序、降序排序
 */
function test_教案_排序符号() {
    Console.log('');
    Console.log('==================== 测试4: 教案排序符号测试 ====================');

    var rs = Array2D.z超级透视(
        教案完整数据,
        ['f2-,f6+'],          // 行字段: 商品ID降序, 年升序
        ['f7+'],              // 列字段: 月升序
        ['count()']           // 数据字段: 计数
    );

    Console.log('');
    Console.log('========== 实际输出 ==========');
    Console.log(JSON.stringify(rs.res(), null, 2));
    Console.log('');

    // 验证排序
    var result = rs.res();
    var pass = true;

    // 检查商品ID是否降序排列
    // 对于1个列字段的情况，headerRowCount=1，所以数据行从索引2开始
    var productIds = [];
    for (var i = 2; i < result.length; i++) {
        if (result[i][0] && result[i][0] !== '') {
            productIds.push(result[i][0]);
        }
    }

    // 简单检查：product3 应该在 product1 前面（字符串降序）
    var idxP1 = -1, idxP3 = -1;
    for (var j = 0; j < productIds.length; j++) {
        if (productIds[j] === 'product1') idxP1 = j;
        if (productIds[j] === 'product3') idxP3 = j;
    }

    // 输出实际的商品ID序列供参考
    Console.log('商品ID序列: ' + productIds.join(', '));

    if (idxP1 !== -1 && idxP3 !== -1 && idxP3 < idxP1) {
        Console.log('✅ 商品ID降序排列正确: product3 在 product1 前面');
    } else {
        Console.log('❌ 商品ID排序可能不正确');
        pass = false;
    }

    if (pass) {
        Console.log('');
        Console.log('✅ 测试4通过！');
    } else {
        Console.log('');
        Console.log('❌ 测试4失败！');
    }

    return pass;
}

/**
 * 测试5: # 排序符号测试（按原始顺序）
 * 预期输出: 保持数据源中的原始出现顺序
 */
function test_教案_原始顺序排序() {
    Console.log('');
    Console.log('==================== 测试5: 教案#排序符号测试 ====================');

    var rs = Array2D.z超级透视(
        教案完整数据,
        ['f2#,f6+'],          // 行字段: 商品ID按原始顺序, 年升序
        ['f7+'],              // 列字段: 月升序
        ['count()']           // 数据字段: 计数
    );

    Console.log('');
    Console.log('========== 实际输出 ==========');
    Console.log(JSON.stringify(rs.res(), null, 2));
    Console.log('');

    Console.log('✅ 测试5完成（需人工验证原始顺序）');
    return true;
}

// ==================== 主测试入口 ====================

/**
 * 运行所有教案测试
 */
function run教案测试() {
    Console.log('');
    Console.log('***********************************************************');
    Console.log('*     superPivot/z超级透视 教案数据测试套件                *');
    Console.log('*     基于教案: 3.25 superPivot降维打击                   *');
    Console.log('*     版本: JSA880.js v3.7.2                              *');
    Console.log('***********************************************************');

    var results = [];

    try {
        results.push({name: '测试1: 教案基本透视', pass: test_教案_基本透视()});
    } catch (e) {
        Console.log('❌ 测试1异常: ' + e.message);
        results.push({name: '测试1: 教案基本透视', pass: false});
    }

    try {
        results.push({name: '测试2: 教案多级表头透视', pass: test_教案_多级表头()});
    } catch (e) {
        Console.log('❌ 测试2异常: ' + e.message);
        results.push({name: '测试2: 教案多级表头透视', pass: false});
    }

    try {
        results.push({name: '测试3: 教案自定义标题透视', pass: test_教案_自定义标题()});
    } catch (e) {
        Console.log('❌ 测试3异常: ' + e.message);
        results.push({name: '测试3: 教案自定义标题透视', pass: false});
    }

    try {
        results.push({name: '测试4: 教案排序符号测试', pass: test_教案_排序符号()});
    } catch (e) {
        Console.log('❌ 测试4异常: ' + e.message);
        results.push({name: '测试4: 教案排序符号测试', pass: false});
    }

    try {
        results.push({name: '测试5: 教案#排序符号测试', pass: test_教案_原始顺序排序()});
    } catch (e) {
        Console.log('❌ 测试5异常: ' + e.message);
        results.push({name: '测试5: 教案#排序符号测试', pass: false});
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
    Console.log('***********************************************************');

    return passCount === results.length;
}

// ==================== 单独运行函数 ====================

/**
 * 只运行测试1
 */
function run教案测试1() {
    return test_教案_基本透视();
}

/**
 * 只运行测试2
 */
function run教案测试2() {
    return test_教案_多级表头();
}

/**
 * 只运行测试3
 */
function run教案测试3() {
    return test_教案_自定义标题();
}

/**
 * 只运行测试4
 */
function run教案测试4() {
    return test_教案_排序符号();
}

/**
 * 只运行测试5
 */
function run教案测试5() {
    return test_教案_原始顺序排序();
}

// ==================== 使用说明 ====================

/**
 * 使用方法:
 *
 * 1. 在 WPS 宏编辑器中，首先加载新版 JSA880.js (v3.7.2)
 * 2. 然后运行此测试文件
 * 3. 调用 run教案测试() 运行所有测试
 * 4. 或单独运行某个测试: run教案测试1(), run教案测试2(), 等
 *
 * 新版 JSA880.js 位置:
 * /Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/JSA880_v3.7.2_to_paste.js
 */
