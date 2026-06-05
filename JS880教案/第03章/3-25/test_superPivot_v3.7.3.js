/**
 * superPivot/z超级透视 API 测试
 * 基于官方 API 文档: https://vbayyds.com/api/jsa880/Array2D.html#.superPivot
 *
 * 测试目的: 验证修复后的 superPivot 函数输出与官方示例一致
 */

// ==================== 辅助函数 ====================

/**
 * 比较两个数组是否相等
 * @param {Array} actual - 实际输出
 * @param {Array} expected - 预期输出
 * @returns {Object} 比较结果
 */
function compareArrays(actual, expected) {
    var result = {
        pass: true,
        message: '',
        differences: []
    };

    // 检查行数
    if (actual.length !== expected.length) {
        result.pass = false;
        result.message = '行数不匹配: 实际=' + actual.length + ', 预期=' + expected.length;
        return result;
    }

    // 检查每行每列
    for (var i = 0; i < actual.length; i++) {
        var actualRow = actual[i];
        var expectedRow = expected[i];

        if (actualRow.length !== expectedRow.length) {
            result.pass = false;
            result.differences.push('第' + (i+1) + '行列数不匹配: 实际=' + actualRow.length + ', 预期=' + expectedRow.length);
            continue;
        }

        for (var j = 0; j < actualRow.length; j++) {
            var actualVal = actualRow[j];
            var expectedVal = expectedRow[j];

            // 数值比较（考虑浮点数精度）
            if (typeof actualVal === 'number' && typeof expectedVal === 'number') {
                if (Math.abs(actualVal - expectedVal) > 0.0001) {
                    result.pass = false;
                    result.differences.push('[' + (i+1) + ',' + (j+1) + '] 实际=' + actualVal + ', 预期=' + expectedVal);
                }
            } else if (actualVal !== expectedVal) {
                result.pass = false;
                result.differences.push('[' + (i+1) + ',' + (j+1) + '] 实际="' + actualVal + '", 预期="' + expectedVal + '"');
            }
        }
    }

    if (result.pass) {
        result.message = '✅ 测试通过！输出与预期完全一致。';
    } else {
        result.message = '❌ 测试失败！发现 ' + result.differences.length + ' 处差异。';
    }

    return result;
}

/**
 * 格式化输出数组
 * @param {Array} arr - 数组
 * @param {String} title - 标题
 */
function printArray(arr, title) {
    Console.log('');
    Console.log('========== ' + title + ' ==========');
    Console.log(JSON.stringify(arr, null, 0));
    Console.log('');
}

// ==================== 测试用例 ====================

/**
 * 测试1: 基本透视 - 带排序符号
 * 行字段: f1+ (第1列升序), f2- (第2列降序)
 * 列字段: f5+ (第5列升序), f6+ (第6列升序)
 * 数据字段: count(), sum("f3")
 */
function test_superPivot_example1() {
    Console.log('');
    Console.log('==================== 测试1: 基本透视（带排序符号）====================');

    var arr = [
        ["销售员","商品ID","数量","数量","价格","年","月","日"],
        [1,"product3","中国",29,10,2023,5,23],
        [2,"product2","德国",20,5,2023,4,20],
        [3,"product1","英国",19,1,2023,5,23],
        [4,"product3","英国",24,10,2023,6,28],
        [5,"product2","中国",30,5,2023,4,19]
    ];

    // 执行 superPivot
    var rs = Array2D.z超级透视(arr, ['f1+,f2-'], ['f5+,f6+'], ['count(),sum("f3")']);
    var actual = rs.res();

    // 预期输出（来自官方API文档）
    var expected = [
        ["","价格",1,1,5,5,10,10],
        ["","年",2023,2023,2023,2023,2023,2023],
        ["销售员","商品ID","计数","求和","计数","求和","计数","求和"],
        [1,"product3","","","","",1,0],
        [2,"product2","","",1,0,"",""],
        [3,"product1",1,0,"","","",""],
        [4,"product3","","","","",1,0],
        [5,"product2","","",1,0,"",""]
    ];

    // 输出实际结果
    printArray(actual, '实际输出');

    // 输出预期结果
    printArray(expected, '预期输出');

    // 比较结果
    var comparison = compareArrays(actual, expected);
    Console.log(comparison.message);

    if (!comparison.pass) {
        Console.log('');
        Console.log('差异详情:');
        for (var i = 0; i < Math.min(comparison.differences.length, 20); i++) {
            Console.log('  ' + comparison.differences[i]);
        }
        if (comparison.differences.length > 20) {
            Console.log('  ... 还有 ' + (comparison.differences.length - 20) + ' 处差异');
        }
    }

    return comparison.pass;
}

/**
 * 测试2: 带自定义标题的透视
 * 行字段: f1,f5,f6 (第1,5,6列) 标题: "prod,year,month"
 * 列字段: f2 (第2列) 标题: "country"
 * 数据字段: count(),sum("f3"),average("f4") 标题: "count,sum,avg"
 */
function test_superPivot_example2() {
    Console.log('');
    Console.log('==================== 测试2: 带自定义标题的透视 ====================');

    var arr = [
        ["销售员","商品ID","数量","数量","价格","年","月","日"],
        [1,"product3","中国",29,10,2023,5,23],
        [2,"product2","德国",20,5,2023,4,20],
        [3,"product1","英国",19,1,2023,5,23],
        [4,"product3","英国",24,10,2023,6,28],
        [5,"product2","中国",30,5,2023,4,19]
    ];

    // 执行 superPivot
    var rs = Array2D.z超级透视(arr,
        ['f1,f5,f6','prod,year,month'],
        ['f2','country'],
        ['count(),sum("f3"),average("f4")','count,sum,avg']
    );
    var actual = rs.res();

    // 预期输出（来自官方API文档）
    var expected = [
        ["","","country","product1","product1","product1","product2","product2","product2","product3","product3","product3"],
        ["prod","year","month","count","sum","avg","count","sum","avg","count","sum","avg"],
        [1,10,2023,"","","","","","",1,0,29],
        [2,5,2023,"","","",1,0,20,"","",""],
        [3,1,2023,1,0,19,"","","","","",""],
        [4,10,2023,"","","","","","",1,0,24],
        [5,5,2023,"","","",1,0,30,"","",""]
    ];

    // 输出实际结果
    printArray(actual, '实际输出');

    // 输出预期结果
    printArray(expected, '预期输出');

    // 比较结果
    var comparison = compareArrays(actual, expected);
    Console.log(comparison.message);

    if (!comparison.pass) {
        Console.log('');
        Console.log('差异详情:');
        for (var i = 0; i < Math.min(comparison.differences.length, 20); i++) {
            Console.log('  ' + comparison.differences[i]);
        }
        if (comparison.differences.length > 20) {
            Console.log('  ... 还有 ' + (comparison.differences.length - 20) + ' 处差异');
        }
    }

    return comparison.pass;
}

/**
 * 测试3: Map 模式 + 回调函数
 * 行字段: f1,f5,f6# (第1,5,6列，第6列按原始顺序) 标题: "期数,年,月"
 * 列字段: f2# (第2列按原始顺序) 标题: "国家"
 * 数据字段: 回调函数 [g=>g.count(), g=>g.sum("f3"), g=>g.average("f4")]
 * 输出: 'map' 模式返回 SuperMap
 */
function test_superPivot_example3() {
    Console.log('');
    Console.log('==================== 测试3: Map模式 + 回调函数 ====================');

    var arr = [
        ["0","1","2","3","4","5","6","7"],
        ["销售员","商品ID","数量","数量","价格","年","月","日"],
        [1,"product3","中国",29,10,2023,5,23],
        [2,"product2","德国",20,5,2023,4,20],
        [3,"product1","英国",19,1,2023,5,23],
        [4,"product3","英国",24,10,2023,6,28],
        [5,"product2","中国",30,5,2023,4,19]
    ];

    // 执行 superPivot - map 模式
    var rs = Array2D.z超级透视(arr,
        ['f1,f5,f6#','期数,年,月'],
        ['f2#','国家'],
        [[g=>g.count(),g=>g.sum("f3"),g=>g.average("f4")],'计数,求和,平均'],
        2,
        'map'
    );

    // 转成 SuperMap
    rs = SuperMap.fromMap(rs);

    // 输出结果
    Console.log('');
    Console.log('========== 实际输出（Map模式）==========');
    Console.log('SuperMap.all 的键:');
    var allKeys = Object.keys(rs.all);
    for (var i = 0; i < Math.min(allKeys.length, 10); i++) {
        Console.log('  [' + i + '] ' + allKeys[i]);
    }
    Console.log('');

    // 验证关键字段
    // Map键格式: '01L0001 1@^@10@^@2023@^@product2' (注意: padStart(4, '0') 产生4位数字)
    var expectedKeys = [
        '1@^@10@^@2023@^@product3',
        '2@^@5@^@2023@^@product2',
        '3@^@1@^@2023@^@product1',
        '4@^@10@^@2023@^@product3',
        '5@^@5@^@2023@^@product2'
    ];

    var pass = true;
    Console.log('========== 验证 Map 键 ==========');
    for (var i = 0; i < expectedKeys.length; i++) {
        var key = expectedKeys[i];
        // 使用 padStart(4, '0') 产生 '0001', '0002' 等（4位数字）
        var seq = String(i + 1);
        while (seq.length < 4) {
            seq = '0' + seq;
        }
        var sortKey = '01L' + seq + ' ' + key;
        if (rs.all[sortKey]) {
            Console.log('✅ 键存在: ' + key + ' (SuperMap键: ' + sortKey + ')');
        } else {
            Console.log('❌ 键缺失: ' + key + ' (期望SuperMap键: ' + sortKey + ')');
            // 输出实际可用的键供调试
            Console.log('   实际可用键: ' + JSON.stringify(Object.keys(rs.all).slice(0, 5)));
            pass = false;
            break;  // 第一次失败就停止，便于调试
        }
    }

    // 验证聚合结果
    Console.log('');
    Console.log('========== 验证聚合结果 ==========');

    var expectedAgg = [
        [1, 0, 29],  // 1, product3
        [1, 0, 20],  // 2, product2
        [1, 0, 19],  // 3, product1
        [1, 0, 24],  // 4, product3
        [1, 0, 30]   // 5, product2
    ];

    for (var j = 0; j < expectedKeys.length; j++) {
        var key = expectedKeys[j];
        // 使用 padStart(4, '0') 产生 '0001', '0002' 等（4位数字）
        var seq = String(j + 1);
        while (seq.length < 4) {
            seq = '0' + seq;
        }
        var sortKey = '01L' + seq + ' ' + key;
        var item = rs.all[sortKey];
        if (item) {
            var actualAgg = item.agg;
            var expAgg = expectedAgg[j];
            var match = true;
            for (var k = 0; k < actualAgg.length; k++) {
                if (actualAgg[k] !== expAgg[k]) {
                    match = false;
                    break;
                }
            }
            if (match) {
                Console.log('✅ 聚合正确: ' + key + ' -> ' + JSON.stringify(actualAgg));
            } else {
                Console.log('❌ 聚合错误: ' + key + ' 实际=' + JSON.stringify(actualAgg) + ', 预期=' + JSON.stringify(expAgg));
                pass = false;
            }
        } else {
            Console.log('❌ 未找到键: ' + key + ' (SuperMap键: ' + sortKey + ')');
            pass = false;
        }
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
 * 测试4: 额外测试 - 验证列字段标题
 */
function test_superPivot_columnTitles() {
    Console.log('');
    Console.log('==================== 测试4: 验证列字段标题 ====================');

    var arr = [
        ["销售员","商品ID","数量","数量","价格","年","月","日"],
        [1,"product3","中国",29,10,2023,5,23],
        [2,"product2","德国",20,5,2023,4,20]
    ];

    // 执行 superPivot
    var rs = Array2D.z超级透视(arr, ['f1+'], ['f5+,f6+'], ['count()']);
    var actual = rs.res();

    Console.log('');
    Console.log('========== 实际输出 ==========');
    Console.log(JSON.stringify(actual, null, 0));
    Console.log('');

    // 验证前两行的列字段标题
    var pass = true;

    // 第1行: 应该有 "价格" 标题
    if (actual[0][1] === '价格') {
        Console.log('✅ 第1行列字段标题正确: "价格"');
    } else {
        Console.log('❌ 第1行列字段标题错误: 实际="' + actual[0][1] + '", 预期="价格"');
        pass = false;
    }

    // 第2行: 应该有 "年" 标题
    if (actual[1][1] === '年') {
        Console.log('✅ 第2行列字段标题正确: "年"');
    } else {
        Console.log('❌ 第2行列字段标题错误: 实际="' + actual[1][1] + '", 预期="年"');
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

// ==================== 主测试入口 ====================

/**
 * 运行所有测试
 */
function runAllTests() {
    Console.log('');
    Console.log('***********************************************************');
    Console.log('*     superPivot/z超级透视 API 测试套件                   *');
    Console.log('*     基于: https://vbayyds.com/api/jsa880/               *');
    Console.log('*     版本: JSA880.js v3.7.3                              *');
    Console.log('***********************************************************');

    var results = [];

    // 运行测试1
    try {
        results.push({name: '测试1: 基本透视（带排序符号）', pass: test_superPivot_example1()});
    } catch (e) {
        Console.log('❌ 测试1异常: ' + e.message);
        results.push({name: '测试1: 基本透视（带排序符号）', pass: false});
    }

    // 运行测试2
    try {
        results.push({name: '测试2: 带自定义标题的透视', pass: test_superPivot_example2()});
    } catch (e) {
        Console.log('❌ 测试2异常: ' + e.message);
        results.push({name: '测试2: 带自定义标题的透视', pass: false});
    }

    // 运行测试3
    try {
        results.push({name: '测试3: Map模式 + 回调函数', pass: test_superPivot_example3()});
    } catch (e) {
        Console.log('❌ 测试3异常: ' + e.message);
        results.push({name: '测试3: Map模式 + 回调函数', pass: false});
    }

    // 运行测试4
    try {
        results.push({name: '测试4: 验证列字段标题', pass: test_superPivot_columnTitles()});
    } catch (e) {
        Console.log('❌ 测试4异常: ' + e.message);
        results.push({name: '测试4: 验证列字段标题', pass: false});
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
function runTest1() {
    return test_superPivot_example1();
}

/**
 * 只运行测试2
 */
function runTest2() {
    return test_superPivot_example2();
}

/**
 * 只运行测试3
 */
function runTest3() {
    return test_superPivot_example3();
}

/**
 * 只运行测试4
 */
function runTest4() {
    return test_superPivot_columnTitles();
}

// ==================== 导出说明 ====================

/**
 * 使用方法:
 *
 * 1. 运行所有测试: runAllTests()
 * 2. 单独运行测试: runTest1(), runTest2(), runTest3(), runTest4()
 *
 * 在 WPS 宏编辑器中执行:
 * - 将此代码复制到宏模块
 * - 调用 runAllTests() 运行所有测试
 * - 查看控制台输出对比结果
 */
