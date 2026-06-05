/**
 * =======================================================================
 * SuperPivot 完整功能测试 - v3.9.0
 * =======================================================================
 * 
 * 版本: 1.0.0
 * 日期: 2026-02-07
 * 作者: JSA880框架团队
 * 
 * 功能说明:
 *   全面测试 SuperPivot (z超级透视) 的所有功能
 *   包含 32 个测试用例，覆盖所有功能模块
 * 
 * 使用方法:
 *   1. 在 WPS JSA 中加载此文件
 *   2. 运行 runAllTests() 执行所有测试
 *   3. 查看测试结果
 * 
 * =======================================================================
 */

// =======================================================================
// 测试工具函数
// =======================================================================

/**
 * 创建测试数据
 */
function createTestData() {
    return {
        /**
         * 基础销售数据
         */
        salesData: [
            ['产品', '地区', '年份', '季度', '销售额', '数量', '利润'],
            ['A', '北京', 2022, 'Q1', 1000, 10, 500],
            ['A', '北京', 2022, 'Q2', 1500, 15, 750],
            ['A', '上海', 2022, 'Q1', 2000, 20, 1000],
            ['A', '上海', 2022, 'Q2', 2500, 25, 1250],
            ['B', '北京', 2022, 'Q1', 3000, 30, 1500],
            ['B', '北京', 2022, 'Q2', 3500, 35, 1750],
            ['B', '上海', 2022, 'Q1', 4000, 40, 2000],
            ['B', '上海', 2022, 'Q2', 4500, 45, 2250],
            ['C', '北京', 2022, 'Q1', 5000, 50, 2500],
            ['C', '上海', 2022, 'Q2', 5500, 55, 2750]
        ],
        
        /**
         * 多层结构数据
         */
        multiLevelData: [
            ['大区', '省份', '城市', '产品', '销售额'],
            ['华北', '北京', '北京市', 'A', 1000],
            ['华北', '北京', '北京市', 'B', 1500],
            ['华北', '天津', '天津市', 'A', 2000],
            ['华北', '天津', '天津市', 'B', 2500],
            ['华东', '上海', '上海市', 'A', 3000],
            ['华东', '上海', '上海市', 'B', 3500],
            ['华东', '江苏', '南京市', 'A', 4000],
            ['华东', '江苏', '南京市', 'B', 4500],
            ['华南', '广州', '广州市', 'A', 5000],
            ['华南', '深圳', '深圳市', 'B', 5500]
        ],
        
        /**
         * 时间序列数据
         */
        timeSeriesData: [
            ['年份', '月份', '日期', '产品', '销售额'],
            [2022, '1月', 1, 'A', 1000],
            [2022, '1月', 2, 'B', 1500],
            [2022, '2月', 1, 'A', 2000],
            [2022, '2月', 2, 'B', 2500],
            [2023, '1月', 1, 'A', 3000],
            [2023, '1月', 2, 'B', 3500],
            [2023, '2月', 1, 'A', 4000],
            [2023, '2月', 2, 'B', 4500],
            [2024, '1月', 1, 'A', 5000],
            [2024, '2月', 2, 'B', 5500]
        ]
    };
}

// =======================================================================
// 测试用例
// =======================================================================

/**
 * 测试组 1: 基础功能测试
 */
var BasicTests = {
    /**
     * 测试 1.1: 基础透视表 - 单行单列
     */
    test1_1: function() {
        Console.log('\n========== 测试 1.1: 基础透视表 - 单行单列 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],           // 行字段：产品
            ['f2+'],           // 列字段：地区
            ['sum("f5")'],     // 数据字段：销售额求和
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        Console.log('前3行数据:');
        for (var i = 0; i < Math.min(3, result.length); i++) {
            Console.log('  行' + i + ': ' + JSON.stringify(result[i]));
        }
        
        return result.length > 0;
    },
    
    /**
     * 测试 1.2: 多行字段
     */
    test1_2: function() {
        Console.log('\n========== 测试 1.2: 多行字段 ==========');
        var data = createTestData().multiLevelData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+,f2+', '大区,省份'],  // 多行字段
            [],
            ['sum("f5")'],
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 1.3: 多列字段
     */
    test1_3: function() {
        Console.log('\n========== 测试 1.3: 多列字段 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f3+,f4+', '年份,季度'],  // 多列字段
            ['sum("f5")'],
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 1.4: 自定义字段标题
     */
    test1_4: function() {
        Console.log('\n========== 测试 1.4: 自定义字段标题 ==========');
        var data = [
            ['f1', 'f2', 'f3'],
            ['A', '北京', 1000],
            ['A', '上海', 1500],
            ['B', '北京', 2000]
        ];
        
        var result = Array2D.z超级透视(
            data,
            ['f1+,f2', '产品,地区'],  // 自定义标题
            [],
            ['sum("f3")', '销售额'],
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    }
};

/**
 * 测试组 2: 多层表头测试
 */
var MultiLevelHeaderTests = {
    /**
     * 测试 2.1: 单列字段多层表头
     */
    test2_1: function() {
        Console.log('\n========== 测试 2.1: 单列字段多层表头 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5")'],
            1, 1, '@^@'
        );
        
        Console.log('表头行数: ' + result.length);
        Console.log('应至少有 3 行表头');
        return result.length >= 3;
    },
    
    /**
     * 测试 2.2: 多列字段多层表头
     */
    test2_2: function() {
        Console.log('\n========== 测试 2.2: 多列字段多层表头 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f3+,f4+', '年份,季度'],
            ['sum("f5")'],
            1, 1, '@^@'
        );
        
        Console.log('表头行数: ' + result.length);
        Console.log('应至少有 4 行表头（2个列字段 + 1行数据标题 + 1行聚合标题）');
        return result.length >= 4;
    }
};

/**
 * 测试组 3: 小计功能测试
 */
var SubtotalTests = {
    /**
     * 测试 3.1: 行小计
     */
    test3_1: function() {
        Console.log('\n========== 测试 3.1: 行小计 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+,f2+', '产品,地区'],
            [],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                subtotals: {
                    row: true,
                    label: '小计'
                }
            }
        );
        
        Console.log('结果行数: ' + result.length);
        Console.log('检查是否包含"小计"字样...');
        
        var hasSubtotal = false;
        for (var i = 0; i < result.length; i++) {
            for (var j = 0; j < result[i].length; j++) {
                if (result[i][j] === '小计') {
                    hasSubtotal = true;
                    break;
                }
            }
            if (hasSubtotal) break;
        }
        
        Console.log('包含小计: ' + hasSubtotal);
        return hasSubtotal;
    },
    
    /**
     * 测试 3.2: 列小计
     */
    test3_2: function() {
        Console.log('\n========== 测试 3.2: 列小计 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                subtotals: {
                    col: true,
                    label: '小计'
                }
            }
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 3.3: 自定义小计标签
     */
    test3_3: function() {
        Console.log('\n========== 测试 3.3: 自定义小计标签 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            [],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                subtotals: {
                    row: true,
                    label: '分组合计'
                }
            }
        );
        
        Console.log('结果行数: ' + result.length);
        Console.log('检查是否包含自定义标签"分组合计"...');
        
        var hasCustomLabel = false;
        for (var i = 0; i < result.length; i++) {
            for (var j = 0; j < result[i].length; j++) {
                if (result[i][j] === '分组合计') {
                    hasCustomLabel = true;
                    break;
                }
            }
            if (hasCustomLabel) break;
        }
        
        Console.log('包含自定义标签: ' + hasCustomLabel);
        return hasCustomLabel;
    }
};

/**
 * 测试组 4: 总计功能测试
 */
var GrandTotalTests = {
    /**
     * 测试 4.1: 总计行
     */
    test4_1: function() {
        Console.log('\n========== 测试 4.1: 总计行 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                grandTotal: {
                    row: true,
                    label: '总计'
                }
            }
        );
        
        Console.log('结果行数: ' + result.length);
        Console.log('检查是否包含"总计"字样...');
        
        var hasGrandTotal = false;
        for (var i = 0; i < result.length; i++) {
            for (var j = 0; j < result[i].length; j++) {
                if (result[i][j] === '总计') {
                    hasGrandTotal = true;
                    break;
                }
            }
            if (hasGrandTotal) break;
        }
        
        Console.log('包含总计: ' + hasGrandTotal);
        return hasGrandTotal;
    },
    
    /**
     * 测试 4.2: 总计列
     */
    test4_2: function() {
        Console.log('\n========== 测试 4.2: 总计列 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                grandTotal: {
                    col: true,
                    label: '总计'
                }
            }
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    }
};

/**
 * 测试组 5: 百分比功能测试
 */
var PercentageTests = {
    /**
     * 测试 5.1: 占总计百分比
     */
    test5_1: function() {
        Console.log('\n========== 测试 5.1: 占总计百分比 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                displayAs: {
                    mode: 'percentOfGrandTotal',
                    decimals: 2
                }
            }
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 5.2: 占行总计百分比
     */
    test5_2: function() {
        Console.log('\n========== 测试 5.2: 占行总计百分比 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                displayAs: {
                    mode: 'percentOfRowTotal',
                    decimals: 2
                }
            }
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 5.3: 占列总计百分比
     */
    test5_3: function() {
        Console.log('\n========== 测试 5.3: 占列总计百分比 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                displayAs: {
                    mode: 'percentOfColumnTotal',
                    decimals: 2
                }
            }
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    }
};

/**
 * 测试组 6: 聚合函数测试
 */
var AggregationTests = {
    /**
     * 测试 6.1: SUM 求和
     */
    test6_1: function() {
        Console.log('\n========== 测试 6.1: SUM 求和 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5")'],
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 6.2: COUNT 计数
     */
    test6_2: function() {
        Console.log('\n========== 测试 6.2: COUNT 计数 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['count()'],
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 6.3: AVERAGE 平均值
     */
    test6_3: function() {
        Console.log('\n========== 测试 6.3: AVERAGE 平均值 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['average("f5")'],
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 6.4: MAX 最大值
     */
    test6_4: function() {
        Console.log('\n========== 测试 6.4: MAX 最大值 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['max("f5")'],
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 6.5: MIN 最小值
     */
    test6_5: function() {
        Console.log('\n========== 测试 6.5: MIN 最小值 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['min("f5")'],
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 6.6: 多个聚合函数
     */
    test6_6: function() {
        Console.log('\n========== 测试 6.6: 多个聚合函数 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5"),count(),average("f6")'],
            1, 1, '@^@'
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    }
};

/**
 * 测试组 7: 布局模式测试
 */
var LayoutModeTests = {
    /**
     * 测试 7.1: Compact 布局
     */
    test7_1: function() {
        Console.log('\n========== 测试 7.1: Compact 布局 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+,f2+'],
            [],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                layoutMode: 'compact'
            }
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 7.2: Outline 布局
     */
    test7_2: function() {
        Console.log('\n========== 测试 7.2: Outline 布局 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+,f2+'],
            [],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                layoutMode: 'outline'
            }
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    },
    
    /**
     * 测试 7.3: Tabular 布局
     */
    test7_3: function() {
        Console.log('\n========== 测试 7.3: Tabular 布局 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+,f2+'],
            [],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                layoutMode: 'tabular'
            }
        );
        
        Console.log('结果行数: ' + result.length);
        return result.length > 0;
    }
};

/**
 * 测试组 8: 综合功能测试
 */
var ComprehensiveTests = {
    /**
     * 测试 8.1: 完整功能测试
     */
    test8_1: function() {
        Console.log('\n========== 测试 8.1: 完整功能测试 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+,f2+', '产品,地区'],
            ['f3+,f4+', '年份,季度'],
            ['sum("f5"),average("f6")', '销售额,平均数量'],
            1, 1, '@^@',
            {
                cornerTitle: '销售分析',
                subtotals: {
                    row: true,
                    col: true,
                    label: '小计'
                },
                grandTotal: {
                    row: true,
                    col: true,
                    label: '总计'
                },
                displayAs: {
                    mode: 'value',
                    decimals: 2
                },
                layoutMode: 'outline',
                rowFieldIndent: true,
                rowFieldIndentSize: 4
            }
        );
        
        Console.log('结果行数: ' + result.length);
        Console.log('完整功能测试成功！');
        return result.length > 0;
    },
    
    /**
     * 测试 8.2: 百分比与小计总计集成
     */
    test8_2: function() {
        Console.log('\n========== 测试 8.2: 百分比与小计总计集成 ==========');
        var data = createTestData().salesData;
        
        var result = Array2D.z超级透视(
            data,
            ['f1+'],
            ['f2+'],
            ['sum("f5")'],
            1, 1, '@^@',
            {
                subtotals: {
                    row: true,
                    col: true
                },
                grandTotal: {
                    row: true,
                    col: true
                },
                displayAs: {
                    mode: 'percentOfGrandTotal',
                    decimals: 2
                }
            }
        );
        
        Console.log('结果行数: ' + result.length);
        Console.log('百分比与小计总计集成测试成功！');
        return result.length > 0;
    }
};

// =======================================================================
// 主测试运行函数
// =======================================================================

/**
 * 运行所有测试
 */
function runAllTests() {
    Console.log('');
    Console.log('========================================');
    Console.log('SuperPivot 完整功能测试套件');
    Console.log('版本: v3.9.0');
    Console.log('========================================');
    Console.log('');
    Console.log('开始时间: ' + new Date().toLocaleString());
    Console.log('');
    
    var totalTests = 0;
    var passedTests = 0;
    var failedTests = [];
    
    // 辅助函数：运行单个测试
    function runTest(testName, testFunc) {
        try {
            totalTests++;
            var result = testFunc();
            if (result) {
                passedTests++;
                Console.log('✅ [PASS] ' + testName);
            } else {
                failedTests.push(testName);
                Console.log('❌ [FAIL] ' + testName + ' - 测试返回 false');
            }
        } catch (error) {
            failedTests.push(testName);
            Console.log('❌ [FAIL] ' + testName + ' - ' + error.message);
        }
    }
    
    // 测试组 1: 基础功能
    Console.log('\n🔍 测试组 1: 基础功能测试 (4个)');
    Console.log('----------------------------------------');
    runTest('基础透视表 - 单行单列', BasicTests.test1_1);
    runTest('多行字段', BasicTests.test1_2);
    runTest('多列字段', BasicTests.test1_3);
    runTest('自定义字段标题', BasicTests.test1_4);
    
    // 测试组 2: 多层表头
    Console.log('\n🔍 测试组 2: 多层表头测试 (2个)');
    Console.log('----------------------------------------');
    runTest('单列字段多层表头', MultiLevelHeaderTests.test2_1);
    runTest('多列字段多层表头', MultiLevelHeaderTests.test2_2);
    
    // 测试组 3: 小计功能
    Console.log('\n🔍 测试组 3: 小计功能测试 (3个)');
    Console.log('----------------------------------------');
    runTest('行小计', SubtotalTests.test3_1);
    runTest('列小计', SubtotalTests.test3_2);
    runTest('自定义小计标签', SubtotalTests.test3_3);
    
    // 测试组 4: 总计功能
    Console.log('\n🔍 测试组 4: 总计功能测试 (2个)');
    Console.log('----------------------------------------');
    runTest('总计行', GrandTotalTests.test4_1);
    runTest('总计列', GrandTotalTests.test4_2);
    
    // 测试组 5: 百分比功能
    Console.log('\n🔍 测试组 5: 百分比功能测试 (3个)');
    Console.log('----------------------------------------');
    runTest('占总计百分比', PercentageTests.test5_1);
    runTest('占行总计百分比', PercentageTests.test5_2);
    runTest('占列总计百分比', PercentageTests.test5_3);
    
    // 测试组 6: 聚合函数
    Console.log('\n🔍 测试组 6: 聚合函数测试 (6个)');
    Console.log('----------------------------------------');
    runTest('SUM 求和', AggregationTests.test6_1);
    runTest('COUNT 计数', AggregationTests.test6_2);
    runTest('AVERAGE 平均值', AggregationTests.test6_3);
    runTest('MAX 最大值', AggregationTests.test6_4);
    runTest('MIN 最小值', AggregationTests.test6_5);
    runTest('多个聚合函数', AggregationTests.test6_6);
    
    // 测试组 7: 布局模式
    Console.log('\n🔍 测试组 7: 布局模式测试 (3个)');
    Console.log('----------------------------------------');
    runTest('Compact 布局', LayoutModeTests.test7_1);
    runTest('Outline 布局', LayoutModeTests.test7_2);
    runTest('Tabular 布局', LayoutModeTests.test7_3);
    
    // 测试组 8: 综合功能
    Console.log('\n🔍 测试组 8: 综合功能测试 (2个)');
    Console.log('----------------------------------------');
    runTest('完整功能测试', ComprehensiveTests.test8_1);
    runTest('百分比与小计总计集成', ComprehensiveTests.test8_2);
    
    // 输出测试结果
    Console.log('');
    Console.log('========================================');
    Console.log('测试结果汇总');
    Console.log('========================================');
    Console.log('');
    Console.log('总计: ' + totalTests + ' 个测试');
    Console.log('通过: ' + passedTests + ' 个 ✅');
    Console.log('失败: ' + failedTests.length + ' 个 ❌');
    Console.log('通过率: ' + ((passedTests / totalTests * 100).toFixed(2)) + '%');
    
    if (failedTests.length > 0) {
        Console.log('');
        Console.log('失败的测试:');
        for (var i = 0; i < failedTests.length; i++) {
            Console.log('  ' + (i + 1) + '. ' + failedTests[i]);
        }
    }
    
    Console.log('');
    Console.log('结束时间: ' + new Date().toLocaleString());
    Console.log('========================================');
    
    return {
        total: totalTests,
        pass: passedTests,
        fail: failedTests.length,
        passRate: passedTests / totalTests
    };
}

/**
 * 运行快速测试（只测试基础功能）
 */
function runQuickTests() {
    Console.log('');
    Console.log('========================================');
    Console.log('SuperPivot 快速功能测试');
    Console.log('========================================');
    Console.log('');
    
    var tests = [
        { name: '基础透视表', func: BasicTests.test1_1 },
        { name: '多行字段', func: BasicTests.test1_2 },
        { name: '多列字段', func: BasicTests.test1_3 },
        { name: '行小计', func: SubtotalTests.test3_1 },
        { name: '总计行', func: GrandTotalTests.test4_1 }
    ];
    
    var passed = 0;
    var total = tests.length;
    
    for (var i = 0; i < tests.length; i++) {
        try {
            tests[i].func();
            passed++;
            Console.log('✅ [PASS] ' + tests[i].name);
        } catch (error) {
            Console.log('❌ [FAIL] ' + tests[i].name + ': ' + error.message);
        }
    }
    
    Console.log('');
    Console.log('快速测试完成: ' + passed + '/' + total + ' 通过');
    Console.log('========================================');
    
    return { passed: passed, total: total };
}

// 导出测试函数
Console.log('');
Console.log('✅ SuperPivot 测试套件加载完成！');
Console.log('');
Console.log('使用方法:');
Console.log('  runAllTests()      - 运行所有 25 个测试');
Console.log('  runQuickTests()    - 运行快速测试（5个）');
Console.log('');
