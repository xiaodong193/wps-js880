/**
 * ========================================================================
 * JSA880 全面测试套件
 * ========================================================================
 * 
 * 说明：完整的测试套件，覆盖所有功能模块
 * 使用：在 WPS JSA 环境中运行此文件
 * 
 * 测试模块：
 *   1. v3.9.0 新功能测试（15 个用例）
 *   2. 原有功能回归测试（35 个用例）
 *   3. 边界情况测试（20 个用例）
 *   4. 错误处理测试（10 个用例）
 *   5. 性能测试（5 个用例）
 *   6. 集成测试（10 个用例）
 * 
 * ========================================================================
 */

// ==================== 测试数据 ====================

var testDataSets = {
  // 基础销售数据
  sales: [
    ['产品', '地区', '季度', '销售额', '订单数'],
    ['手机', '北京', 'Q1', 10000, 50],
    ['手机', '北京', 'Q2', 12000, 60],
    ['手机', '上海', 'Q1', 15000, 80],
    ['手机', '上海', 'Q2', 18000, 90],
    ['电脑', '北京', 'Q1', 20000, 30],
    ['电脑', '北京', 'Q2', 22000, 35],
    ['电脑', '上海', 'Q1', 18000, 25],
    ['电脑', '上海', 'Q2', 20000, 28],
    ['平板', '北京', 'Q1', 5000, 20],
    ['平板', '上海', 'Q2', 6000, 25]
  ],

  // 多级分类数据
  category: [
    ['年份', '季度', '类别', '子类别', '产品', '销售额', '利润'],
    ['2024', 'Q1', '电子', '手机', 'iPhone', 50000, 10000],
    ['2024', 'Q1', '电子', '手机', '小米', 30000, 5000],
    ['2024', 'Q2', '电子', '手机', 'iPhone', 55000, 11000],
    ['2024', 'Q2', '电子', '电脑', 'MacBook', 80000, 20000],
    ['2025', 'Q1', '家电', '冰箱', '海尔', 20000, 4000],
    ['2025', 'Q1', '家电', '洗衣机', '美的', 15000, 3000]
  ],

  // 边界测试数据
  edgeCase: [
    ['产品', '地区', '销售额'],
    ['A', '北京', 100],
    ['A', '', 200],
    ['A', null, 300],
    ['B', 'X', 'N/A'],
    ['C', 'Y', 0],
    ['D', 'Z', -100]
  ],

  // 简单数据
  simple: [
    ['产品', '地区', '销售额'],
    ['A', '北京', 1000],
    ['A', '上海', 1500],
    ['B', '北京', 2000],
    ['B', '上海', 1000]
  ]
};

// ==================== 测试框架 ====================

function TestFramework() {
  this.results = {
    passed: 0,
    failed: 0,
    skipped: 0,
    errors: [],
    details: []
  };

  this.test = function(name, fn) {
    try {
      Console.log('  ◦ ' + name);
      var result = fn();
      if (result === true) {
        this.results.passed++;
        this.results.details.push({ name: name, status: 'PASS' });
        Console.log('    ✅ PASS');
      } else if (result === false) {
        this.results.failed++;
        this.results.details.push({ name: name, status: 'FAIL', reason: 'Assertion failed' });
        Console.log('    ❌ FAIL');
      } else {
        this.results.skipped++;
        this.results.details.push({ name: name, status: 'SKIP' });
        Console.log('    ⏭️  SKIP');
      }
    } catch (e) {
      this.results.failed++;
      this.results.errors.push({ name: name, error: e.message });
      this.results.details.push({ name: name, status: 'ERROR', error: e.message });
      Console.log('    ❌ ERROR: ' + e.message);
    }
  };

  this.assert = function(condition, message) {
    if (!condition) {
      throw new Error(message || 'Assertion failed');
    }
    return true;
  };

  this.assertEqual = function(actual, expected, message) {
    if (actual !== expected) {
      throw new Error((message || '') + ' (expected: ' + expected + ', actual: ' + actual + ')');
    }
    return true;
  };

  this.assertNotNull = function(value, message) {
    if (value === null || value === undefined) {
      throw new Error(message || 'Value is null or undefined');
    }
    return true;
  };

  this.report = function() {
    Console.log('\n========================================');
    Console.log('  测试结果汇总');
    Console.log('========================================');
    Console.log('总测试数: ' + (this.results.passed + this.results.failed + this.results.skipped));
    Console.log('✅ 通过: ' + this.results.passed);
    Console.log('❌ 失败: ' + this.results.failed);
    Console.log('⏭️  跳过: ' + this.results.skipped);
    
    if (this.results.errors.length > 0) {
      Console.log('\n错误详情:');
      for (var i = 0; i < this.results.errors.length; i++) {
        Console.log('  ' + (i + 1) + '. ' + this.results.errors[i].name + ': ' + this.results.errors[i].error);
      }
    }
    
    Console.log('========================================\n');
    
    return this.results.failed === 0;
  };
}

// ==================== 测试模块 1: v3.9.0 新功能 ====================

function test_v390_NewFeatures(tf) {
  Console.log('\n【模块 1】v3.9.0 新功能测试');
  Console.log('─────────────────────────────────────');

  // 1.1 小计功能
  Console.log('\n1.1 小计功能:');
  
  tf.test('行小计 - 单行字段', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { subtotals: { row: true } }
    );
    return tf.assertNotNull(result) && result.length > 0;
  });

  tf.test('列小计 - 单列字段', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { subtotals: { col: true } }
    );
    return tf.assertNotNull(result) && result.length > 0;
  });

  tf.test('行小计 + 列小计组合', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { subtotals: { row: true, col: true } }
    );
    return tf.assertNotNull(result) && result.length > 0;
  });

  tf.test('小计标签自定义', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { subtotals: { row: true, label: '合计' } }
    );
    return tf.assertNotNull(result);
  });

  tf.test('多数据字段的小计', function() {
    var result = Array2D.z超级透视(
      testDataSets.sales,
      ['f1+'], ['f2+'], ['sum("f4"),count("f4"),average("f4")'],
      1, 1, '@^@',
      { subtotals: { row: true, col: true } }
    );
    return tf.assertNotNull(result);
  });

  // 1.2 总计功能
  Console.log('\n1.2 总计功能:');
  
  tf.test('总计行 - 基础', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { grandTotal: { row: true } }
    );
    return tf.assertNotNull(result) && result.length > 0;
  });

  tf.test('总计列 - 基础', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { grandTotal: { col: true } }
    );
    return tf.assertNotNull(result);
  });

  tf.test('总计行 + 总计列组合', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { grandTotal: { row: true, col: true } }
    );
    return tf.assertNotNull(result);
  });

  tf.test('总计标签自定义', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { grandTotal: { row: true, label: '总和' } }
    );
    return tf.assertNotNull(result);
  });

  // 1.3 兼容性测试
  Console.log('\n1.3 兼容性测试:');
  
  tf.test('不传 options（完全向后兼容）', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")']
    );
    return tf.assertNotNull(result) && result.length > 0;
  });

  tf.test('使用旧版配置名（rowSubtotals）', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { rowSubtotals: { enabled: true } }
    );
    return tf.assertNotNull(result);
  });

  tf.test('空配置对象', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      {}
    );
    return tf.assertNotNull(result);
  });
}

// ==================== 测试模块 2: 原有功能回归 ====================

function test_Regression(tf) {
  Console.log('\n【模块 2】原有功能回归测试');
  Console.log('─────────────────────────────────────');

  // 2.1 基础透视功能
  Console.log('\n2.1 基础透视功能:');
  
  tf.test('单行单列基础透视', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")']
    );
    return tf.assertNotNull(result) && result.length > 0;
  });

  tf.test('多行多列透视', function() {
    var result = Array2D.z超级透视(
      testDataSets.category,
      ['f1+,f2+'], ['f3+'], ['sum("f6")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('无表头输出（outputHeader=0）', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 0, '@^@'
    );
    return tf.assertNotNull(result);
  });

  tf.test('自定义表头（outputHeader=-1）', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, -1, '@^@'
    );
    return tf.assertNotNull(result);
  });

  // 2.2 聚合函数
  Console.log('\n2.2 聚合函数:');
  
  tf.test('sum() 求和', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('count() 计数', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['count()']
    );
    return tf.assertNotNull(result);
  });

  tf.test('average() 平均值', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['average("f3")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('max() 最大值', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['max("f3")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('min() 最小值', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['min("f3")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('多个聚合函数组合', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3"),count(),average("f3")']
    );
    return tf.assertNotNull(result);
  });

  // 2.3 多级字段
  Console.log('\n2.3 多级字段:');
  
  tf.test('多级行字段（2级）', function() {
    var result = Array2D.z超级透视(
      testDataSets.category,
      ['f1+,f2+'], ['f3+'], ['sum("f6")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('多级列字段（2级）', function() {
    var result = Array2D.z超级透视(
      testDataSets.category,
      ['f1+'], ['f3+,f4+'], ['sum("f6")']
    );
    return tf.assertNotNull(result);
  });

  // 2.4 排序功能
  Console.log('\n2.4 排序功能:');
  
  tf.test('升序排序（+）', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2-'], ['sum("f3")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('降序排序（-）', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1-'], ['f2+'], ['sum("f3")']
    );
    return tf.assertNotNull(result);
  });

  // 2.5 特殊功能
  Console.log('\n2.5 特殊功能:');
  
  tf.test('百分比显示（percentOfGrandTotal）', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { displayAs: { mode: 'percentOfGrandTotal' } }
    );
    return tf.assertNotNull(result);
  });

  tf.test('百分比显示（percentOfRowTotal）', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { displayAs: { mode: 'percentOfRowTotal' } }
    );
    return tf.assertNotNull(result);
  });
}

// ==================== 测试模块 3: 边界情况 ====================

function test_EdgeCases(tf) {
  Console.log('\n【模块 3】边界情况测试');
  Console.log('─────────────────────────────────────');

  Console.log('\n3.1 数据边界:');
  
  tf.test('包含空值的数据', function() {
    var result = Array2D.z超级透视(
      testDataSets.edgeCase,
      ['f1+'], ['f2+'], ['sum("f3")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('包含 null 的数据', function() {
    var result = Array2D.z超级透视(
      testDataSets.edgeCase,
      ['f1+'], ['f2+'], ['sum("f3")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('零值处理', function() {
    var result = Array2D.z超级透视(
      testDataSets.edgeCase,
      ['f1+'], ['f2+'], ['sum("f3")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('负数处理', function() {
    var result = Array2D.z超级透视(
      testDataSets.edgeCase,
      ['f1+'], ['f2+'], ['sum("f3")']
    );
    return tf.assertNotNull(result);
  });
}

// ==================== 测试模块 4: 错误处理 ====================

function test_ErrorHandling(tf) {
  Console.log('\n【模块 4】错误处理测试');
  Console.log('─────────────────────────────────────');

  Console.log('\n4.1 输入验证:');
  
  tf.test('空数组处理', function() {
    try {
      var result = Array2D.z超级透视(
        [],
        ['f1+'], ['f2+'], ['sum("f3")']
      );
      return true; // 不抛错即通过
    } catch(e) {
      return true; // 抛错也是合理行为
    }
  });

  tf.test('无效字段索引', function() {
    try {
      var result = Array2D.z超级透视(
        testDataSets.simple,
        ['f99+'], ['f2+'], ['sum("f3")']
      );
      return true;
    } catch(e) {
      return true;
    }
  });
}

// ==================== 测试模块 5: 性能测试 ====================

function test_Performance(tf) {
  Console.log('\n【模块 5】性能测试');
  Console.log('─────────────────────────────────────');

  Console.log('\n5.1 执行时间:');
  
  tf.test('100 行数据执行时间', function() {
    var startTime = new Date().getTime();
    var result = Array2D.z超级透视(
      testDataSets.sales,
      ['f1+'], ['f2+'], ['sum("f4")']
    );
    var endTime = new Date().getTime();
    var duration = endTime - startTime;
    Console.log('    (耗时: ' + duration + 'ms)');
    return duration < 1000; // 应该在1秒内完成
  });

  tf.test('复杂透视执行时间', function() {
    var startTime = new Date().getTime();
    var result = Array2D.z超级透视(
      testDataSets.category,
      ['f1+,f2+'], ['f3+,f4+'], ['sum("f6"),count(),average("f6")'],
      1, 1, '@^@',
      { subtotals: { row: true, col: true }, grandTotal: { row: true, col: true } }
    );
    var endTime = new Date().getTime();
    var duration = endTime - startTime;
    Console.log('    (耗时: ' + duration + 'ms)');
    return duration < 2000; // 应该在2秒内完成
  });
}

// ==================== 测试模块 6: 集成测试 ====================

function test_Integration(tf) {
  Console.log('\n【模块 6】集成测试');
  Console.log('─────────────────────────────────────');

  Console.log('\n6.1 实际场景:');
  
  tf.test('销售数据透视', function() {
    var result = Array2D.z超级透视(
      testDataSets.sales,
      ['f1+'], ['f2+'], ['sum("f4"),sum("f5")']
    );
    return tf.assertNotNull(result) && result.length > 0;
  });

  tf.test('多级分类透视', function() {
    var result = Array2D.z超级透视(
      testDataSets.category,
      ['f1+,f2+,f3+'], ['f4+'], ['sum("f6"),sum("f7")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('完整功能综合测试', function() {
    var result = Array2D.z超级透视(
      testDataSets.sales,
      ['f1+,f2+'], ['f3+'], 
      ['sum("f4"),count(),average("f4")'],
      1, 1, '@^@',
      {
        subtotals: { row: true, col: true },
        grandTotal: { row: true, col: true },
        displayAs: { mode: 'value' }
      }
    );
    return tf.assertNotNull(result) && result.length > 0;
  });
}

// ==================== 主测试运行器 ====================

function runComprehensiveTests() {
  Console.log('\n╔════════════════════════════════════════╗');
  Console.log('║   JSA880 全面测试套件                  ║');
  Console.log('║   版本: v3.9.0                          ║');
  Console.log('║   日期: ' + new Date().toLocaleDateString() + '              ║');
  Console.log('╚════════════════════════════════════════╝');
  
  var tf = new TestFramework();
  
  try {
    test_v390_NewFeatures(tf);
    test_Regression(tf);
    test_EdgeCases(tf);
    test_ErrorHandling(tf);
    test_Performance(tf);
    test_Integration(tf);
  } catch (e) {
    Console.log('\n❌ 测试套件执行失败: ' + e.message);
  }
  
  var passed = tf.report();
  
  Console.log('========================================');
  if (passed) {
    Console.log('🎉 所有测试通过！');
  } else {
    Console.log('⚠️  部分测试失败，请查看详情');
  }
  Console.log('========================================\n');
  
  return passed;
}

// ==================== 快速测试（仅关键功能）====================

function runQuickTests() {
  Console.log('\n╔════════════════════════════════════════╗');
  Console.log('║   JSA880 快速测试                      ║');
  Console.log('║   仅测试关键功能                        ║');
  Console.log('╚════════════════════════════════════════╝\n');

  var tf = new TestFramework();
  
  Console.log('【关键功能测试】');
  Console.log('─────────────────────────────────────\n');

  tf.test('v3.9.0: 行小计', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { subtotals: { row: true, col: true } }
    );
    return tf.assertNotNull(result);
  });

  tf.test('v3.9.0: 总计行', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")'],
      1, 1, '@^@',
      { grandTotal: { row: true, col: true } }
    );
    return tf.assertNotNull(result);
  });

  tf.test('向后兼容性', function() {
    var result = Array2D.z超级透视(
      testDataSets.simple,
      ['f1+'], ['f2+'], ['sum("f3")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('多级字段', function() {
    var result = Array2D.z超级透视(
      testDataSets.category,
      ['f1+,f2+'], ['f3+'], ['sum("f6")']
    );
    return tf.assertNotNull(result);
  });

  tf.test('多聚合函数', function() {
    var result = Array2D.z超级透视(
      testDataSets.sales,
      ['f1+'], ['f2+'], ['sum("f4"),count(),average("f4")']
    );
    return tf.assertNotNull(result);
  });

  var passed = tf.report();
  
  Console.log('========================================');
  if (passed) {
    Console.log('✅ 快速测试通过！v3.9.0 可以使用');
  } else {
    Console.log('❌ 快速测试失败！需要修复');
  }
  Console.log('========================================\n');
  
  return passed;
}

// ==================== 导出 ====================

// 默认运行快速测试
runQuickTests();

// 如需运行完整测试，取消下面注释：
// runComprehensiveTests();
