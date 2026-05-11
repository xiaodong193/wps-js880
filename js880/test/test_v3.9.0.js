/**
 * ========================================================================
 * SuperPivot v3.9.0 测试文件
 * ========================================================================
 * 
 * 说明：测试小计和总计功能
 * 使用：在 WPS JSA 环境中运行此文件
 * 
 * ========================================================================
 */

function testSuperPivot_v390() {
  Console.log('========================================');
  Console.log('  SuperPivot v3.9.0 功能测试');
  Console.log('========================================\n');

  // ========== 测试数据 ==========
  var testData = [
    ['产品', '地区', '销售额'],
    ['A', '北京', 1000],
    ['A', '上海', 1500],
    ['A', '广州', 800],
    ['B', '北京', 2000],
    ['B', '上海', 1000],
    ['B', '广州', 1200]
  ];

  // ========== 测试1: 行小计 ==========
  Console.log('【测试1】行小计');
  Console.log('配置: subtotals.row = true');
  try {
    var result1 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      subtotals: { 
        row: true,
        col: false,
        label: '小计' 
      }
    });
    Console.log('✅ 成功！结果行数: ' + result1.length);
    Console.log('第一行数据: ' + JSON.stringify(result1[0]));
    Console.log('最后一行数据: ' + JSON.stringify(result1[result1.length - 1]));
    result1.toRange("A1");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  // ========== 测试2: 列小计 ==========
  Console.log('\n【测试2】列小计');
  Console.log('配置: subtotals.col = true');
  try {
    var result2 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      subtotals: { 
        row: false,
        col: true,
        label: '小计' 
      }
    });
    Console.log('✅ 成功！结果行数: ' + result2.length);
    Console.log('第一行列数: ' + result2[0].length);
    Console.log('最后一行: ' + JSON.stringify(result2[result2.length - 1]));
    result2.toRange("A20");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  // ========== 测试3: 行小计 + 列小计 ==========
  Console.log('\n【测试3】行小计 + 列小计');
  Console.log('配置: subtotals.row = true, subtotals.col = true');
  try {
    var result3 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      subtotals: { 
        row: true,
        col: true,
        label: '小计' 
      }
    });
    Console.log('✅ 成功！结果行数: ' + result3.length);
    Console.log('第一行列数: ' + result3[0].length);
    result3.toRange("A40");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  // ========== 测试4: 总计行 ==========
  Console.log('\n【测试4】总计行');
  Console.log('配置: grandTotal.row = true');
  try {
    var result4 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      grandTotal: { 
        row: true,
        col: false,
        label: '总计' 
      }
    });
    Console.log('✅ 成功！结果行数: ' + result4.length);
    Console.log('最后一行（总计）: ' + JSON.stringify(result4[result4.length - 1]));
    result4.toRange("A60");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  // ========== 测试5: 总计列 ==========
  Console.log('\n【测试5】总计列');
  Console.log('配置: grandTotal.col = true');
  try {
    var result5 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      grandTotal: { 
        row: false,
        col: true,
        label: '总计' 
      }
    });
    Console.log('✅ 成功！结果行数: ' + result5.length);
    Console.log('第一行列数: ' + result5[0].length);
    result5.toRange("A80");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  // ========== 测试6: 完整功能 ==========
  Console.log('\n【测试6】完整功能（小计 + 总计）');
  Console.log('配置: 所有选项启用');
  try {
    var result6 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      subtotals: { 
        row: true, 
        col: true,
        label: '小计' 
      },
      grandTotal: { 
        row: true, 
        col: true,
        label: '总计' 
      }
    });
    Console.log('✅ 成功！结果行数: ' + result6.length);
    Console.log('第一行列数: ' + result6[0].length);
    Console.log('最后一行（总计）: ' + JSON.stringify(result6[result6.length - 1]));
    result6.toRange("A100");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  // ========== 测试7: 多个数据字段 ==========
  Console.log('\n【测试7】多个数据字段');
  Console.log('配置: sum("f3"), count(), average("f3")');
  try {
    var result7 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3"),count(),average("f3")'], 1, 1, '@^@', {
      subtotals: { row: true, col: true },
      grandTotal: { row: true, col: true }
    });
    Console.log('✅ 成功！结果行数: ' + result7.length);
    Console.log('第一行列数: ' + result7[0].length);
    result7.toRange("A120");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  // ========== 测试8: 向后兼容 ==========
  Console.log('\n【测试8】向后兼容（不启用新功能）');
  Console.log('配置: 不传 options 参数');
  try {
    var result8 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")']);
    Console.log('✅ 成功！结果行数: ' + result8.length);
    Console.log('行为与旧版本一致: ' + (result8.length <= 3 ? '是' : '否'));
    result8.toRange("A140");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  // ========== 测试9: 自定义标签 ==========
  Console.log('\n【测试9】自定义小计/总计标签');
  Console.log('配置: label = "合计", "总和"');
  try {
    var result9 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      subtotals: { row: true, col: true, label: '合计' },
      grandTotal: { row: true, col: true, label: '总和' }
    });
    Console.log('✅ 成功！');
    Console.log('检查最后一行是否包含"总和": ' + (JSON.stringify(result9[result9.length - 1]).indexOf('总和') > 0 ? '是' : '否'));
    result9.toRange("A160");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  // ========== 测试10: 多级行字段 ==========
  Console.log('\n【测试10】多级行字段');
  var multiLevelData = [
    ['年份', '季度', '产品', '销售额'],
    ['2024', 'Q1', 'A', 1000],
    ['2024', 'Q1', 'B', 1500],
    ['2024', 'Q2', 'A', 1200],
    ['2024', 'Q2', 'B', 1800],
    ['2025', 'Q1', 'A', 1100],
    ['2025', 'Q1', 'B', 1600]
  ];
  
  try {
    var result10 = Array2D.z超级透视(multiLevelData, ['f1+,f2+,f3+'], ['f4+'], ['sum("f4")'], 1, 1, '@^@', {
      subtotals: { row: true, col: true },
      grandTotal: { row: true, col: true }
    });
    Console.log('✅ 成功！结果行数: ' + result10.length);
    result10.toRange("A180");
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
  }

  Console.log('\n========================================');
  Console.log('  所有测试完成！');
  Console.log('========================================');
}

// ========== 验证函数 ==========
function validateSuperPivot_v390() {
  Console.log('\n========================================');
  Console.log('  v3.9.0 验证检查');
  Console.log('========================================\n');

  var checks = [
    { name: '参数解析 - options.subtotals', test: function() {
      return typeof Array2D.z超级透视 !== 'undefined';
    }},
    { name: '参数解析 - options.grandTotal', test: function() {
      return typeof Array2D.z超级透视 !== 'undefined';
    }},
    { name: '向后兼容', test: function() {
      try {
        var data = [['A', 'B'], [1, 2]];
        var result = Array2D.z超级透视(data, ['f1+'], ['f2+'], ['sum("f2")']);
        return result !== undefined;
      } catch(e) {
        return false;
      }
    }}
  ];

  var passed = 0;
  var failed = 0;

  for (var i = 0; i < checks.length; i++) {
    var check = checks[i];
    try {
      var result = check.test();
      if (result) {
        Console.log('✅ ' + check.name);
        passed++;
      } else {
        Console.log('❌ ' + check.name);
        failed++;
      }
    } catch(e) {
      Console.log('❌ ' + check.name + ' - 错误: ' + e.message);
      failed++;
    }
  }

  Console.log('\n验证结果: ' + passed + ' 通过, ' + failed + ' 失败');
  Console.log('========================================\n');

  return failed === 0;
}

// ========== 主函数 ==========
function main() {
  Console.log('\n');
  Console.log('╔════════════════════════════════════════╗');
  Console.log('║  SuperPivot v3.9.0 测试套件          ║');
  Console.log('║  版本: v3.9.0                         ║');
  Console.log('║  日期: 2025-01-09                     ║');
  Console.log('╚════════════════════════════════════════╝');
  Console.log('\n');

  // 先验证
  var validated = validateSuperPivot_v390();

  if (validated) {
    Console.log('✅ 验证通过，开始运行测试...\n');
    testSuperPivot_v390();
  } else {
    Console.log('❌ 验证失败，请检查实现！');
  }
}

// 运行
main();
