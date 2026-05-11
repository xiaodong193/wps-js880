/**
 * v3.9.0 实施后快速验证
 */

function quickTest_v390() {
  Console.log('========================================');
  Console.log('  v3.9.0 实施后快速验证');
  Console.log('========================================\n');

  // 简单测试数据
  var testData = [
    ['产品', '地区', '销售额'],
    ['A', '北京', 1000],
    ['A', '上海', 1500],
    ['B', '北京', 2000],
    ['B', '上海', 1000]
  ];

  Console.log('【测试 1】基础功能（向后兼容）');
  try {
    var result1 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")']);
    Console.log('✅ 通过 - 结果行数: ' + result1.length);
    Console.log('   第一行: ' + JSON.stringify(result1[0]));
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
    return false;
  }

  Console.log('\n【测试 2】列小计');
  try {
    var result2 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      subtotals: { row: false, col: true }
    });
    Console.log('✅ 通过 - 结果行数: ' + result2.length);
    Console.log('   表头列数: ' + result2[0].length);
    Console.log('   第一行: ' + JSON.stringify(result2[0]));
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
    return false;
  }

  Console.log('\n【测试 3】行小计 + 列小计');
  try {
    var result3 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      subtotals: { row: true, col: true, label: '小计' }
    });
    Console.log('✅ 通过 - 结果行数: ' + result3.length);
    Console.log('   表头列数: ' + result3[0].length);
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
    return false;
  }

  Console.log('\n【测试 4】总计行');
  try {
    var result4 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      grandTotal: { row: true, col: false, label: '总计' }
    });
    Console.log('✅ 通过 - 结果行数: ' + result4.length);
    Console.log('   最后一行: ' + JSON.stringify(result4[result4.length - 1]));
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
    return false;
  }

  Console.log('\n【测试 5】完整功能（小计 + 总计）');
  try {
    var result5 = Array2D.z超级透视(testData, ['f1+'], ['f2+'], ['sum("f3")'], 1, 1, '@^@', {
      subtotals: { row: true, col: true },
      grandTotal: { row: true, col: true }
    });
    Console.log('✅ 通过 - 结果行数: ' + result5.length);
    Console.log('   表头列数: ' + result5[0].length);
    Console.log('   最后一行: ' + JSON.stringify(result5[result5.length - 1]));
  } catch(e) {
    Console.log('❌ 失败: ' + e.message);
    return false;
  }

  Console.log('\n========================================');
  Console.log('🎉 所有快速测试通过！');
  Console.log('v3.9.0 实施成功！');
  Console.log('========================================\n');

  return true;
}

// 运行测试
quickTest_v390();
