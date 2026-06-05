/**
 * SuperPivot Field Coordination Debug - JSA880 v3.9.4+
 */

(function() {
    var Console = typeof Console !== 'undefined' ? Console : { log: function() { console.log.apply(console, arguments); } };
    var Array2D;

    try {
        var fs = require('fs');
        var path = require('path');
        var jsaPath = path.join(__dirname, '../JSA880.js');
        if (fs.existsSync(jsaPath)) {
            var code = fs.readFileSync(jsaPath, 'utf8');
            code = code.replace(/isWPS\s*=\s*typeof\s+Application[^;]+;/g, 'isWPS = false;');
            eval(code);
            Console.log('JSA880 loaded');
        }
    } catch (e) {
        Console.log('Load error: ' + e.message);
    }

    var TEST_DATA = [
        ['产品', '地区', '年份', '季度', '销售额', '数量'],
        ['A产品', '北京', '2024', 'Q1', 1000, 10],
        ['A产品', '北京', '2024', 'Q2', 1500, 15],
        ['A产品', '上海', '2024', 'Q1', 800, 8],
        ['A产品', '上海', '2024', 'Q2', 1200, 12],
        ['B产品', '北京', '2024', 'Q1', 900, 9],
        ['B产品', '北京', '2024', 'Q2', 1100, 11],
        ['B产品', '上海', '2024', 'Q1', 700, 7],
        ['B产品', '上海', '2024', 'Q2', 1300, 13],
        ['A产品', '北京', '2025', 'Q1', 1100, 11],
        ['A产品', '北京', '2025', 'Q2', 1600, 16],
        ['A产品', '上海', '2025', 'Q1', 900, 9],
        ['A产品', '上海', '2025', 'Q2', 1400, 14],
        ['B产品', '北京', '2025', 'Q1', 1000, 10],
        ['B产品', '北京', '2025', 'Q2', 1200, 12],
        ['B产品', '上海', '2025', 'Q1', 800, 8],
        ['B产品', '上海', '2025', 'Q2', 1500, 15]
    ];

    Console.log('========================================');
    Console.log('Field Coordination Debug');
    Console.log('========================================');

    if (!Array2D || !Array2D.z超级透视) {
        Console.log('z超级透视 not available');
        return;
    }

    // Test 1: Check how col field expansion works
    Console.log('\n[Test 1] Col field expansion (地区+年份 -> 4 columns)');
    var r1 = Array2D.z超级透视(TEST_DATA, ['f1+'], ['f2+,f3+'], ['sum("f5")'], 1, 0);
    Console.log('Rows: ' + r1.length + ', Cols per row: ' + r1[0].length);
    Console.log('A产品: ' + JSON.stringify(r1[0]));
    Console.log('B产品: ' + JSON.stringify(r1[1]));
    // The expected structure: each row is [产品, 北京2024, 北京2025, 上海2024, 上海2025] = 5 values
    // But with data field, each "cell" has 1 value (sum), so: [产品, 4300(北京2024), ..., 5200(上海2024), 52(上海Q2?)]

    // Test 2: Check 3-level col header
    Console.log('\n[Test 2] 3-level col header (地区+年份+季度)');
    var r2 = Array2D.z超级透视(TEST_DATA, ['f1+'], ['f2+,f3+,f4+'], ['sum("f5")'], 1, 1);
    Console.log('Total rows: ' + r2.length);
    Console.log('Header row 0: ' + JSON.stringify(r2[0]));
    Console.log('Header row 1: ' + JSON.stringify(r2[1]));
    Console.log('Header row 2: ' + JSON.stringify(r2[2]));
    Console.log('Header row 3: ' + JSON.stringify(r2[3]));
    Console.log('Data row 1: ' + JSON.stringify(r2[4]));
    Console.log('Data row 2: ' + JSON.stringify(r2[5]));

    // Test 3: Check single level col
    Console.log('\n[Test 3] Single level col (地区 only)');
    var r3 = Array2D.z超级透视(TEST_DATA, ['f1+'], ['f2+'], ['sum("f5"),sum("f6")'], 1, 0);
    Console.log('Rows: ' + r3.length);
    Console.log('A产品: ' + JSON.stringify(r3[0]));
    Console.log('B产品: ' + JSON.stringify(r3[1]));
    // Expected: [产品, 北京求和, 上海求和, 北京求和, 上海求和] = 5 cols
    // Or [产品, 北京销售额, 北京数量, 上海销售额, 上海数量] = 5 cols

    // Test 4: Check with no row field
    Console.log('\n[Test 4] No row field');
    var r4 = Array2D.z超级透视(TEST_DATA, [], ['f2+,f3+'], ['sum("f5")'], 1, 0);
    Console.log('Rows: ' + r4.length);
    Console.log('First row: ' + JSON.stringify(r4[0]));

    // Test 5: Check cornerTitle
    Console.log('\n[Test 5] Corner title');
    var r5 = Array2D.z超级透视(TEST_DATA, ['f1+'], ['f2+,f3+'], ['sum("f5")'], 1, 1, '@^@', {
        cornerTitle: '产品'
    });
    Console.log('Header row 0: ' + JSON.stringify(r5[0]));

    // Test 6: Check outputHeader=-1 (no row titles)
    Console.log('\n[Test 6] outputHeader=-1 (hide row titles)');
    var r6 = Array2D.z超级透视(TEST_DATA, ['f1+'], ['f2+'], ['sum("f5")'], 1, -1);
    Console.log('Header row 0: ' + JSON.stringify(r6[0]));
    Console.log('Data row 1: ' + JSON.stringify(r6[2]));

    // Test 7: Check filter with 3-level
    Console.log('\n[Test 7] Filter on 3-level cols');
    var r7 = Array2D.z超级透视(TEST_DATA, ['f1+'], ['f2+,f3+,f4+'], ['sum("f5")'], 1, 0, '@^@', {
        filterCols: { f1: ['北京'], f2: ['2024'] }
    });
    Console.log('Filtered rows: ' + r7.length);
    if (r7.length > 0) {
        Console.log('A产品 filtered: ' + JSON.stringify(r7[0]));
    }

    // Test 8: Check multiple data ops
    Console.log('\n[Test 8] Multiple data ops (count,sum,avg)');
    var r8 = Array2D.z超级透视(TEST_DATA, ['f1+'], ['f2+'], ['count(),sum("f5"),average("f6")'], 1, 0);
    Console.log('A产品: ' + JSON.stringify(r8[0]));
    Console.log('B产品: ' + JSON.stringify(r8[1]));
    // Expected: [产品, 北京count, 北京sum, 北京avg, 上海count, 上海sum, 上海avg] = 7 cols

    Console.log('\n========================================');
    Console.log('Debug complete');
    Console.log('========================================');

})();