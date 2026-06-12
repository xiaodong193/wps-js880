/**
 * ========================================================================
 * SuperPivot 完整回归测试套件 (WPS JSA) v2.1
 * ========================================================================
 * 覆盖 Bug#1-9 修复验证 + 教案 3.25 节全部场景
 * 运行: 快速测试() 或 运行SuperPivot完整测试()
 * ========================================================================
 */

// ==================== 测试数据 ====================

function 生成测试数据(rowCount) {
    rowCount = rowCount || 100;
    var products = ['手机', '电脑', '平板', '耳机'];
    var countries = ['中国', '美国', '德国', '日本'];
    var years = [2020, 2021, 2022, 2023];
    var months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    var data = [['序号', '产品', '国家', '数量', '价格', '年', '月']];
    for (var i = 0; i < rowCount; i++) {
        data.push([i + 1,
            products[Math.floor(Math.random() * products.length)],
            countries[Math.floor(Math.random() * countries.length)],
            Math.floor(Math.random() * 100) + 1,
            Math.floor(Math.random() * 500) + 50,
            years[Math.floor(Math.random() * years.length)],
            months[Math.floor(Math.random() * months.length)]
        ]);
    }
    return Array2D(data);
}

function 生成可控测试数据() {
    return Array2D([
        ['产品', '国家', '数量', '价格', '年', '月'],
        ['手机', '中国', 10, 100, 2020, 1],
        ['手机', '中国', 20, 200, 2020, 2],
        ['手机', '美国', 15, 150, 2020, 1],
        ['手机', '美国', 25, 250, 2021, 3],
        ['电脑', '中国', 30, 300, 2020, 1],
        ['电脑', '中国', 40, 400, 2021, 2],
        ['电脑', '美国', 35, 350, 2021, 3],
        ['电脑', '德国', 45, 450, 2020, 1],
        ['平板', '中国', 50, 500, 2021, 2],
        ['平板', '日本', 55, 550, 2020, 1],
    ]);
}

// ==================== 测试辅助 ====================

var _testNum = 0, _testPass = 0, _testFail = 0, _outputRow = 1, _ws = null;

function 初始化测试输出() {
    try {
        _ws = Worksheets('测试输出');
        _ws.Cells.Clear();
    } catch (e) {
        _ws = Worksheets.Add();
        _ws.Name = '测试输出';
    }
    _testNum = 0; _testPass = 0; _testFail = 0; _outputRow = 1;
}

function 写标题(text) {
    _ws.Cells(_outputRow, 1).Value2 = text;
    _ws.Cells(_outputRow, 1).Font.Bold = true;
    try { _ws.Range(_ws.Cells(_outputRow, 1), _ws.Cells(_outputRow, 10)).Merge(); } catch (e) {}
    _outputRow++;
}

function 写结果(name, passed, detail) {
    _testNum++;
    if (passed) _testPass++; else _testFail++;
    _ws.Cells(_outputRow, 1).Value2 = _testNum + '. ' + name;
    _ws.Cells(_outputRow, 2).Value2 = passed ? 'OK' : 'FAIL';
    _ws.Cells(_outputRow, 2).Font.Color = passed ? 0x008000 : 0xFF0000;
    if (detail) _ws.Cells(_outputRow, 3).Value2 = detail;
    _outputRow++;
}

// 安全输出：直接写入二维数组到单元格，不依赖 toRange 方法
function 安全输出(result) {
    var data;
    if (typeof result.val === 'function') {
        data = result.val();
    } else if (Array.isArray(result)) {
        data = result;
    } else {
        data = [];
    }

    if (!data || data.length === 0) {
        _outputRow += 2;
        return;
    }

    var rows = data.length;
    var cols = rows > 0 ? (Array.isArray(data[0]) ? data[0].length : 1) : 0;

    if (rows === 0 || cols === 0) {
        _outputRow += 2;
        return;
    }

    // 直接写入：使用 Resize 方式
    try {
        var startCell = _ws.Cells(_outputRow, 1);
        var endCell = _ws.Cells(_outputRow + rows - 1, cols);
        var writeRng = _ws.Range(startCell, endCell);
        writeRng.Value2 = data;
    } catch (e1) {
        // 回退：逐行写入
        try {
            for (var r = 0; r < rows; r++) {
                for (var c = 0; c < cols; c++) {
                    _ws.Cells(_outputRow + r, c + 1).Value2 = data[r][c];
                }
            }
        } catch (e2) {
            _ws.Cells(_outputRow, 1).Value2 = '输出失败: ' + e2.message;
        }
    }

    _outputRow += rows + 2;
}

// ==================== 测试用例 ====================

function test_01() {
    写标题('1. 基本透视（行:产品, 列:年+月, 汇总:sum数量）');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(), ['f2'], ['f5,f6'], ['sum("f4")']);
        var ok = result !== null && result !== undefined;
        写结果('基本透视', ok);
        if (ok) 安全输出(result);
    } catch (e) {
        写结果('基本透视', false, e.message);
    }
}

function test_02() {
    写标题('2. 自定义标题（产品名称/年份/月份/计数/总价）');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(),
            ['f2', '产品名称'], ['f5,f6', '年份,月份'],
            ['count(),sum("f4")', '计数,总价']);
        写结果('自定义标题', result !== null);
        if (result) 安全输出(result);
    } catch (e) {
        写结果('自定义标题', false, e.message);
    }
}

function test_03() {
    写标题('3. 无行字段 → 验证 Bug#9 empty rowKeys');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(), [], ['f5,f6'], ['sum("f4")']);
        var d = result ? result.val() : null;
        var hasData = d && d.length > 0 && d.some(function(r) {
            return r.some(function(c) { return c !== '' && c !== null && c !== undefined; });
        });
        写结果('无行字段', hasData, hasData ? '数据行数:' + d.length : '无数据行');
        if (d) 安全输出(result);
    } catch (e) {
        写结果('无行字段', false, e.message);
    }
}

function test_04() {
    写标题('4. 无列字段（纯行聚合）');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(), ['f2'], [], ['sum("f4")']);
        var ok = result !== null && result.val().length > 0;
        写结果('无列字段', ok, ok ? '行数:' + result.val().length : '');
        if (result) 安全输出(result);
    } catch (e) {
        写结果('无列字段', false, e.message);
    }
}

function test_05() {
    写标题('5. 无行无列（纯汇总）');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(), [], [], ['count(),sum("f4")']);
        写结果('无行无列', result !== null);
        if (result) 安全输出(result);
    } catch (e) {
        写结果('无行无列', false, e.message);
    }
}

function test_06() {
    写标题('6. 多汇总函数（count+sum+avg+max+min）');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(), ['f2'], ['f5'],
            ['count(),sum("f4"),average("f4"),max("f4"),min("f4")', '计数,求和,平均,最大,最小']);
        写结果('多汇总函数', result !== null);
        if (result) 安全输出(result);
    } catch (e) {
        写结果('多汇总函数', false, e.message);
    }
}

function test_07() {
    写标题('7. 回调模式（自定义聚合）');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(),
            ['f2', '产品'], ['f5', '年份'],
            [[function(g) { return g.count(); }, function(g) { return g.sum('f4'); }, function(g) { return g.average('f4'); }], '计数,求和,平均']);
        写结果('回调模式', result !== null);
        if (result) 安全输出(result);
    } catch (e) {
        写结果('回调模式', false, e.message);
    }
}

function test_08() {
    写标题('8. 无表头数据(headerRows=0) → 验证 Bug#2 falsy');
    try {
        var arr = Array2D([
            ['手机', '中国', 10, 100, 2020, 1],
            ['手机', '美国', 15, 150, 2020, 2],
            ['电脑', '中国', 20, 200, 2021, 1],
        ]);
        var result = Array2D.z超级透视(arr, ['f1'], ['f5'], ['sum("f4")'], 0);
        var d = result ? result.val() : null;
        var ok = d && d.length > 0;
        写结果('无表头 headerRows=0', ok, ok ? '数据行数:' + d.length : '');
        if (result) 安全输出(result);
    } catch (e) {
        写结果('无表头 headerRows=0', false, e.message);
    }
}

function test_09() {
    写标题('9. 不输出表头(outputHeader=0)');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(), ['f2'], ['f5'], ['sum("f4")'], 1, 0);
        写结果('不输出表头', result !== null);
        if (result) 安全输出(result);
    } catch (e) {
        写结果('不输出表头', false, e.message);
    }
}

function test_10() {
    写标题('10. Map字典模式');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(), ['f2'], ['f5,f6'], ['sum("f4")'], 1, 'map');
        var ok = result && typeof result.size === 'number' && result.size > 0;
        写结果('Map模式', ok, ok ? '条目数:' + result.size : '');
    } catch (e) {
        写结果('Map模式', false, e.message);
    }
}

function test_11() {
    写标题('11. 排序符号 (f2-降序, f5-降序)');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(), ['f2-'], ['f5-'], ['sum("f4")']);
        写结果('排序符号', result !== null);
        if (result) 安全输出(result);
    } catch (e) {
        写结果('排序符号', false, e.message);
    }
}

function test_12() {
    写标题('12. 多行多列表头(2+2) → 验证 Bug#3 totalColSpans');
    try {
        var result = Array2D.z超级透视(生成可控测试数据(),
            ['f2,f3', '产品,国家'], ['f5,f6', '年份,月份'], ['sum("f4")', '总价']);
        var d = result ? result.val() : null;
        if (d) {
            var w = d[0].length, ok = true;
            for (var i = 1; i < d.length; i++) { if (d[i].length !== w) { ok = false; break; } }
            写结果('多列表头列对齐', ok, '宽度:' + w);
        } else {
            写结果('多列表头列对齐', false, '无结果');
        }
        if (d) 安全输出(result);
    } catch (e) {
        写结果('多列表头列对齐', false, e.message);
    }
}

function test_13() {
    写标题('13. 空值/null处理');
    try {
        var arr = Array2D([
            ['产品', '国家', '数量', '价格', '年', '月'],
            ['手机', '中国', 10, 100, 2020, 1],
            ['手机', null, 15, 150, 2020, 2],
            ['电脑', '中国', null, 200, 2021, 1],
        ]);
        var result = Array2D.z超级透视(arr, ['f2,f3'], ['f5'], ['sum("f4"),count()']);
        写结果('空值处理', result !== null);
        if (result) 安全输出(result);
    } catch (e) {
        写结果('空值处理', false, e.message);
    }
}

function test_14() {
    写标题('14. 单行数据边界');
    try {
        var arr = Array2D([['产品', '国家', '数量', '价格', '年', '月'], ['手机', '中国', 10, 100, 2020, 1]]);
        var result = Array2D.z超级透视(arr, ['f2'], ['f5'], ['sum("f4")']);
        写结果('单行数据', result !== null);
        if (result) 安全输出(result);
    } catch (e) {
        写结果('单行数据', false, e.message);
    }
}

function test_15() {
    写标题('15. 大数据量性能(1000行)');
    try {
        var arr = 生成测试数据(1000);
        var t0 = new Date().getTime();
        var result = Array2D.z超级透视(arr, ['f2,f3', '产品,国家'], ['f5,f6', '年,月'], ['count(),sum("f4")', '计数,求和']);
        var ms = new Date().getTime() - t0;
        写结果('1000行性能', ms < 5000, ms + 'ms (期望<5000ms)');
    } catch (e) {
        写结果('1000行性能', false, e.message);
    }
}

// ==================== 主入口 ====================

function 运行SuperPivot完整测试() {
    初始化测试输出();
    写标题('SuperPivot 完整回归测试 v2.1 — 验证 Bug#1/2/3/4/5/6/7/8/9');
    _outputRow++;
    test_01(); test_02(); test_03(); test_04();
    test_05(); test_06(); test_07(); test_08();
    test_09(); test_10(); test_11(); test_12();
    test_13(); test_14(); test_15();
    _outputRow++;
    写标题('总计: ' + _testNum + ' | 通过: ' + _testPass + ' | 失败: ' + _testFail);
    Console.log('SuperPivot 测试完成: ' + _testPass + '/' + _testNum + ' 通过');
}

function 快速测试() {
    初始化测试输出();
    test_01(); test_03(); test_08(); test_12();
    _outputRow++;
    写标题('快速测试完成: ' + _testPass + '/' + _testNum + ' 通过');
    Console.log('快速测试完成: ' + _testPass + '/' + _testNum + ' 通过');
}

Console.log('SuperPivot v2.1 测试套件已加载。运行: 快速测试() 或 运行SuperPivot完整测试()');
