var clsPerformanceTester = (function() {
    function clsPerformanceTester() {
        this.MODULE_NAME = "clsPerformanceTester";
        this.testResults = [];
    }
    
    clsPerformanceTester.prototype.runTimedTest = function(testName, testFunc, iterations) {
        iterations = iterations || 1;
        
        var startTime = this._getTime();
        var result;
        
        for (var i = 0; i < iterations; i++) {
            result = testFunc();
        }
        
        var endTime = this._getTime();
        var duration = endTime - startTime;
        var avgDuration = duration / iterations;
        
        return {
            testName: testName,
            iterations: iterations,
            totalDuration: duration,
            avgDuration: avgDuration,
            result: result,
            success: true
        };
    };
    
    clsPerformanceTester.prototype._getTime = function() {
        if (typeof Date.now === "function") {
            return Date.now();
        }
        return new Date().getTime();
    };
    
    clsPerformanceTester.prototype.runAllTests = function() {
        console.log("========================================");
        console.log("   Array2D 性能测试开始");
        console.log("========================================");
        
        this.testResults = [];
        
        this.testArrToArrData();
        this.testGenerateRateArray();
        this.testBatchWriteRates();
        this.testFilterAndTransform();
        this.testAggregateByGroup();
        
        this.printSummary();
        
        return this.testResults;
    };
    
    clsPerformanceTester.prototype.testArrToArrData = function() {
        console.log("\n--- 测试 arrToArrData ---");
        
        var arrFormula = this._createFormulaTemplate(36, 13);
        
        var optimizer = new clsArray2DOptimizer(null);
        
        var resultSmall = this.runTimedTest("arrToArrData-10期", function() {
            return optimizer.arrToArrDataOptimized(arrFormula, 10, { FIRST: 1, MIDDLE: 2, LAST: 3 });
        }, 100);
        this.testResults.push(resultSmall);
        console.log("  10期 (100次迭代): 平均 " + resultSmall.avgDuration.toFixed(3) + "ms");
        
        var resultMedium = this.runTimedTest("arrToArrData-36期", function() {
            return optimizer.arrToArrDataOptimized(arrFormula, 36, { FIRST: 1, MIDDLE: 2, LAST: 3 });
        }, 100);
        this.testResults.push(resultMedium);
        console.log("  36期 (100次迭代): 平均 " + resultMedium.avgDuration.toFixed(3) + "ms");
        
        var resultLarge = this.runTimedTest("arrToArrData-120期", function() {
            return optimizer.arrToArrDataOptimized(arrFormula, 120, { FIRST: 1, MIDDLE: 2, LAST: 3 });
        }, 50);
        this.testResults.push(resultLarge);
        console.log("  120期 (50次迭代): 平均 " + resultLarge.avgDuration.toFixed(3) + "ms");
        
        var resultOriginal = this.runTimedTest("arrToArrData-原实现-36期", function() {
            return optimizer._arrToArrDataOriginal(arrFormula, 36);
        }, 100);
        console.log("  原实现 36期 (100次迭代): 平均 " + resultOriginal.avgDuration.toFixed(3) + "ms");
        
        var improvement = ((resultOriginal.avgDuration - resultMedium.avgDuration) / resultOriginal.avgDuration * 100).toFixed(1);
        console.log("  性能提升: " + improvement + "%");
    };
    
    clsPerformanceTester.prototype.testGenerateRateArray = function() {
        console.log("\n--- 测试 generateRateArray ---");
        
        var optimizer = new clsArray2DOptimizer(null);
        var baseRate = 0.035;
        
        var resultSmall = this.runTimedTest("generateRateArray-10期", function() {
            return optimizer.generateRateArrayOptimized(10, baseRate);
        }, 1000);
        this.testResults.push(resultSmall);
        console.log("  10期 (1000次迭代): 平均 " + resultSmall.avgDuration.toFixed(4) + "ms");
        
        var resultLarge = this.runTimedTest("generateRateArray-360期", function() {
            return optimizer.generateRateArrayOptimized(360, baseRate);
        }, 100);
        this.testResults.push(resultLarge);
        console.log("  360期 (100次迭代): 平均 " + resultLarge.avgDuration.toFixed(4) + "ms");
        
        var resultOriginal = this.runTimedTest("generateRateArray-原实现-360期", function() {
            var arr2D = [];
            for (var i = 0; i < 360; i++) {
                arr2D.push([baseRate]);
            }
            return arr2D;
        }, 100);
        console.log("  原实现 360期 (100次迭代): 平均 " + resultOriginal.avgDuration.toFixed(4) + "ms");
        
        var improvement = ((resultOriginal.avgDuration - resultLarge.avgDuration) / resultOriginal.avgDuration * 100).toFixed(1);
        console.log("  性能提升: " + improvement + "%");
    };
    
    clsPerformanceTester.prototype.testBatchWriteRates = function() {
        console.log("\n--- 测试 batchWriteRates (模拟) ---");
        
        var optimizer = new clsArray2DOptimizer(null);
        var rates = optimizer.generateRateArrayOptimized(36, 0.035);
        
        var result = this.runTimedTest("batchWriteRates-模拟批量写入", function() {
            var dummy = rates[0][0];
            return rates;
        }, 100);
        
        this.testResults.push(result);
        console.log("  36期 (100次迭代): 平均 " + result.avgDuration.toFixed(4) + "ms");
        console.log("  注: 实际批量写入性能取决于 WPS API 响应");
    };
    
    clsPerformanceTester.prototype.testFilterAndTransform = function() {
        console.log("\n--- 测试 filterAndTransform ---");
        
        var testData = [];
        for (var i = 0; i < 1000; i++) {
            testData.push([
                i + 1,
                "类别" + (i % 5),
                Math.random() * 10000,
                Math.random() * 1000,
                Math.random() * 500,
                Math.random() * 200,
                Math.random() * 100,
                Math.random() * 50,
                Math.random() * 20,
                Math.random() * 10
            ]);
        }
        
        var optimizer = new clsArray2DOptimizer(null);
        
        var resultFilter = this.runTimedTest("filter-1000行", function() {
            return testData.filter(function(row) { return row[2] > 5000; });
        }, 50);
        console.log("  筛选 1000行 (50次迭代): 平均 " + resultFilter.avgDuration.toFixed(4) + "ms");
        
        var resultTransform = this.runTimedTest("filter+transform-1000行", function() {
            return testData.filter(function(row) { return row[2] > 5000; })
                           .map(function(row) { return [row[0], row[1], row[2] * 1.1]; });
        }, 50);
        console.log("  筛选+转换 1000行 (50次迭代): 平均 " + resultTransform.avgDuration.toFixed(4) + "ms");
        
        this.testResults.push(resultFilter);
        this.testResults.push(resultTransform);
    };
    
    clsPerformanceTester.prototype.testAggregateByGroup = function() {
        console.log("\n--- 测试 aggregateByGroup ---");
        
        var testData = [];
        for (var i = 0; i < 500; i++) {
            testData.push([
                i + 1,
                "产品" + (i % 20),
                "地区" + (i % 5),
                Math.random() * 10000,
                1
            ]);
        }
        
        var optimizer = new clsArray2DOptimizer(null);
        
        if (typeof Array2D !== "undefined" && Array2D.groupInto) {
            var resultArray2D = this.runTimedTest("aggregateByGroup-Array2D-500行", function() {
                return Array2D.groupInto(testData, "f2,f3", 'count(),sum("f4")');
            }, 50);
            console.log("  Array2D分组 500行 (50次迭代): 平均 " + resultArray2D.avgDuration.toFixed(4) + "ms");
            this.testResults.push(resultArray2D);
        }
        
        var resultFallback = this.runTimedTest("aggregateByGroup-降级-500行", function() {
            return optimizer.aggregateByGroup(testData, "f2", "count(),sum('f4')");
        }, 50);
        console.log("  降级分组 500行 (50次迭代): 平均 " + resultFallback.avgDuration.toFixed(4) + "ms");
        this.testResults.push(resultFallback);
    };
    
    clsPerformanceTester.prototype.printSummary = function() {
        console.log("\n========================================");
        console.log("   性能测试总结");
        console.log("========================================");
        
        var totalTests = this.testResults.length;
        var successCount = 0;
        var totalDuration = 0;
        
        for (var i = 0; i < this.testResults.length; i++) {
            var result = this.testResults[i];
            if (result.success) {
                successCount++;
                totalDuration = totalDuration + result.avgDuration;
            }
        }
        
        console.log("\n总测试数: " + totalTests);
        console.log("成功数: " + successCount);
        console.log("失败数: " + (totalTests - successCount));
        console.log("平均执行时间: " + (totalDuration / totalTests).toFixed(4) + "ms");
        
        console.log("\n各测试详情:");
        for (var j = 0; j < this.testResults.length; j++) {
            var r = this.testResults[j];
            var status = r.success ? "OK" : "FAIL";
            console.log("  " + status + " " + r.testName + ": " + r.avgDuration.toFixed(4) + "ms");
        }
    };
    
    clsPerformanceTester.prototype.generateReport = function() {
        var totalDur = 0;
        for (var i = 0; i < this.testResults.length; i++) {
            totalDur = totalDur + this.testResults[i].totalDuration;
        }
        var avgDur = 0;
        if (this.testResults.length > 0) {
            avgDur = totalDur / this.testResults.length;
        }
        
        var report = {
            timestamp: new Date().toLocaleString("zh-CN"),
            tests: this.testResults,
            summary: {
                totalTests: this.testResults.length,
                totalDuration: totalDur,
                avgDuration: avgDur
            }
        };
        return report;
    };
    
    clsPerformanceTester.prototype._createFormulaTemplate = function(rows, cols) {
        var arr = [];
        for (var i = 0; i <= 3; i++) {
            arr[i] = [];
            for (var j = 0; j <= cols; j++) {
                arr[i][j] = "=FORMULA" + i + "_" + j;
            }
        }
        return arr;
    };
    
    return clsPerformanceTester;
})();

if (typeof clsArray2DOptimizer !== "undefined") {
    clsArray2DOptimizer.prototype._arrToArrDataOriginal = function(arrFormula, length) {
        var FORMULA_ROW = { FIRST: 1, MIDDLE: 2, LAST: 3 };
        var maxCol = arrFormula[FORMULA_ROW.FIRST].length - 1;
        
        var arrData = [];
        for (var i = 0; i < length; i++) {
            arrData[i] = new Array(arrFormula[FORMULA_ROW.FIRST].length);
        }
        
        for (var row = 0; row < length; row++) {
            var rowIndex;
            if (length === 1) {
                rowIndex = FORMULA_ROW.FIRST;
            } else if (row === 0) {
                rowIndex = FORMULA_ROW.FIRST;
            } else if (row === length - 1) {
                rowIndex = FORMULA_ROW.LAST;
            } else {
                rowIndex = FORMULA_ROW.MIDDLE;
            }
            for (var col = 1; col <= maxCol && col < arrFormula[FORMULA_ROW.FIRST].length; col++) {
                arrData[row][col - 1] = arrFormula[rowIndex][col];
            }
        }
        
        return arrData;
    };
}

function runPerformanceTest(iterations) {
    var tester = new clsPerformanceTester();
    tester.runAllTests();
    return tester.generateReport();
}

function runSingleTest(testName, testFunc) {
    var tester = new clsPerformanceTester();
    return tester.runTimedTest(testName, testFunc, 1);
}

function printPerformanceReport(report) {
    console.log("\n========================================");
    console.log("   Array2D 性能对比报告");
    console.log("========================================");
    console.log("生成时间: " + report.timestamp);
    console.log("总测试数: " + report.summary.totalTests);
    console.log("总耗时: " + report.summary.totalDuration.toFixed(2) + "ms");
    console.log("平均耗时: " + report.summary.avgDuration.toFixed(4) + "ms");
    
    console.log("\n各测试详情:");
    for (var i = 0; i < report.tests.length; i++) {
        var t = report.tests[i];
        console.log("  " + t.testName + ": " + t.avgDuration.toFixed(4) + "ms (" + t.iterations + "次迭代)");
    }
}

function benchmark(name, func, iterations) {
    iterations = iterations || 100;
    
    var start = Date.now();
    for (var i = 0; i < iterations; i++) {
        func();
    }
    var end = Date.now();
    
    return {
        name: name,
        iterations: iterations,
        totalDuration: end - start,
        avgDuration: (end - start) / iterations
    };
}

console.log("[mArray2DOptimizer_test.js] Array2D性能测试套件加载完成");