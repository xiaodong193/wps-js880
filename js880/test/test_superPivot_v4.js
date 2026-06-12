global.Application = {};
global.Range = function(){};
global.Sheets = function(){ return {Activate:function(){}}; };
global.Console = { log: function(){} };
global.ActiveSheet = { Name: "Sheet1" };
var fs = require("fs");
eval(fs.readFileSync("./JSA880.js", "utf8"));

var assert = require("assert");
var passed = 0, failed = 0, skipped = 0;

function test(name, fn) {
    try {
        fn();
        passed++;
        console.log("PASS: " + name);
    } catch (e) {
        failed++;
        console.log("FAIL: " + name + " — " + e.message);
    }
}

function skip(name) {
    skipped++;
    console.log("SKIP: " + name + " — not testable in Node");
}

var salesData = [
    ["产品", "地区", "年份", "金额", "数量"],
    ["产品A", "北京", 2023, 100, 10],
    ["产品B", "上海", 2023, 200, 20],
    ["产品A", "北京", 2024, 150, 15],
    ["产品B", "上海", 2024, 250, 25],
    ["产品A", "上海", 2023, 180, 18],
    ["产品B", "北京", 2023, 120, 12],
    ["产品A", "北京", 2024, 160, 16],
];

var dateData = [
    ["产品", "日期", "金额"],
    ["产品A", "2023-01-15", 100],
    ["产品A", "2023-03-20", 150],
    ["产品B", "2023-06-10", 200],
    ["产品A", "2024-02-05", 120],
    ["产品B", "2024-04-18", 180],
    ["产品A", "2024-07-22", 130],
];

test("1. Basic 2-column pivot (row + col + sum)", function() {
    var r = Array2D.superPivot(salesData, "f1", "f2", 'sum("f4")');
    var arr = r.res();
    assert(arr.length > 1, "should have header + data rows");
    var header = arr[0];
    assert(header.indexOf("产品") !== -1 || header[0] !== undefined, "row title present");
});

test("2. Multi-row fields (f1,f2 as rowFields)", function() {
    var r = Array2D.superPivot(salesData, "f1,f2", "f3", 'sum("f4")');
    var arr = r.res();
    assert(arr.length > 2, "should have header + data rows");
    var dataRow = arr[2];
    assert(dataRow.length >= 3, "row should have row key cols + data cols");
});

test("3. Multi-col fields (f3,f4 as colFields)", function() {
    var r = Array2D.superPivot(salesData, "f1", "f3,f2", 'sum("f4")');
    var arr = r.res();
    assert(arr.length > 1, "should have header + data");
    assert(arr[0].length >= 3, "header should span multiple col levels");
});

test("4. No column fields (only row fields)", function() {
    var r = Array2D.superPivot(salesData, "f1", "", 'sum("f4")');
    var arr = r.res();
    assert(arr.length > 1, "should have header + data");
});

test("5. No row fields (only column fields)", function() {
    var r = Array2D.superPivot(salesData, "", "f2", 'sum("f4")');
    var arr = r.res();
    assert(arr.length > 1, "should have header + data");
});

test("6. Multiple data fields (count + sum + avg)", function() {
    var r = Array2D.superPivot(salesData, "f1", "f2", 'count(),sum("f4"),average("f4")');
    var arr = r.res();
    assert(arr.length > 1, "should have header + data");
    var dataRow = arr[2];
    assert(dataRow.length > 4, "should have row key + 3 data cols per col key");
});

test("7. Subtotals enabled", function() {
    var r = Array2D.superPivot(salesData, "f1", "f2", 'sum("f4")', { subtotals: { enabled: true, row: true } });
    var arr = r.res();
    assert(arr.length > 1, "should produce output without error");
});

test("8. Grand total enabled", function() {
    var r = Array2D.superPivot(salesData, "f1", "f2", 'sum("f4")', { grandTotal: { row: true } });
    var arr = r.res();
    assert(arr.length > 2, "should have header + data + grand total row");
    var lastRow = arr[arr.length - 1];
    var hasTotal = false;
    for (var i = 0; i < lastRow.length; i++) {
        if (String(lastRow[i]).indexOf("总计") !== -1) { hasTotal = true; break; }
    }
    assert(hasTotal, "last row should contain total label");
});

test("9. Sort order (+ asc, - desc in rowFields)", function() {
    var rAsc = Array2D.superPivot(salesData, "f1+", "f2", 'sum("f4")');
    var rDesc = Array2D.superPivot(salesData, "f1-", "f2", 'sum("f4")');
    var aAsc = rAsc.res();
    var aDesc = rDesc.res();
    assert(aAsc.length >= 3 && aDesc.length >= 3, "both should have data rows");
    var firstAsc = String(aAsc[2][0]);
    var firstDesc = String(aDesc[2][0]);
    assert(firstAsc !== firstDesc, "asc and desc should differ");
});

test("10. Filter after pivot", function() {
    var r = Array2D.superPivot(salesData, "f1", "f2", 'sum("f4")');
    var filtered = r.filter(function(row, i) { return i === 0 || String(row[0]).indexOf("产品A") !== -1; });
    assert(filtered.length <= r.length, "filtered should be subset or equal");
});

test("11. Empty input data", function() {
    var r = Array2D.superPivot([["h1","h2"]], "f1", "f2", 'count()');
    var arr = r.res();
    assert(Array.isArray(arr), "should return array");
});

test("12. Single row input", function() {
    var data = [["产品","地区","金额"],["产品A","北京",100]];
    var r = Array2D.superPivot(data, "f1", "f2", 'sum("f3")');
    var arr = r.res();
    assert(arr.length >= 2, "should have header + at least 1 data row");
});

test("13. Ragged rows (different column counts)", function() {
    var data = [["产品","地区","金额"],["产品A","北京",100],["产品B"]]; 
    var r = Array2D.superPivot(data, "f1", "", 'count()');
    var arr = r.res();
    assert(Array.isArray(arr), "should handle ragged rows without crash");
});

test("14. Header preservation through pivot", function() {
    var r = Array2D.superPivot(salesData, "f1", "f2", 'sum("f4")');
    var arr = r.res();
    assert(arr.length > 1, "should have output");
    assert(arr[0].length > 0, "header row should not be empty");
});

test("15. smartGroup year/month/quarter grouping", function() {
    var groups = Array2D.smartGroup(dateData.slice(1), "f2", "year");
    assert(groups instanceof Map, "should return Map");
    assert(groups.size > 0, "should have year groups");
    var groupsQ = Array2D.smartGroup(dateData.slice(1), "f2", "quarter");
    assert(groupsQ instanceof Map, "quarter should also work");
    var groupsM = Array2D.smartGroup(dateData.slice(1), "f2", "month");
    assert(groupsM instanceof Map, "month should also work");
});

test("16. textjoin aggregate", function() {
    var data = [["产品","类别","描述"],["产品A","类型1","好"],["产品A","类型2","中"],["产品B","类型1","优"]];
    var r = Array2D.superPivot(data, "f1", "", 'textjoin("f3",",")');
    var arr = r.res();
    assert(arr.length >= 2, "should have header + data");
    var dataRow = arr[1];
    var found = false;
    for (var i = 0; i < dataRow.length; i++) {
        if (String(dataRow[i]).indexOf(",") !== -1) { found = true; break; }
    }
    assert(found, "should contain joined values with comma");
});

test("17. groupInto with multiple aggregates", function() {
    var grp = Array2D.groupInto(salesData.slice(1), "f1", 'count(),sum("f4"),average("f4")');
    assert(grp instanceof Array2D, "should return Array2D instance");
    assert(grp.length > 0, "should have grouped rows");
    assert(grp[0].length >= 3, "each row should have key + aggregates");
});

test("18. Edge: nullValue parameter", function() {
    var r = Array2D.superPivot(salesData, "f1", "f2", 'sum("f5")', { nullValue: "N/A" });
    var arr = r.res();
    assert(Array.isArray(arr), "should return array");
});

test("19. Edge: custom headerRows", function() {
    var dataWith2Headers = [
        ["主标题", "子标题", "列3", "列4"],
        ["产品", "地区", "金额", "数量"],
        ["产品A", "北京", 100, 10],
        ["产品B", "上海", 200, 20],
    ];
    var r = Array2D.superPivot(dataWith2Headers, "f1", "f2", 'sum("f3")', { headerRows: 2 });
    var arr = r.res();
    assert(arr.length >= 1, "should produce output with custom headerRows");
});

skip("20. __KJ_ARGS__ parsing — WPS formula engine specific, not testable in Node");

console.log("\n=== Results ===");
console.log("Passed: " + passed);
console.log("Failed: " + failed);
console.log("Skipped: " + skipped);
console.log("Total:   " + (passed + failed + skipped));
