// XXD-33 verification: SuperMap._aggregate post-fix (let→var) — exercise 5 cases.
// Mirrors the function body exactly as it appears in JSA880.js L18749-18763.
"use strict";

function _aggregate(values, aggType) {
    var numericValues = [];
    for (var i = 0; i < values.length; i++) {
        if (typeof values[i] === "number" && !isNaN(values[i])) numericValues.push(values[i]);
    }
    if (numericValues.length === 0) {
        switch (aggType) {
            case "sum": return 0;
            case "avg": return 0;
            case "count": return values.length;
            case "min": return null;
            case "max": return null;
            default: return null;
        }
    }
    switch (aggType) {
        case "sum":   { var sum = 0;   for (var j = 0; j < numericValues.length; j++) sum += numericValues[j]; return sum; }
        case "avg":
        case "average": { var total = 0; for (var k = 0; k < numericValues.length; k++) total += numericValues[k]; return total / numericValues.length; }
        case "count": return values.length;
        case "min":   { var minVal = numericValues[0]; for (var m = 1; m < numericValues.length; m++) if (numericValues[m] < minVal) minVal = numericValues[m]; return minVal; }
        case "max":   { var maxVal = numericValues[0]; for (var n = 1; n < numericValues.length; n++) if (numericValues[n] > maxVal) maxVal = numericValues[n]; return maxVal; }
        default: return null;
    }
}

var failures = 0;
function eq(name, actual, expected) {
    var ok = actual === expected || (typeof actual === "number" && typeof expected === "number" && Math.abs(actual - expected) < 1e-9);
    console.log((ok ? "PASS " : "FAIL ") + name + "  actual=" + JSON.stringify(actual) + " expected=" + JSON.stringify(expected));
    if (!ok) failures++;
}

// Issue's repro: SuperMap._aggregate([1,2,3], 'sum') → 6
eq("sum([1,2,3])",        _aggregate([1, 2, 3], "sum"),   6);
eq("count([1,2,3])",      _aggregate([1, 2, 3], "count"), 3);
eq("avg([1,2,3])",        _aggregate([1, 2, 3], "avg"),   2);
eq("min([3,1,2])",        _aggregate([3, 1, 2], "min"),   1);
eq("max([3,1,2])",        _aggregate([3, 1, 2], "max"),   3);

// Edge cases called out in the scope: non-numeric filtering + empty-numeric path
eq("sum with NaN/null",   _aggregate([1, NaN, 2, null, 3], "sum"),   6);
eq("count includes null", _aggregate([1, NaN, 2, null, 3], "count"), 5);
eq("avg on numeric-only", _aggregate([10, 20, 30], "avg"),           20);
eq("min empty-numeric",   _aggregate([NaN, null, "x"], "min"),       null);
eq("max empty-numeric",   _aggregate([NaN, null, "x"], "max"),       null);
eq("sum empty-numeric",   _aggregate([NaN, null, "x"], "sum"),       0);
eq("avg empty-numeric",   _aggregate([NaN, null, "x"], "avg"),       0);
eq("count empty-numeric", _aggregate([NaN, null, "x"], "count"),     3);

// 'average' alias
eq("average alias",       _aggregate([1, 2, 3, 4], "average"), 2.5);

if (failures > 0) {
    console.error("XXD-33 verification FAILED: " + failures + " case(s)");
    process.exit(1);
} else {
    console.log("XXD-33 verification OK — 14 cases passed (sum/count/avg/min/max + edge).");
}
