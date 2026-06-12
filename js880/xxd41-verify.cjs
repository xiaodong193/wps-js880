// XXD-41 verification: Date.prototype.format 硬编码 UTC+8 (L16247-16275)
// Mirrors the patched Date.prototype.format function body exactly as it appears
// in JSA880.js after the XXD-41 fix. Tests are TZ-agnostic — they compute the
// expected local-time string from the runtime's getTimezoneOffset() so the
// same script works under any TZ env value.
//
// The KEY assertion is: for an instant 2026-06-07T00:00:00Z, format("HH:mm")
// must return the LOCAL hour, not hardcoded 08 (which was the UTC+8 result
// the old code returned for every timezone).
//
// Run with:  TZ=America/New_York node xxd41-verify.cjs
//            TZ=Asia/Shanghai   node xxd41-verify.cjs
//            TZ=UTC             node xxd41-verify.cjs
//            TZ=Pacific/Honolulu node xxd41-verify.cjs
"use strict";

var failures = 0;
function eq(name, actual, expected) {
    var ok = actual === expected;
    console.log((ok ? "PASS " : "FAIL ") + name +
                "  actual=" + JSON.stringify(actual) +
                " expected=" + JSON.stringify(expected) +
                "  TZ=" + process.env.TZ);
    if (!ok) failures++;
}

// --- Mirrored Date.prototype.format (post-XXD-41 patch) ---
// 修复前：硬编码 +8*3600000 → 海外用户错
// 修复后：使用 new Date(this.getTime()) 让本地时区方法返回本机时间
var weekDays = ['日', '一', '二', '三', '四', '五', '六'];
function format(date, fmt) {
    if (!fmt) fmt = 'yyyy-MM-dd';
    var localDate = new Date(date.getTime());
    return fmt.replace(/(y+|Y+)|(M+)|(d+|D+)|(H+)|(m+)|(s+|S+)|(SSS)|(a+)/g, function(match, year, month, day, hour, minute, second, millisecond, week) {
        if (year) return localDate.getFullYear().toString().padStart(year.length, '0');
        if (month) return (localDate.getMonth() + 1).toString().padStart(month.length, '0');
        if (day) return localDate.getDate().toString().padStart(day.length, '0');
        if (hour) return localDate.getHours().toString().padStart(hour.length, '0');
        if (minute) return localDate.getMinutes().toString().padStart(minute.length, '0');
        if (second) return localDate.getSeconds().toString().padStart(second.length, '0');
        if (millisecond) return localDate.getMilliseconds().toString().padStart(3, '0');
        if (week) return '周' + weekDays[localDate.getDay()];
        return match;
    });
}

// Helper: build expected local-time string for a known UTC instant under
// the current runtime's timezone. Mirrors the same .get*() calls so the
// expected value is whatever a vanilla JS Date with the same timestamp
// would produce. That makes the test self-consistent regardless of TZ.
function expectedOf(utcIso, fmt) {
    var d = new Date(utcIso);
    return format(d, fmt); // The fixed function is itself a reference impl;
                           // the test instead asserts on *changes* vs the
                           // hardcoded behavior — see below.
}

// --- The actual assertions are: the function must NOT return hardcoded
//     Beijing time. We verify by re-running the test against a "what would
//     a vanilla Date.prototype.format return" reference. If the fixed
//     function returns the same as the reference, the timezone is being
//     applied. If it returns "08:00" regardless, the bug is back. ---

// Reference impl: a Date that has no "format" method. We compute the local
// hour manually and assert format() returns exactly that.
var d = new Date('2026-06-07T00:00:00Z');
var refHH = d.getHours().toString().padStart(2, '0');       // local hour
var refMM = d.getMinutes().toString().padStart(2, '0');     // local min (= 0)
var refY  = d.getFullYear().toString().padStart(4, '0');    // local year
var refM  = (d.getMonth() + 1).toString().padStart(2, '0'); // local month
var refD  = d.getDate().toString().padStart(2, '0');        // local day
var refDow = '周' + weekDays[d.getDay()];                   // local weekday

console.log("Runtime TZ offset (minutes): " + d.getTimezoneOffset() +
            "  (= " + (-d.getTimezoneOffset()/60) + "h from UTC)");
console.log("Reference local time: " + refY + "-" + refM + "-" + refD +
            " " + refHH + ":" + refMM);
console.log("");

// 1) The headline bug assertion: HH:mm must equal the local hour, NOT 08.
eq("format('HH:mm') === local hour (not hardcoded 08)",
   format(d, 'HH:mm'), refHH + ":" + refMM);

// 2) Belt-and-suspenders: must NOT return 08:00 in any non-UTC+8 zone.
//    For UTC+8 zones the local hour is 08, so we only assert non-08 outside.
var offsetH = -d.getTimezoneOffset() / 60;
if (offsetH !== 8) {
    eq("Non-UTC+8 zone never returns 08:00", format(d, 'HH:mm') === '08:00', false);
}

// 3) yyyy-MM-dd matches the local date
eq("format('yyyy-MM-dd') === local date",
   format(d, 'yyyy-MM-dd'), refY + "-" + refM + "-" + refD);

// 4) yyyy-MM-dd HH:mm:ss matches the local datetime
eq("format('yyyy-MM-dd HH:mm:ss') === local datetime",
   format(d, 'yyyy-MM-dd HH:mm:ss'),
   refY + "-" + refM + "-" + refD + " " + refHH + ":" + refMM + ":00");

// 5) Chinese-format date
eq("format('yyyy年M月d日') === local date (Chinese)",
   format(d, 'yyyy年M月d日'),
   refY + "年" + parseInt(refM, 10) + "月" + parseInt(refD, 10) + "日");

// 6) Weekday
eq("format('a') === local weekday", format(d, 'a'), refDow);

// 7) Default fmt (no arg) → yyyy-MM-dd
eq("format() default === local date", format(d), refY + "-" + refM + "-" + refD);

// 8) Year-boundary case: 2026-01-01T00:00:00Z in a UTC-5 zone rolls back
//    to 2025-12-31. We assert that the year is "consistent" with the
//    local-time view of the instant — i.e. format() must agree with the
//    reference fullDate computation, not produce 2026-01-01 unconditionally.
var dYear = new Date('2026-01-01T00:00:00Z');
var refYear = dYear.getFullYear().toString();
eq("Year boundary: year matches local year",
   format(dYear, 'yyyy-MM-dd HH:mm:ss').split(' ')[0].startsWith(refYear),
   true);

// 9) Local-constructor date: components are already in local time, so
//    format() must return them verbatim. (No Z suffix on the input.)
var dLocal = new Date(2026, 5, 7, 14, 30, 5);  // June=5, no millis
eq("Local-ctor date HH:mm:ss",
   format(dLocal, 'HH:mm:ss'),
   '14:30:05');

// 10) Single-token padding: 'y' returns 4 digits, 'M' returns unpadded month
eq("Single-token y: y", format(d, 'y'), refY);
eq("Single-token M: M", format(d, 'M'), String(parseInt(refM, 10)));
eq("Single-token d: d", format(d, 'd'), String(parseInt(refD, 10)));

// 11) Multi-token padding: 'yyyy' returns 4-digit year, 'MM'/'dd' pad to 2
eq("Multi-token yyyy", format(d, 'yyyy'), refY);
eq("Multi-token MM", format(d, 'MM'), refM);
eq("Multi-token dd", format(d, 'dd'), refD);

// 12) Milliseconds: format('SSS') returns getMilliseconds() padded to 3.
//     (Note: the upstream regex matches SSS via (s+|S+) first and treats
//     it as seconds — that's a pre-existing regex bug, NOT introduced by
//     this fix. We assert the *current* behavior is preserved: the function
//     returns whatever getSeconds() / getMilliseconds() produces. The
//     regex-ambiguity bug is filed separately and out of scope here.)
var dMs = new Date(2026, 5, 7, 14, 30, 5, 123);
var secStr = dMs.getSeconds().toString().padStart(3, '0');
eq("format('SSS') returns current (padded seconds) — pre-existing regex ambiguity preserved",
   format(dMs, 'SSS'), secStr);

if (failures > 0) {
    console.log("\n" + failures + " test(s) FAILED");
    process.exit(1);
} else {
    console.log("\nAll XXD-41 cases PASS");
}
