// XXD-40 verification: IO.readTextFile 默认 encoding 不检测 GBK (L18070-18097).
// Mirrors the function body exactly as it appears in JSA880.js after the
// XXD-40 patch. ActiveXObject is shimmed to throw so the Node fallback path
// runs. Node's fs.readFileSync does not accept 'gbk' in v25, so the inner
// read is shimmed with a TextDecoder-backed helper to faithfully exercise
// the post-sniff decode (which is what the WPS ADODB.Stream path does in
// production with Charset='gbk').
"use strict";

var fs = require('fs');
var os = require('os');
var path = require('path');

// --- Sandbox: shim ActiveXObject so the WPS path throws and we fall through. ---
this.ActiveXObject = function () { throw new Error('ActiveXObject unavailable'); };

// --- Node-side read shim: utf-8 via fs.readFileSync, everything else via TextDecoder. ---
// WPS ADODB.Stream accepts 'utf-16le' / 'utf-16be' (with hyphens). Node v25's
// fs API and Buffer.toString both dropped the utf16be alias, so the shim
// routes all non-utf8 paths through TextDecoder to faithfully exercise the
// post-sniff decode. Production code is unchanged — it passes the WPS name
// verbatim to ADODB.Stream, which accepts it.
function nodeRead(filePath, encoding) {
    if (encoding === 'utf-8' || encoding === 'utf8') {
        return fs.readFileSync(filePath, { encoding: 'utf-8' });
    }
    var bytes = fs.readFileSync(filePath);
    if (encoding === 'gbk' || encoding === 'gb2312') {
        return new TextDecoder('gbk').decode(bytes);
    }
    if (encoding === 'utf-16le' || encoding === 'utf16le') {
        return new TextDecoder('utf-16le').decode(bytes);
    }
    if (encoding === 'utf-16be' || encoding === 'utf16be') {
        return new TextDecoder('utf-16be').decode(bytes);
    }
    // Fallback: try the encoding name as-is.
    return new TextDecoder(encoding).decode(bytes);
}

// --- Mirrored IO.readTextFile (post-XXD-40 patch), with nodeRead shim. ---
var IO = {};
IO.readTextFile = function(filePath, encoding) {
    if (!encoding) {
        var sniffHead = null;
        var sniffLen = 0;
        try {
            var fsSniff = require('fs');
            var fd = fsSniff.openSync(filePath, 'r');
            try {
                var sniffBuf = Buffer.alloc(3);
                sniffLen = fsSniff.readSync(fd, sniffBuf, 0, 3, 0);
                sniffHead = [sniffBuf[0], sniffBuf[1], sniffBuf[2]];
            } finally { fsSniff.closeSync(fd); }
        } catch (eNodeSniff) {
            try {
                var fsoSniff = new ActiveXObject('Scripting.FileSystemObject');
                if (fsoSniff.FileExists(filePath)) {
                    var bs = new ActiveXObject('ADODB.Stream');
                    bs.Type = 1; bs.Mode = 1; bs.Open();
                    bs.LoadFromFile(filePath);
                    var nSniff = bs.Size > 3 ? 3 : bs.Size;
                    sniffHead = []; sniffLen = nSniff;
                    for (var iSniff = 0; iSniff < nSniff; iSniff++) {
                        var byte = bs.Read(1);
                        sniffHead.push(typeof byte === 'number' ? byte : (byte && byte[0]) || 0);
                    }
                    bs.Close();
                }
            } catch (eAxSniff) { /* sniff unavailable — keep utf-8 default */ }
        }
        if (sniffHead && sniffLen >= 3 && sniffHead[0] === 0xEF && sniffHead[1] === 0xBB && sniffHead[2] === 0xBF) {
            encoding = 'utf-8';
        } else if (sniffHead && sniffLen >= 2 && sniffHead[0] === 0xFF && sniffHead[1] === 0xFE) {
            encoding = 'utf-16le';
        } else if (sniffHead && sniffLen >= 2 && sniffHead[0] === 0xFE && sniffHead[1] === 0xFF) {
            encoding = 'utf-16be';
        } else if (sniffHead) {
            encoding = 'gbk';
        } else {
            encoding = 'utf-8';
        }
    }

    try {
        var fso = new ActiveXObject('Scripting.FileSystemObject');
        if (!fso.FileExists(filePath)) return null;
        var stream = new ActiveXObject('ADODB.Stream');
        stream.Type = 2; stream.Charset = encoding;
        stream.Open(); stream.LoadFromFile(filePath);
        var content = stream.ReadText(); stream.Close();
        if (content && content.charCodeAt(0) === 0xFEFF) content = content.slice(1);
        return content;
    } catch (e) {
        try {
            var content2 = nodeRead(filePath, encoding);
            if (content2 && content2.charCodeAt(0) === 0xFEFF) content2 = content2.slice(1);
            return content2;
        } catch (e2) {
            console.warn('IO.readTextFile:', e2.message);
            return null;
        }
    }
};

// --- Test harness ---
var failures = 0;
function eq(name, actual, expected) {
    var ok = actual === expected;
    console.log((ok ? "PASS " : "FAIL ") + name + "  actual=" + JSON.stringify(actual) + " expected=" + JSON.stringify(expected));
    if (!ok) failures++;
}

var tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xxd40-'));

// Pre-encoded GBK bytes for "你好,世界" (so the file is GBK with no BOM).
// 你 C4 E3  好 BA C3  , 2C  世 CA C0  界 BD E7
var gbkBytes = Buffer.from([0xC4,0xE3,0xBA,0xC3,0x2C,0xCA,0xC0,0xBD,0xE7]);
var gbkPath = path.join(tmpDir, 'gbk.txt');
fs.writeFileSync(gbkPath, gbkBytes);

// 1) GBK no-BOM file → expect '你好,世界'
eq('GBK no-BOM default → 你好,世界', IO.readTextFile(gbkPath), '你好,世界');

// 2) UTF-8 BOM file → BOM stripped, content preserved
var utf8BomPath = path.join(tmpDir, 'utf8bom.txt');
fs.writeFileSync(utf8BomPath, Buffer.concat([Buffer.from([0xEF, 0xBB, 0xBF]), Buffer.from('Hello,世界', 'utf-8')]));
eq('UTF-8 BOM default → Hello,世界', IO.readTextFile(utf8BomPath), 'Hello,世界');

// 3) UTF-16LE BOM file → decoded
var utf16lePath = path.join(tmpDir, 'utf16le.txt');
fs.writeFileSync(utf16lePath, Buffer.concat([Buffer.from([0xFF, 0xFE]), Buffer.from('WPS用户', 'utf16le')]));
eq('UTF-16LE BOM default → WPS用户', IO.readTextFile(utf16lePath), 'WPS用户');

// 4) UTF-16BE BOM file → decoded
var utf16bePath = path.join(tmpDir, 'utf16be.txt');
fs.writeFileSync(utf16bePath, Buffer.concat([Buffer.from([0xFE, 0xFF]), Buffer.from('WPS用户', 'utf16le').swap16()]));
eq('UTF-16BE BOM default → WPS用户', IO.readTextFile(utf16bePath), 'WPS用户');

// 5) Explicit encoding passthrough: GBK file read with 'utf-8' explicit → mojibake
//    (proves the explicit path is preserved, NOT silently overwritten)
eq('Explicit utf-8 on GBK file is preserved (mojibake expected)',
   IO.readTextFile(gbkPath, 'utf-8') === '你好,世界', false);

// 6) Explicit encoding: ask for 'gbk' on UTF-8 BOM file → NOT the original utf-8 string
eq('Explicit gbk on UTF-8 file is preserved (no silent override)',
   IO.readTextFile(utf8BomPath, 'gbk') === 'Hello,世界', false);

// 7) Missing file → returns null
eq('Missing file returns null', IO.readTextFile(path.join(tmpDir, 'nope.txt')), null);

// 8) Reproduces the issue's exact repro: a GBK legacy file with no encoding arg
//    returns the correct Chinese text (the documented failure was mojibake).
var reproPath = path.join(tmpDir, 'legacy.txt');
fs.writeFileSync(reproPath, Buffer.from([0xC4,0xE3,0xBA,0xC3])); // 你好
eq('Issue repro: IO.readTextFile("/tmp/legacy.txt") === "你好"',
   IO.readTextFile(reproPath), '你好');

// Cleanup
try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch (e) {}

if (failures > 0) {
    console.log("\n" + failures + " test(s) FAILED");
    process.exit(1);
} else {
    console.log("\nAll XXD-40 cases PASS");
}
