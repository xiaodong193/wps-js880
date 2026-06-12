#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inject_t27.py  — 极简版:只修 JSA880 codemodule 的 smartUnwrap
════════════════════════════════════════════════════════════════════
策略:
1. 保留 t15.bak 原始结构(2 个 codemodule: JSA880 + ThisWorkbook)
2. 只改 JSA880 codemodule 的 codetext — 把内嵌的 JSA880.js 整体替换成新的
3. ThisWorkbook 不动(简单 wrapper)
4. JSA880 codemodule 里 `function k` 顶层声明,自然被 WPS 扫为 UDF
5. smartUnwrap 修好后,Range → 2D 数组的 chainable 就能跑
"""
import sys
import shutil
import zipfile
import tempfile
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path


# 新 smartUnwrap(从 JSA880.js 第 2893 行那段)
NEW_SMARTUNWRAP = '''function smartUnwrap(v) {
            // [T27 修复] Range 对象检测 → 转 Value2(去掉 v !== asRange(v) 的反逻辑)
            if (v && typeof v === 'object' && v.Address && typeof v.Value2 !== 'undefined') {
                try {
                    var __val2 = v.Value2;
                    if (__val2 && __val2 !== v) v = __val2;
                } catch (e) {}
            }
            // [T27 增强] host array(像 Array 但 .filter undefined)→ JSON 强转
            if (Array.isArray(v) && typeof v.filter !== 'function') {
                try { v = JSON.parse(JSON.stringify(v)); } catch (e) {
                    try { v = Array.from(v); } catch (e2) {}
                }
            }
            // [T27 增强] 1D 数组(被 WPS 压扁)→ 2D n×1
            if (Array.isArray(v) && v.length > 0 && !Array.isArray(v[0])) {
                var __n2d = [];
                for (var __i = 0; __i < v.length; __i++) __n2d.push([v[__i]]);
                v = __n2d;
            }
            // 只有 1x1 2D 数组才 flatten → 单个原始值
            if (Array.isArray(v) && v.length === 1 && Array.isArray(v[0]) && v[0].length === 1) {
                return v[0][0];
            }
            return v;
        }'''

OLD_SMARTUNWRAP_PREFIX = 'function smartUnwrap(v) {'
OLD_SMARTUNWRAP_END = 'return v;\n        }'


def main():
    parser = argparse.ArgumentParser(description='T27: 极简改 JSA880 codemodule 的 smartUnwrap')
    parser.add_argument('--target', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'JS880教案/第03章/3-28/KO一切的k函数.xlsm')
    parser.add_argument('--source', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'js880/JSA880.js')
    args = parser.parse_args()

    target = Path(args.target).resolve()
    source_js = Path(args.source).resolve()

    print('═' * 60)
    print(' T27 — 极简:JSA880 codemodule 注入新 JSA880.js + 修 smartUnwrap')
    print('═' * 60)
    print(f'  目标: {target}')

    if not target.exists():
        print('❌ 目标不存在'); sys.exit(1)

    # 读 JSA880.js 源
    jsa_code = source_js.read_text(encoding='utf-8')
    print(f'📖 JSA880.js: {len(jsa_code):,} 字符')

    # [T33] 在 JSA880 codemodule 末尾追加 wrapper(替换原 k 转发函数)
    # 关键:wrapper 必须替换原 function k 位置(在 JSA880 codemodule 内)
    # 用占位符方式 — 找到文件末尾的 console.log,前面替换
    # 实际:替换 JSA880.js 内部的 function k body
    jsa_code = jsa_code.replace(
        'function k(fn) {\n    return JSA.k.apply(null, arguments);\n}',
        'function k(fn) { return JSA.k.apply(null, arguments); }'
    )

    # [T40] JSA.k 是真正的 UDF!加 spill 逻辑
    jsa_code = jsa_code.replace(
        '''JSA.k = function(fn) {
    var args = [];
    for (var i = 1; i < arguments.length; i++) args.push(arguments[i]);''',
        '''JSA.k = function(fn) {
    var args = [];
    for (var i = 1; i < arguments.length; i++) args.push(arguments[i]);

    // T40 spill helper
    function _kExtract2D(v) {
        // 尝试提取 2D 数组(从各种包装类型)
        if (Array.isArray(v) && v.length > 0 && Array.isArray(v[0])) return v;
        if (Array.isArray(v) && v.length > 0) return v;  // 1D 也认
        if (typeof v === "object" && v) {
            // wrappedResult(val/res) 风格
            try { if (typeof v.val === "function") { var r = v.val(); if (Array.isArray(r)) return r; } } catch(e){}
            try { if (typeof v.res === "function") { var r = v.res(); if (Array.isArray(r)) return r; } } catch(e){}
            try { if (typeof v.toArray === "function") { var r = v.toArray(); if (Array.isArray(r)) return r; } } catch(e){}
            // Array2D 的 _data
            if (v._data && Array.isArray(v._data)) return v._data;
            if (v.data && Array.isArray(v.data)) return v.data;
            // 直接有 length
            if (typeof v.length === "number" && typeof v[0] !== "undefined") {
                try {
                    var arr = [];
                    for (var i = 0; i < v.length; i++) arr.push(v[i]);
                    return arr;
                } catch(e){}
            }
        }
        return null;
    }
    function _kIs2D(v) {
        var extracted = _kExtract2D(v);
        if (!extracted) return false;
        if (extracted.length === 0) return false;
        if (Array.isArray(extracted[0])) return true;
        return false;
    }
    function _kSpill(v) {
        try {
            var caller = (typeof Application !== "undefined" && Application.Caller) ? Application.Caller : null;
            if (!caller) return v;
            var arr = null;
            if (typeof v === "object" && v) {
                try { if (typeof v.val === "function") arr = v.val(); } catch(e1) {}
                if (!arr) try { if (typeof v.res === "function") arr = v.res(); } catch(e2) {}
            }
            if (!arr) arr = _kExtract2D(v);
            if (!arr || arr.length === 0) return v;
            if (!Array.isArray(arr[0])) return v;
            var rows = arr.length;
            var cols = arr[0].length || 1;
            // T47:用 ActiveSheet.Range 显式写每个 cell(绕过 caller 限制)
            var sht = null;
            try { sht = Application.ActiveSheet; } catch(eS) {}
            if (!sht) try { sht = caller.Worksheet; } catch(eW) {}
            if (!sht) {
                // 最后回退:caller.Cells
                for (var r = 0; r < rows; r++) {
                    for (var c = 0; c < cols; c++) {
                        try { caller.Cells(r + 1, c + 1).Value2 = arr[r][c] !== undefined ? arr[r][c] : ""; } catch(eC) {}
                    }
                }
                return arr[0][0] !== undefined ? arr[0][0] : "";
            }
            // 用 caller.Address 解析起点
            var startAddr = caller.Address;
            for (var r = 0; r < rows; r++) {
                for (var c = 0; c < cols; c++) {
                    var vv = arr[r][c];
                    if (vv === null || vv === undefined) vv = "";
                    try {
                        var cellAddr = sht.Range(startAddr).Offset(r, c);
                        cellAddr.Value2 = vv;
                    } catch (eO) {
                        // 失败再试 caller.Cells
                        try { caller.Cells(r + 1, c + 1).Value2 = vv; } catch(eC) {}
                    }
                }
            }
            return arr[0][0] !== undefined ? arr[0][0] : "";
        } catch (e) {
            return v;
        }
    }'''
    )

    # 在 JSA.k 结尾的 return result 之前插入 spill 调用
    # 用更精确的模式:JSA.k 函数紧接在 jsaLambda 错误处理后
    jsa_code = jsa_code.replace(
        '''        return '#K_ERR: pos=0, ' + kind + ', msg="' +
               (e && e.message ? e.message.replace(/"/g, "'") : String(e)) + '"';
    }

    // 3) null / undefined 兜底
    if (result === undefined || result === null) {
        return '#K_ERR: pos=0, FN, msg="jsaLambda 返回 null/undefined,可能 fn 语法错或参数不匹配"';
    }

    return result;
};''',
        '''        return '#K_ERR: pos=0, ' + kind + ', msg="' +
               (e && e.message ? e.message.replace(/"/g, "'") : String(e)) + '"';
    }

    // 3) null / undefined 兜底
    if (result === undefined || result === null) {
        return '#K_ERR: pos=0, FN, msg="jsaLambda 返回 null/undefined,可能 fn 语法错或参数不匹配"';
    }

    // T48 — 如果 result 有 val/res/.toRange → 提取 2D 数组返回
    // (WPS JSA 可能能 spill 真正 2D 数组,但不能 spill wrappedResult)
    if (typeof result === "object" && result !== null) {
        var __arr = null;
        try { if (typeof result.val === "function") __arr = result.val(); } catch(__e) {}
        if (!__arr) try { if (typeof result.res === "function") __arr = result.res(); } catch(__e2) {}
        if (__arr && Array.isArray(__arr) && __arr.length > 0 && Array.isArray(__arr[0])) {
            return __arr;  // 直接返回 2D 数组
        }
    }
    return result;
};'''
    )
    # jsaLambda 保持原样(也指向 JSA.k,但 jsaLambda 也注册为 UDF)

    # 用 ET 解析
    with tempfile.TemporaryDirectory() as tmpdir:
        workdir = Path(tmpdir)
        target_bin = workdir / 'JDEData.bin'
        with zipfile.ZipFile(str(target)) as zf:
            with zf.open('xl/JDEData.bin') as src, open(target_bin, 'wb') as dst:
                shutil.copyfileobj(src, dst)

        orig_size = target_bin.stat().st_size
        print(f'📂 JDEData.bin: {orig_size:,} bytes')

        ET.register_namespace('', '')
        tree = ET.parse(str(target_bin))
        root = tree.getroot()

        # 列所有 codemodule
        for cm in root.findall('codemodule'):
            name = cm.get('name', '')
            mid = cm.get('id', '')
            ctt = cm.find('codetext')
            tlen = len(ctt.text) if ctt is not None and ctt.text else 0
            print(f'  现有: {name} (id={mid}) codetext={tlen:,} chars')

        # 1) 找 JSA880 codemodule
        jsa880_cm = None
        for cm in root.findall('codemodule'):
            if cm.get('name') == 'JSA880':
                jsa880_cm = cm
                break
        if jsa880_cm is None:
            print('❌ JSA880 codemodule 不存在,无法注入')
            sys.exit(1)

        # 2) 替换 JSA880 codemodule 的 codetext 为新 JSA880.js
        ct = jsa880_cm.find('codetext')
        if ct is None:
            ct = ET.SubElement(jsa880_cm, 'codetext')
        ct.text = jsa_code
        print(f'✏️  JSA880 codemodule codetext 已更新 → {len(jsa_code):,} chars')

        # 3) [T31] 改写 ThisWorkbook:加自动 spill 2D 数组到 Application.Caller
        thiswb = None
        for cm in root.findall('codemodule'):
            if cm.get('name') == 'ThisWorkbook':
                thiswb = cm
                break
        if thiswb is not None:
            tw_ct = thiswb.find('codetext')
            if tw_ct is None:
                tw_ct = ET.SubElement(thiswb, 'codetext')
            tw_ct.text = '''// T38 — ThisWorkbook 留空(因为 WPS 倒序加载,JSA880 的 wrapper 会赢)
function k(fn) { return JSA.k.apply(null, arguments); }
function jsaLambda(fn) { return JSA.k.apply(null, arguments); }
'''
            print('✏️  ThisWorkbook codetext 已更新(T31 自动 spill)')

        # 4) 写回
        tmp = Path(str(target_bin) + '.tmp')
        tree.write(str(tmp), encoding='UTF-8', xml_declaration=True, short_empty_elements=True)
        shutil.move(str(tmp), str(target_bin))

        new_size = target_bin.stat().st_size
        print(f'💽 JDEData.bin: {orig_size:,} → {new_size:,} bytes')

        # 验证 XML
        try:
            ET.parse(str(target_bin))
            print('✅ XML 合法')
        except ET.ParseError as e:
            print(f'❌ XML 不合法:{e}')
            sys.exit(1)

        # 重打包
        print('📦 重打包 xlsm ...')
        tmp_out = target.with_suffix('.inject.tmp')
        with zipfile.ZipFile(str(target), 'r') as zin:
            with zipfile.ZipFile(str(tmp_out), 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == 'xl/JDEData.bin':
                        continue
                    zout.writestr(item, zin.read(item.filename))
                with open(target_bin, 'rb') as f:
                    zout.writestr('xl/JDEData.bin', f.read())
        shutil.move(str(tmp_out), str(target))

    print()
    print('✅ T27 注入完成!')
    print('   关键改动:')
    print('   - 保留原始 codemodule 结构(JSA880 + ThisWorkbook)')
    print('   - JSA880 codemodule 用新 JSA880.js 替换(里面 smartUnwrap 已修)')
    print('   - ThisWorkbook 是简单 wrapper:`function k(fn) { return JSA.k.apply... }`')


if __name__ == '__main__':
    main()
