#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inject_jsa880_simple.py  v3 — 最稳妥版本
═══════════════════════════════════════════════════════════════
策略:用 ElementTree 安全解析现有 JDEData.bin,只删/改 ThisWorkbook 模块,
其他 codemodule 完整保留(不重新序列化它们,避免引入 XML 格式差异)。
"""
import sys
import re
import shutil
import zipfile
import tempfile
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path


_AMP = chr(38); _LT = chr(60); _GT = chr(62); _QUOT = chr(34); _APOS = chr(39)


def encode_for_xml(s):
    """
    完整实体编码(只转 & < > 和控制字符,' 和 " 留 raw)。
    注意:ElementTree 会自动转义 & < > 但不转义 ' "。我们手动转所有,
    这样结果 100% 是 ElementTree 直读直写的格式。
    """
    out = []
    for ch in s:
        cp = ord(ch)
        if ch == _AMP:    out.append(_AMP + 'amp;')
        elif ch == _LT:   out.append(_LT + 'lt;')
        elif ch == _GT:   out.append(_GT + 'gt;')
        elif ch == _NEWLINE: out.append('&#x0A;')
        elif ch == _CR:   out.append('&#x0D;')
        elif ch == _TAB:  out.append('&#x09;')
        elif cp < 0x20:
            out.append('&#x' + format(cp, 'x') + ';')
        else:
            out.append(ch)
    return ''.join(out)


_NEWLINE = chr(10); _CR = chr(13); _TAB = chr(9)


WRAPPER = '''/* ╔════════════════════════════════════════════════════════════╗
   ║ T24 wrapper:用 throw 把 WPS 实际类型推到 cell(用户可见)  ║
   ╚════════════════════════════════════════════════════════════╝ */
console.log('🔧 [T24 wrapper] loaded');

function _kStruct(v) {
    if (v == null) return String(v);
    var t = typeof v;
    var isArr = Array.isArray(v);
    var hasFilter = isArr ? (typeof v.filter) : '?';
    var keys = [];
    try { for (var k in v) { if (keys.length < 8) keys.push(k); } } catch(e){}
    return t + (isArr ? '[Array,len=' + v.length + ',filter=' + hasFilter + ']' : '') +
           ' {hasAddr=' + (!!v.Address) + ',hasCells=' + (!!v.Cells) +
           ',methods=[' + (isArr ? hasFilter : '?') + '],keys=[' + keys.join(',') + ']}';
}

function _kCoerce(v) {
    if (v == null) return v;
    // (1) Range → Value2
    if (typeof v === 'object' && v && v.Address && typeof v.Value2 !== 'undefined') {
        try {
            var __val = v.Value2;
            if (__val && __val !== v) v = __val;
        } catch (__e) {}
    }
    // (2) 非 Array 但有 length → 转
    if (!Array.isArray(v) && v != null && typeof v === 'object' && typeof v.length === 'number' && v.length >= 0) {
        try {
            var __tmp = [];
            for (var __i = 0; __i < v.length; __i++) {
                var __it = v[__i];
                if (typeof __it !== 'undefined') __tmp.push(__it);
            }
            v = __tmp;
        } catch (__e) {}
    }
    // (3) host array(无 .filter)→ JSON 强转
    if (Array.isArray(v) && typeof v.filter !== 'function') {
        try { v = JSON.parse(JSON.stringify(v)); } catch (__e) {
            try { v = Array.from(v); } catch (__e2) {}
        }
    }
    // (4) 1D → 2D n×1
    if (Array.isArray(v) && v.length > 0 && !Array.isArray(v[0])) {
        var __n2d = [];
        for (var __i = 0; __i < v.length; __i++) __n2d.push([v[__i]]);
        v = __n2d;
    }
    // (5) 2D 但 row 不是 Array → normalize
    if (Array.isArray(v) && v.length > 0) {
        var __bad = false;
        for (var __i = 0; __i < v.length; __i++) {
            if (!Array.isArray(v[__i])) { __bad = true; break; }
        }
        if (__bad) {
            var __fix = [];
            for (var __i = 0; __i < v.length; __i++) {
                if (Array.isArray(v[__i])) __fix.push(v[__i]);
                else __fix.push([v[__i]]);
            }
            v = __fix;
        }
    }
    return v;
}

function k(fn) {
    // 🎯 T80 v4.0.13: 真正调 JSA.k(WPS UDF 入口); 不再是 T25 调试存根
    var args = [];
    for (var i = 1; i < arguments.length; i++) args.push(_kCoerce(arguments[i]));
    return JSA.k.apply(null, [fn].concat(args));
}
function jsaLambda(fn) {
    var args = [];
    for (var i = 1; i < arguments.length; i++) args.push(_kCoerce(arguments[i]));
    return JSA.k.apply(null, [fn].concat(args));
}

'''


def main():
    parser = argparse.ArgumentParser(description='最简化 inline 注入 v3 — ET 解析+只改 ThisWorkbook')
    parser.add_argument('--target', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'JS880教案/第03章/3-28/KO一切的k函数.xlsm')
    parser.add_argument('--source', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'js880/JSA880.js')
    args = parser.parse_args()

    target = Path(args.target).resolve()
    source_js = Path(args.source).resolve()

    print('═' * 60)
    print(' T15 simple v3 — ElementTree 安全修改 ThisWorkbook')
    print('═' * 60)
    print(f'  目标: {target}')

    # 备份
    backup = target.with_suffix(target.suffix + '.t15.bak')
    if not backup.exists():
        shutil.copy2(str(target), str(backup))
        print(f'💾 备份: {backup.name}')

    jsa_code = source_js.read_text(encoding='utf-8')
    print(f'📖 JSA880.js: {len(jsa_code):,} 字符')

    # 关键:JSA880.js 内部的顶层 function k() / function jsaLambda() 会和 wrapper 冲突
    # WPS UDF scanner 只扫开头,我们的 wrapper 必须在开头
    # 但 JSA880.js 内部的同名函数会 shadow 我们的 wrapper
    # 修法:把 JSA880.js 内部的顶层 function k / jsaLambda 重命名
    # 匹配: `function k(fn) {` 和 `function jsaLambda(fn) {`(顶层,只此 1 处)
    jsa_code = jsa_code.replace(
        'function k(fn) {\n    return JSA.k.apply(null, arguments);\n}',
        'function __JSA_k_internal(fn) {\n    return JSA.k.apply(null, arguments);\n}'
    ).replace(
        'function jsaLambda(fn) {\n    return JSA.k.apply(null, arguments);\n}',
        'function __JSA_jL_internal(fn) {\n    return JSA.k.apply(null, arguments);\n}'
    )

    # wrapper 必须在最开头(WPS UDF scanner 只扫开头)
    combined = WRAPPER + '// ═══════════════════════════════════════════════════════════════\n' \
               '// 以下为 JSA880 v5.0 主体代码(inline 到 ThisWorkbook)\n' \
               '// ═══════════════════════════════════════════════════════════════\n\n' \
               + jsa_code

    # 用 ElementTree 解析
    with tempfile.TemporaryDirectory() as tmpdir:
        workdir = Path(tmpdir)
        target_bin = workdir / 'JDEData.bin'
        with zipfile.ZipFile(str(target)) as zf:
            with zf.open('xl/JDEData.bin') as src, open(target_bin, 'wb') as dst:
                shutil.copyfileobj(src, dst)

        orig_size = target_bin.stat().st_size
        print(f'📂 JDEData.bin: {orig_size:,} bytes')

        # 用 ET 解析(允许 CDATA)
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

        # 1) 删除所有 JSA880 / mJSA880 codemodule
        purged = []
        for cm in list(root.findall('codemodule')):
            nm = cm.get('name', '')
            if nm == 'JSA880' or nm.startswith('mJSA880'):
                root.remove(cm)
                purged.append(nm)
        if purged:
            print(f'🧹 删 codemodule: {purged}')

        # 2) 找/创建 ThisWorkbook
        thiswb = None
        for cm in root.findall('codemodule'):
            if cm.get('name') == 'ThisWorkbook':
                thiswb = cm
                break
        if thiswb is None:
            # 新建
            thiswb = ET.Element('codemodule', {'name': 'ThisWorkbook', 'id': '999'})
            ET.SubElement(thiswb, 'window', {'cursorpos': '0', 'actived': 'true', 'visible': 'false'})
            funcdata = root.find('functionsdata')
            if funcdata is not None:
                idx = list(root).index(funcdata)
                root.insert(idx, thiswb)
            else:
                root.append(thiswb)
        else:
            # 改 id
            thiswb.set('id', '999')
            # 确保有 window
            win = thiswb.find('window')
            if win is None:
                win = ET.Element('window', {'cursorpos': '0', 'actived': 'true', 'visible': 'false'})
                thiswb.insert(0, win)
            # 清空所有子节点(window 之外)
            for child in list(thiswb):
                if child.tag != 'window':
                    thiswb.remove(child)

        # 3) 替换/添加 codetext
        ct = thiswb.find('codetext')
        if ct is None:
            ct = ET.SubElement(thiswb, 'codetext')
        # ⚠️ 不调用 encode_for_xml,ElementTree 会自动处理 & < >
        # 但需要手动把 \n 替换为 ElementTree 能正确序列化的形式
        # 直接放 raw text,ElementTree 在 write 时处理
        ct.text = combined

        # 4) 改 activemodule
        am = root.find('activemodule')
        if am is None:
            am = ET.SubElement(root, 'activemodule')
        am.text = '999'

        # 5) [T26] 新建独立 codemodule "kWrapper" 放真正的 k() UDF
        # 原因:ThisWorkbook 里的 k() 似乎没被 WPS UDF scanner 识别
        # 独立 codemodule 在最前,确保被扫描到
        # 先删除已有的 kWrapper(如果存在)
        for cm in list(root.findall('codemodule')):
            if cm.get('name') == 'kWrapper':
                root.remove(cm)
        # 找一个新的 id
        existing_ids = [int(cm.get('id', '0')) for cm in root.findall('codemodule') if cm.get('id', '').isdigit()]
        new_id = (max(existing_ids) if existing_ids else 0) + 1
        kwrap = ET.Element('codemodule', {'name': 'kWrapper', 'id': str(new_id)})
        ET.SubElement(kwrap, 'window', {'cursorpos': '0', 'actived': 'true', 'visible': 'true'})
        kw_ct = ET.SubElement(kwrap, 'codetext')
        kw_ct.text = '''// T26 kWrapper — 独立 codemodule 的 k() UDF
function _kStruct(v) {
    if (v == null) return String(v);
    var t = typeof v;
    var isArr = Array.isArray(v);
    var hasFilter = isArr ? (typeof v.filter) : '?';
    var keys = [];
    try { for (var k in v) { if (keys.length < 8) keys.push(k); } } catch(e){}
    return t + (isArr ? '[Array,len=' + v.length + ',filter=' + hasFilter + ']' : '') +
           ' {hasAddr=' + (!!v.Address) + ',hasCells=' + (!!v.Cells) +
           ',methods=[' + (isArr ? hasFilter : '?') + '],keys=[' + keys.join(',') + ']}';
}
function _kCoerce(v) {
    if (v == null) return v;
    if (typeof v === 'object' && v && v.Address && typeof v.Value2 !== 'undefined') {
        try { var __val = v.Value2; if (__val && __val !== v) v = __val; } catch (__e) {}
    }
    if (!Array.isArray(v) && v != null && typeof v === 'object' && typeof v.length === 'number' && v.length >= 0) {
        try {
            var __tmp = [];
            for (var __i = 0; __i < v.length; __i++) { var __it = v[__i]; if (typeof __it !== 'undefined') __tmp.push(__it); }
            v = __tmp;
        } catch (__e) {}
    }
    if (Array.isArray(v) && typeof v.filter !== 'function') {
        try { v = JSON.parse(JSON.stringify(v)); } catch (__e) { try { v = Array.from(v); } catch (__e2) {} }
    }
    if (Array.isArray(v) && v.length > 0 && !Array.isArray(v[0])) {
        var __n2d = [];
        for (var __i = 0; __i < v.length; __i++) __n2d.push([v[__i]]);
        v = __n2d;
    }
    return v;
}
function k(fn) {
    // 诊断分支
    if (fn === '__T26_TYPE__') {
        var a = arguments[1];
        return 'T26_TYPE|IN:' + _kStruct(a) + '|OUT:' + _kStruct(_kCoerce(a));
    }
    // 正常分支
    var args = [];
    for (var i = 1; i < arguments.length; i++) args.push(_kCoerce(arguments[i]));
    return JSA.k.apply(null, [fn].concat(args));
}
'''
        # 插到 ThisWorkbook 之前(让 WPS scanner 先扫到 kWrapper)
        idx_tw = list(root).index(thiswb) if thiswb in list(root) else len(list(root))
        root.insert(idx_tw, kwrap)
        print(f'🆕 新建 codemodule kWrapper (id={new_id}) 在 ThisWorkbook 之前')

        # 6) 写回
        tmp = Path(str(target_bin) + '.tmp')
        # ET.indent 会重新格式化,会破坏原 codemodule 顺序
        # 直接 write 不 indent
        tree.write(str(tmp), encoding='UTF-8', xml_declaration=True, short_empty_elements=True)
        shutil.move(str(tmp), str(target_bin))

        new_size = target_bin.stat().st_size
        print(f'💽 JDEData.bin: {orig_size:,} → {new_size:,} bytes')

        # 验证 XML 合法
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
    print('✅ 注入完成!')


if __name__ == '__main__':
    main()
