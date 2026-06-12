#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inject_jsa880_inline.py  v4 — base64 编码注入
═══════════════════════════════════════════════════════════════
JS 源码 base64 编码后存到 ThisWorkbook.codetext(CDATA 内),
ThisWorkbook 顶层 wrapper 解码后 eval 加载。

完全规避 XML 实体编码的坑(不依赖 & < > ' " 编码,不依赖 CDATA 分段)。
"""
import os
import sys
import re
import shutil
import zipfile
import tempfile
import base64
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path


def encode_for_xml(s):
    """
    XML text 节点完整实体编码。
    ElementTree 不自动处理 ' 和 "(只在属性里),也不转控制字符。
    手动编码所有 XML 敏感字符。
    """
    out = []
    for ch in s:
        cp = ord(ch)
        if ch == '&':    out.append('&amp;')
        elif ch == '<':   out.append('&lt;')
        elif ch == '>':   out.append('&gt;')
        elif ch == '"':   out.append('&quot;')
        elif ch == "'":   out.append('&apos;')
        elif ch == '\n':  out.append('&#x0A;')
        elif ch == '\r':  out.append('&#x0D;')
        elif ch == '\t':  out.append('&#x09;')
        elif cp < 0x20:
            out.append('&#x' + format(cp, 'x') + ';')
        else:
            out.append(ch)
    return ''.join(out)


# ThisWorkbook 顶层 wrapper:
# 1) 顶层 function k() / jsaLambda() — WPS UDF 注册
# 2) 解码 JSA880 源码并 eval
# 3) 注册 JSA.k 包装 + 链式 parser(若 JSA880 没自带)
# 4) Workbook_Open 自检
WRAPPER_TEMPLATE = '''function k(fn) { return JSA.k.apply(null, arguments); }
function jsaLambda(fn) { return JSA.k.apply(null, arguments); }

/* === T15 inline: 解码并执行 JSA880 主体 === */
var __JSA880_B64 = "{b64}";
(function _t15LoadJSA880() {
    try {
        var __code = atob(__JSA880_B64);
        // 删末尾重复的顶层 function k/jsaLambda(若 JSA880 还带)
        __code = __code.replace(/function\\s+k\\s*\\(\\s*fn\\s*\\)\\s*\\{\\s*return\\s+JSA\\.k\\.apply\\(null,\\s*arguments\\)\\s*;\\s*\\}/g, '');
        __code = __code.replace(/function\\s+jsaLambda\\s*\\(\\s*fn\\s*\\)\\s*\\{\\s*return\\s+JSA\\.k\\.apply\\(null,\\s*arguments\\)\\s*;\\s*\\}/g, '');
        (0, eval)(__code);
        if (typeof Console !== 'undefined') {
            Console.log('✅ [T15] JSA880 主体已 inline 加载');
        }
    } catch (__e) {
        // atob 不存在?用手动 base64 解码
        try {
            var __b = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
            var __code = '';
            for (var __i = 0; __i < __JSA880_B64.length; __i += 4) {
                var __a = __b.indexOf(__JSA880_B64[__i]);
                var __bb = __b.indexOf(__JSA880_B64[__i+1]);
                var __c = __b.indexOf(__JSA880_B64[__i+2]);
                var __d = __b.indexOf(__JSA880_B64[__i+3]);
                __code += String.fromCharCode((__a << 2) | (__bb >> 4));
                if (__c < 64) __code += String.fromCharCode(((__bb & 15) << 4) | (__c >> 2));
                if (__d < 64) __code += String.fromCharCode(((__c & 3) << 6) | __d);
            }
            (0, eval)(__code);
            if (typeof Console !== 'undefined') {
                Console.log('✅ [T15] JSA880 主体已 inline 加载(手动 base64)');
            }
        } catch (__e2) {
            if (typeof Console !== 'undefined') {
                Console.log('❌ [T15] 加载失败:' + (__e2 && __e2.message ? __e2.message : __e2));
            }
        }
    }
})();
'''


def extract_existing_modules(bin_path):
    """提取所有非 JSA880 codemodule 的原始 inner xml"""
    with open(bin_path, 'rb') as f:
        d_bytes = f.read()
    d = d_bytes.decode('utf-8', errors='replace')
    modules = []
    for m in re.finditer(r'<codemodule\s+name="([^"]+)"\s+id="(\d+)"[^>]*>(.*?)</codemodule>', d, re.DOTALL):
        name = m.group(1)
        mid = m.group(2)
        inner = m.group(3)
        if name == 'JSA880' or name.startswith('mJSA880'):
            continue
        modules.append((name, mid, inner))
    return modules


def build_jdedata_xml(modules, thiswb_inner):
    """手写 JDEData.bin 的 XML — ThisWorkbook 不再用 CDATA,用普通 text 节点"""
    out = ['<?xml version="1.0" encoding="UTF-8"?>\n']
    out.append('<document version="2.0">\n')
    out.append('    <name>Project</name>\n')
    out.append('    <property desc="" lock="false" password="" />\n')
    out.append('    <activemodule>999</activemodule>\n')
    for (name, mid, inner) in modules:
        if name == 'ThisWorkbook':
            continue
        out.append(f'    <codemodule name="{name}" id="{mid}">\n')
        if 'window' not in inner.split('</codemodule>')[0][:200]:
            out.append('        <window cursorpos="0" actived="true" visible="true" />\n')
        out.append(inner.rstrip())
        out.append('\n    </codemodule>\n')
    # ThisWorkbook - 不用 CDATA,直接放实体编码后的 text
    out.append('    <codemodule name="ThisWorkbook" id="999">\n')
    out.append('        <window cursorpos="0" actived="true" visible="false" />\n')
    out.append(thiswb_inner.rstrip())
    out.append('\n    </codemodule>\n')
    out.append('    <functionsdata />\n')
    out.append('</document>\n')
    return ''.join(out)


def encode_cdata(s):
    """CDATA 内只需避免 ]]>"""
    return s.replace(']]>', ']]]]><![CDATA[>')


def main():
    parser = argparse.ArgumentParser(description='T15 v4 — base64 inline JSA880 到 ThisWorkbook')
    parser.add_argument('--target', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'JS880教案/第03章/3-28/KO一切的k函数.xlsm')
    parser.add_argument('--source', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'js880/JSA880.js')
    args = parser.parse_args()

    target = Path(args.target).resolve()
    source_js = Path(args.source).resolve()

    print('═' * 60)
    print(' T15 v4 — base64 inline 注入')
    print('═' * 60)
    print(f'  目标: {target}')
    print(f'  源:   {source_js}')

    if not target.exists():
        print('❌ 目标不存在'); sys.exit(1)
    if not source_js.exists():
        print('❌ 源不存在'); sys.exit(1)

    # 备份
    backup = target.with_suffix(target.suffix + '.t15.bak')
    if not backup.exists():
        shutil.copy2(str(target), str(backup))
        print(f'💾 备份: {backup.name}')

    # 读源
    jsa_code = source_js.read_text(encoding='utf-8')
    print(f'📖 JSA880.js: {len(jsa_code):,} 字符')

    # base64 编码
    b64 = base64.b64encode(jsa_code.encode('utf-8')).decode('ascii')
    print(f'🔐 base64: {len(b64):,} 字符')

    # 组装 wrapper(用 replace 不用 format,避免 { } 冲突)
    wrapper = WRAPPER_TEMPLATE.replace('{b64}', b64)

    # inner XML for ThisWorkbook - 用实体编码的普通 text 节点(WPS 可能不认 CDATA)
    thiswb_inner = '<codetext>' + encode_for_xml(wrapper) + '</codetext>'

    # 提取 JDEData.bin
    with tempfile.TemporaryDirectory() as tmpdir:
        workdir = Path(tmpdir)
        target_bin = workdir / 'JDEData.bin'
        with zipfile.ZipFile(str(target)) as zf:
            with zf.open('xl/JDEData.bin') as src, open(target_bin, 'wb') as dst:
                shutil.copyfileobj(src, dst)

        orig_size = target_bin.stat().st_size
        print(f'📂 JDEData.bin: {orig_size:,} bytes')

        existing = extract_existing_modules(target_bin)
        print(f'📋 保留 codemodule: {[m[0] for m in existing]}')

        new_xml = build_jdedata_xml(existing, thiswb_inner)
        with open(target_bin, 'w', encoding='utf-8') as f:
            f.write(new_xml)

        new_size = target_bin.stat().st_size
        print(f'💽 JDEData.bin: {orig_size:,} → {new_size:,} bytes')
        print(f'📝 ThisWorkbook codetext: {len(wrapper):,} 字符 (含 b64)')

        # 重打包
        print('📦 重打包 xlsm ...')
        replace_in_zip(target, 'xl/JDEData.bin', target_bin)

    print()
    print('✅ base64 inline 注入完成!')
    print('   ThisWorkbook 顶层 wrapper:')
    print('   1. function k() / jsaLambda() (UDF 转发)')
    print('   2. atob() 解码 JSA880 主体,eval 加载')
    print('   3. 全在同一上下文,不会 #VALUE!')


def replace_in_zip(zip_path, internal_path, source_file):
    tmp_out = zip_path.with_suffix('.inject.tmp')
    with zipfile.ZipFile(str(zip_path), 'r') as zin:
        with zipfile.ZipFile(str(tmp_out), 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == internal_path:
                    continue
                zout.writestr(item, zin.read(item.filename))
            with open(source_file, 'rb') as f:
                zout.writestr(internal_path, f.read())
    shutil.move(str(tmp_out), str(zip_path))


if __name__ == '__main__':
    main()
