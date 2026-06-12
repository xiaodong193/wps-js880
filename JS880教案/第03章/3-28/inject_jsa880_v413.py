#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inject_jsa880_v413.py  v4.0.13
═══════════════════════════════════════════════════════════════
策略:完全照搬 v4.0.11 原结构 —— JSA880 作为独立 codemodule (id=3)
而不是塞进 ThisWorkbook。WPS UDF scanner 从独立 codemodule 顶层扫
function k / function jsaLambda,直接在 formula 里可调。

不复用 inject_jsa880_simple.py 的 wrapper 套法(那套会导致 #NAME?)。
"""
import sys
import re
import shutil
import zipfile
import tempfile
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description='v4.0.13 注入: JSA880 独立 codemodule')
    parser.add_argument('--target', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'JS880教案/第03章/3-28/KO一切的k函数.xlsm')
    parser.add_argument('--source', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'js880/JSA880.js')
    args = parser.parse_args()

    target = Path(args.target).resolve()
    source_js = Path(args.source).resolve()

    print('═' * 60)
    print(' v4.0.13 注入 — JSA880 独立 codemodule 结构')
    print('═' * 60)
    print(f'  目标: {target}')

    # 备份
    backup = target.with_suffix(target.suffix + '.t81.before_v413.bak')
    if not backup.exists():
        shutil.copy2(str(target), str(backup))
        print(f'💾 备份: {backup.name}')

    jsa_code = source_js.read_text(encoding='utf-8')
    print(f'📖 JSA880.js: {len(jsa_code):,} 字符')

    # ⚠️ 关键: 不重命名 k / jsaLambda — WPS UDF scanner 要从 codemodule
    # 顶层看到 function k(...) / function jsaLambda(...) 才能注册

    # 顶部加一行 console.log 标记,方便 WPS 端确认加载版本
    banner = (
        "// ╔════════════════════════════════════════════════════════════╗\n"
        "// ║ JSA880 v4.0.13 — 2026-06-07                              ║\n"
        "// ║ Bug 修复:                                                 ║\n"
        "// ║   - Array2D.z筛选 加 fN proxy (x.f1, x.f2 ...)             ║\n"
        "// ║   - 链式 .filter((x,i)=>x.fN==...) 走 Array2D.z筛选       ║\n"
        "// ╚════════════════════════════════════════════════════════════╝\n"
        "console.log('✅ JSA880 v4.0.13 loaded (chainable filter + fN proxy)');\n\n"
    )
    combined = banner + jsa_code

    # ThisWorkbook: workbook_open + 顶层 UDF shim(双保险)
    # ⚠️ v4.0.13 关键: 必须在 ThisWorkbook 顶层保留 function k / function jsaLambda
    # 因为 WPS UDF scanner 主要从 ThisWorkbook 顶层扫,JSA880 codemodule 是补充
    # 不加这俩 shim,=k() 报 #NAME?,输入 =k 也没有自动补全
    thiswb_code = (
        "// ThisWorkbook — v4.0.13\n"
        "// 同时提供 workbook_open 事件和顶层 UDF shim(k / jsaLambda)\n"
        "// ⚠️ 不要删 function k / function jsaLambda — WPS 公式引擎从 ThisWorkbook 顶层注册\n"
        "function Workbook_Open() {\n"
        "    try {\n"
        "        if (typeof Console !== 'undefined') Console.log('✅ JSA880 v4.0.13 ThisWorkbook loaded');\n"
        "    } catch (e) {}\n"
        "}\n"
        "\n"
        "/**\n"
        " * [v4.0.13] 顶层 k() UDF — 转发到 JSA.k (JSA880.js 内)\n"
        " * 单元格公式: =k(\"JSA.getIndexs\", 1, 5, 1)\n"
        " */\n"
        "function k(fn) {\n"
        "    return JSA.k.apply(null, arguments);\n"
        "}\n"
        "\n"
        "/**\n"
        " * [v4.0.13] 顶层 jsaLambda() UDF — k() 的全名版本,完全等价\n"
        " * 单元格公式: =jsaLambda(\"JSA.getIndexs\", 1, 5, 1)\n"
        " */\n"
        "function jsaLambda(fn) {\n"
        "    return JSA.k.apply(null, arguments);\n"
        "}\n"
    )

    with tempfile.TemporaryDirectory() as tmpdir:
        workdir = Path(tmpdir)
        target_bin = workdir / 'JDEData.bin'
        with zipfile.ZipFile(str(target)) as zf:
            with zf.open('xl/JDEData.bin') as src, open(target_bin, 'wb') as dst:
                shutil.copyfileobj(src, dst)

        orig_size = target_bin.stat().st_size
        print(f'📂 JDEData.bin: {orig_size:,} bytes')

        # ET 解析
        ET.register_namespace('', '')
        tree = ET.parse(str(target_bin))
        root = tree.getroot()

        # 列出现有 codemodule
        for cm in root.findall('codemodule'):
            name = cm.get('name', '')
            mid = cm.get('id', '')
            ctt = cm.find('codetext')
            tlen = len(ctt.text) if ctt is not None and ctt.text else 0
            print(f'  现有: {name} (id={mid}) {tlen:,} chars')

        # 1) 删除所有 JSA880 / mJSA880 / kWrapper codemodule(清理旧注入产物)
        purged = []
        for cm in list(root.findall('codemodule')):
            nm = cm.get('name', '')
            if nm == 'JSA880' or nm.startswith('mJSA880') or nm == 'kWrapper':
                root.remove(cm)
                purged.append(nm)
        if purged:
            print(f'🧹 删 codemodule: {purged}')

        # 2) 找/创建 JSA880 codemodule (id=3, 模拟 v4.0.11 原始结构)
        jsa_cm = None
        for cm in root.findall('codemodule'):
            if cm.get('name') == 'JSA880':
                jsa_cm = cm
                break
        if jsa_cm is None:
            jsa_cm = ET.Element('codemodule', {'name': 'JSA880', 'id': '3'})
            # ⚠️ v4.0.13 修复: actived 必须是 "true" — 否则 WPS 跳过 UDF 扫描,=k() 报 #NAME?
            ET.SubElement(jsa_cm, 'window', {'cursorpos': '0', 'actived': 'true', 'visible': 'true'})
            # 插到 Module2 之后
            inserted = False
            for i, elem in enumerate(list(root)):
                if elem.tag == 'codemodule' and elem.get('name') == 'Module2':
                    root.insert(i + 1, jsa_cm)
                    inserted = True
                    break
            if not inserted:
                root.append(jsa_cm)
            print('🆕 新建 codemodule JSA880 (id=3)')
        else:
            jsa_cm.set('id', '3')
            # 清空子节点(window 之外)
            for child in list(jsa_cm):
                if child.tag != 'window':
                    jsa_cm.remove(child)
            # ⚠️ v4.0.13 修复: 强制把 window 改成 actived="true"
            # 否则 WPS 跳过 UDF 扫描 =k() 报 #NAME?
            win = jsa_cm.find('window')
            if win is None:
                win = ET.Element('window', {'cursorpos': '0', 'actived': 'true', 'visible': 'true'})
                jsa_cm.insert(0, win)
            else:
                win.set('actived', 'true')
                win.set('visible', 'true')
            print('♻️  复用 codemodule JSA880 (id=3) — 强制 actived=true')

        # 写 JSA880 codemodule 的 codetext
        ct = jsa_cm.find('codetext')
        if ct is None:
            ct = ET.SubElement(jsa_cm, 'codetext')
        ct.text = combined

        # 3) 找/创建 ThisWorkbook(只放 workbook_open)
        thiswb = None
        for cm in root.findall('codemodule'):
            if cm.get('name') == 'ThisWorkbook':
                thiswb = cm
                break
        if thiswb is None:
            thiswb = ET.Element('codemodule', {'name': 'ThisWorkbook', 'id': '999'})
            ET.SubElement(thiswb, 'window', {'cursorpos': '0', 'actived': 'true', 'visible': 'false'})
            root.append(thiswb)
        else:
            thiswb.set('id', '999')
            win = thiswb.find('window')
            if win is None:
                win = ET.Element('window', {'cursorpos': '0', 'actived': 'true', 'visible': 'false'})
                thiswb.insert(0, win)
            # 清空所有子节点(window 之外)
            for child in list(thiswb):
                if child.tag != 'window':
                    thiswb.remove(child)
            print('♻️  复用 ThisWorkbook (id=999)')

        ct2 = thiswb.find('codetext')
        if ct2 is None:
            ct2 = ET.SubElement(thiswb, 'codetext')
        ct2.text = thiswb_code

        # 4) 改 activemodule 指向 JSA880 (id=3)
        am = root.find('activemodule')
        if am is None:
            am = ET.SubElement(root, 'activemodule')
        am.text = '3'
        print('🎯 activemodule → 3 (JSA880)')

        # 5) 写回
        tmp = Path(str(target_bin) + '.tmp')
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
    print('✅ 注入完成! v4.0.13 JSA880 作为独立 codemodule (id=3)')
    print('   function k / function jsaLambda 顶层定义未被重命名')
    print('   WPS UDF scanner 直接扫到 → 单元格 =k(...) 即可调用')


if __name__ == '__main__':
    main()
