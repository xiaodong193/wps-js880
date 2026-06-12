#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inject_jsa880_agg.py  v4.2.3
═══════════════════════════════════════════════════════════════
策略:沿用 v4.0.13 注入脚本结构(JSA880 作为独立 codemodule id=3)
针对 3.5 agg.xlsm: JSA880 模块替换 mJSA880-V1.6.1,保留 Module2 测试代码
修复合并: JSA.agg 从 fn 升级为 namespace(支持 count/sum/avg/min/max/textjoin/qctextjoin/addFunction)
═══════════════════════════════════════════════════════════════
"""
import sys
import shutil
import zipfile
import tempfile
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description='v4.2.3 注入: JSA880 → 3.5 agg.xlsm')
    parser.add_argument('--target', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'JS880教案/第03章/3-05/3.5 JSA二维数组按列求和计数平均值最大最小值聚合函数agg.xlsm')
    parser.add_argument('--source', default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                                             'js880/JSA880.js')
    args = parser.parse_args()

    target = Path(args.target).resolve()
    source_js = Path(args.source).resolve()

    print('═' * 60)
    print(' v4.2.3 注入 — JSA.agg namespace 升级到 3.5 agg.xlsm')
    print('═' * 60)
    print(f'  目标: {target.name}')
    print(f'  源:   {source_js.name}')

    # 备份
    backup = target.with_suffix(target.suffix + '.v423.bak')
    if not backup.exists():
        shutil.copy2(str(target), str(backup))
        print(f'💾 备份: {backup.name}')
    else:
        print(f'💾 备份已存在: {backup.name} (跳过)')

    jsa_code = source_js.read_text(encoding='utf-8')
    print(f'📖 JSA880.js: {len(jsa_code):,} 字符')

    # 顶部 banner
    banner = (
        "// ╔════════════════════════════════════════════════════════════╗\n"
        "// ║ JSA880 v4.2.3 — 2026-06-11                              ║\n"
        "// ║ 新增: JSA.agg = { count, sum, avg, min, max,             ║\n"
        "// ║               textjoin, qctextjoin, addFunction }       ║\n"
        "// ║ 修复: agg.addFunction(name, fn) + arr.<name>() 方法注册 ║\n"
        "// ╚════════════════════════════════════════════════════════════╝\n"
        "console.log('✅ JSA880 v4.2.3 loaded (JSA.agg namespace + addFunction)');\n\n"
    )
    combined = banner + jsa_code

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

        # 列出当前 codemodules
        print('📋 现有 codemodules:')
        for cm in root.findall('codemodule'):
            name = cm.get('name', '')
            mid = cm.get('id', '')
            ctt = cm.find('codetext')
            tlen = len(ctt.text) if ctt is not None and ctt.text else 0
            print(f'    {name} (id={mid}) {tlen:,} chars')

        # 1) 删旧的 mJSA880-* / JSA880 / kWrapper codemodules
        purged = []
        for cm in list(root.findall('codemodule')):
            nm = cm.get('name', '')
            if nm == 'JSA880' or nm.startswith('mJSA880') or nm == 'kWrapper':
                root.remove(cm)
                purged.append(nm)
        if purged:
            print(f'🧹 清理旧 codemodule: {purged}')

        # 2) 新建/复用 JSA880 codemodule (id=3)
        jsa_cm = None
        for cm in root.findall('codemodule'):
            if cm.get('name') == 'JSA880':
                jsa_cm = cm
                break
        if jsa_cm is None:
            jsa_cm = ET.Element('codemodule', {'name': 'JSA880', 'id': '1'})
            ET.SubElement(jsa_cm, 'window', {'cursorpos': '0', 'actived': 'true', 'visible': 'true'})
            inserted = False
            for i, elem in enumerate(list(root)):
                if elem.tag == 'codemodule' and elem.get('name') == 'Module2':
                    root.insert(i, jsa_cm)  # 插到 Module2 之前,确保先加载
                    inserted = True
                    break
            if not inserted:
                root.append(jsa_cm)
            print('🆕 新建 codemodule JSA880 (id=1) — 插在 Module2 之前')
        else:
            jsa_cm.set('id', '1')
            for child in list(jsa_cm):
                if child.tag != 'window':
                    jsa_cm.remove(child)
            win = jsa_cm.find('window')
            if win is None:
                win = ET.Element('window', {'cursorpos': '0', 'actived': 'true', 'visible': 'true'})
                jsa_cm.insert(0, win)
            else:
                win.set('actived', 'true')
                win.set('visible', 'true')
            print('♻️  复用 codemodule JSA880 (id=3)')

        # 写 codetext
        ct = jsa_cm.find('codetext')
        if ct is None:
            ct = ET.SubElement(jsa_cm, 'codetext')
        ct.text = combined
        print(f'📝 写入 JSA880 codetext: {len(combined):,} 字符')

        # 3) 找/创建 ThisWorkbook(放 workbook_open + UDF shim)
        thiswb_code = (
            "// ThisWorkbook — v4.2.3\n"
            "function Workbook_Open() {\n"
            "    try {\n"
            "        if (typeof Console !== 'undefined') Console.log('✅ JSA880 v4.2.3 ThisWorkbook loaded');\n"
            "    } catch (e) {}\n"
            "}\n"
        )
        thiswb = None
        for cm in root.findall('codemodule'):
            if cm.get('name') == 'ThisWorkbook':
                thiswb = cm
                break
        if thiswb is None:
            thiswb = ET.Element('codemodule', {'name': 'ThisWorkbook', 'id': '999'})
            ET.SubElement(thiswb, 'window', {'cursorpos': '0', 'actived': 'true', 'visible': 'false'})
            root.append(thiswb)
            print('🆕 新建 ThisWorkbook (id=999)')
        else:
            thiswb.set('id', '999')
            win = thiswb.find('window')
            if win is None:
                win = ET.Element('window', {'cursorpos': '0', 'actived': 'true', 'visible': 'false'})
                thiswb.insert(0, win)
            for child in list(thiswb):
                if child.tag != 'window':
                    thiswb.remove(child)
            print('♻️  复用 ThisWorkbook (id=999)')

        ct2 = thiswb.find('codetext')
        if ct2 is None:
            ct2 = ET.SubElement(thiswb, 'codetext')
        ct2.text = thiswb_code

        # 4) 改 activemodule 指向 JSA880
        am = root.find('activemodule')
        if am is None:
            am = ET.SubElement(root, 'activemodule')
        am.text = '1'
        print('🎯 activemodule → 1 (JSA880)')

        # 5) 写回 JDEData.bin
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
    print('✅ 注入完成! JSA.agg namespace 已就位')
    print('   JSA880 codemodule (id=3) + Module2 测试代码 (id=2) + ThisWorkbook (id=999)')
    print('   用户可执行 agg聚合函数test() 验证 count/sum/textjoin/qctextjoin')
    print('   或 平方和/aggStDev 验证 addFunction 注册')


if __name__ == '__main__':
    main()
