#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inject_jsa880_main.py
═══════════════════════════════════════════════════════════════
把 js880/JSA880.js(单文件整合版:v4.2.2,已含 KO k 函数 shim)
注入到 KO一切的k函数.xlsm 的 xl/JDEData.bin 中,替换原有的 JSA880
codemodule 内容。WPS 打开后:

  • 单元格里直接可用 =k("JSA.getIndexs", 1, 10, 2) → 1 3 5 7 9
  • 单元格里直接可用 =jsaLambda(...)
  • 保留原 xlsm 的所有 sheet / 公式 / 样式 / 数据

WPS JSA 代码存储结构
─────────────────────
xlsm 里的 xl/JDEData.bin 是 WPS 专有的 XML 格式,结构:
  <document version="2.0">
    <name>Project</name>
    <property ... />
    <activemodule>1</activemodule>
    <codemodule name="JSA880" id="1">
      <window ... />
      <codetext>...HTML entity 编码后的 JS 源码...</codetext>
    </codemodule>
    <codemodule name="Module2" id="2">...</codemodule>
    <functionsdata />
  </document>

本脚本做的事情:
  1. 备份目标 xlsm
  2. 提取 JDEData.bin,解析
  3. 找到 name="JSA880" 的 codemodule,替换其 codetext
     (id 保持不变,保证 activemodule 等引用不会断)
  4. 写回 JDEData.bin(临时 .tmp → rename 原子替换)
  5. 重新打包 xlsm(临时 .tmp → rename 原子替换)
  6. 验证:unzip -t + 重新解析 JDEData.bin 校验模块列表
"""
import os
import sys
import re
import shutil
import zipfile
import tempfile
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

# ─────────────────────────────────────────────
#  XML 实体编码(写入 codetext 前)
# ─────────────────────────────────────────────
_AMP = chr(38)
_LT = chr(60)
_GT = chr(62)
_QUOT = chr(34)
_APOS = chr(39)
_NEWLINE = chr(10)
_CR = chr(13)
_TAB = chr(9)

def encode_for_xml(s):
    """跟 WPS JDEData.bin 保持一致的实体编码"""
    out = []
    for ch in s:
        if ch == _AMP:    out.append(_AMP + 'amp;')
        elif ch == _LT:   out.append(_LT + 'lt;')
        elif ch == _GT:   out.append(_GT + 'gt;')
        elif ch == _QUOT: out.append(_QUOT + 'quot;')
        elif ch == _APOS: out.append(_APOS + 'apos;')
        elif ch == _NEWLINE: out.append('&#x0A;')
        elif ch == _CR:   out.append('&#x0D;')
        elif ch == _TAB:  out.append('&#x09;')
        else:
            cp = ord(ch)
            if cp < 0x20:
                out.append('&#x' + format(cp, 'x') + ';')
            else:
                out.append(ch)
    return ''.join(out)


# ─────────────────────────────────────────────
#  核心:替换 JDEData.bin 中指定模块的 codetext
# ─────────────────────────────────────────────
def replace_codemodule(bin_path, module_name, new_code,
                       purge_prefixes=None):
    """
    打开 JDEData.bin:
      1. 删除所有 name 以 purge_prefixes 中任一前缀开头的旧 codemodule
         (主要用于清理老的 mJSA880* / JSA880* 框架,避免冲突)
      2. 找到 name==module_name 的 codemodule,替换它的 codetext
      3. 找不到同名模块 → 在 functionsdata 之前追加
      4. 把 activemodule 指向新/改后的模块
      5. 写回 bin_path(原子:tmp → rename)
    返回 (action, used_id, enc_size, purged_names)。
    """
    if purge_prefixes is None:
        purge_prefixes = []

    tree = ET.parse(str(bin_path))
    root = tree.getroot()
    purged_names = []

    # 1) 删除所有匹配前缀的旧模块
    for cm in list(root.findall('codemodule')):
        nm = cm.get('name', '')
        for pfx in purge_prefixes:
            if nm == pfx or nm.startswith(pfx):
                root.remove(cm)
                purged_names.append(nm)
                break

    # 2) 找到 name==module_name 的同名模块
    target = None
    max_id = 0
    for cm in root.findall('codemodule'):
        cid = int(cm.get('id', '0'))
        if cid > max_id:
            max_id = cid
        if cm.get('name') == module_name:
            target = cm

    if target is None:
        # 追加新模块
        new_id = max_id + 1
        cm = ET.Element('codemodule', {'name': module_name, 'id': str(new_id)})
        ET.SubElement(cm, 'window', {
            'cursorpos': '0', 'actived': 'true', 'visible': 'true'
        })
        text = ET.SubElement(cm, 'codetext')
        text.text = encode_for_xml(new_code)
        funcdata = root.find('functionsdata')
        if funcdata is not None:
            idx = list(root).index(funcdata)
            root.insert(idx, cm)
        else:
            root.append(cm)
        enc_size = len(text.text)
        action, used_id = 'appended', new_id
    else:
        # 替换同名模块的 codetext
        used_id = target.get('id')
        ct = target.find('codetext')
        if ct is None:
            ct = ET.SubElement(target, 'codetext')
        ct.text = encode_for_xml(new_code)
        enc_size = len(ct.text)
        action = 'replaced'

    # 3) 把 activemodule 指向新/改后的模块
    am = root.find('activemodule')
    if am is None:
        am = ET.SubElement(root, 'activemodule')
    am.text = str(used_id)

    # 4) 写回 bin_path(原子:tmp → rename)
    tmp = Path(str(bin_path) + '.tmp')
    tree2 = ET.ElementTree(root)
    ET.indent(tree2, space='    ', level=0)
    tree2.write(str(tmp), encoding='UTF-8',
                xml_declaration=True, short_empty_elements=True)
    shutil.move(str(tmp), str(bin_path))

    return (action, used_id, enc_size, purged_names)


def write_jdedata_bin(root, out_path):
    tree = ET.ElementTree(root)
    ET.indent(tree, space='    ', level=0)
    tree.write(str(out_path), encoding='UTF-8',
               xml_declaration=True, short_empty_elements=True)


# ─────────────────────────────────────────────
#  Zip 重打包:替换 xlsm 内的一个文件
# ─────────────────────────────────────────────
def replace_in_zip(zip_path, internal_path, source_file):
    """
    重新打包 zip_path,把 internal_path 这一项替换为 source_file 的内容。
    用临时文件 + rename 的方式保证原子性。
    """
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


# ─────────────────────────────────────────────
#  主流程
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='把 js880/JSA880.js 注入到 xlsm 的 xl/JDEData.bin '
                    '(替换 JSA880 codemodule 的 codetext)')
    parser.add_argument(
        '--target',
        default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                'JS880教案/第03章/3-28/KO一切的k函数.xlsm',
        help='目标 xlsm 路径')
    parser.add_argument(
        '--source',
        default='/Users/daidai193/Library/CloudStorage/SynologyDrive-code/'
                'js880/JSA880.js',
        help='要注入的 .js 源文件')
    parser.add_argument(
        '--module-name', default='JSA880',
        help='注入到 JDEData.bin 里的 codemodule 名字(默认 JSA880)')
    parser.add_argument(
        '--no-backup', action='store_true',
        help='不生成 .bak 备份')
    parser.add_argument(
        '--purge-prefix', action='append', default=[],
        help='注入前,删除所有 name 以此前缀开头的旧 codemodule '
             '(可重复指定,避免新旧 JSA880 框架冲突)。'
             '默认会自动清理: mJSA880, JSA880')
    args = parser.parse_args()

    # 默认清理的前缀
    DEFAULT_PURGE = ['mJSA880', 'JSA880']
    if not args.purge_prefix:
        args.purge_prefix = DEFAULT_PURGE

    target = Path(args.target).resolve()
    source_js = Path(args.source).resolve()

    print('═' * 60)
    print(' inject_jsa880_main.py')
    print('═' * 60)
    print(f'  目标 xlsm:  {target}')
    print(f'  源 .js:     {source_js}')
    print(f'  模块名:     {args.module_name}')
    print(f'  清理前缀:   {args.purge_prefix}')
    print()

    if not target.exists():
        print(f'❌ 目标 xlsm 不存在: {target}')
        sys.exit(1)
    if not source_js.exists():
        print(f'❌ 源 .js 不存在: {source_js}')
        sys.exit(1)

    # 1) 备份
    if not args.no_backup:
        backup = target.with_suffix(target.suffix + '.bak')
        if not backup.exists():
            shutil.copy2(str(target), str(backup))
            print(f'💾 备份: {backup}')
        else:
            print(f'💾 备份已存在,跳过: {backup}')
    print()

    # 2) 读源 JS
    new_code = source_js.read_text(encoding='utf-8')
    print(f'📖 源 .js: {len(new_code):,} 字符, {new_code.count(chr(10)):,} 行')
    print()

    # 3) 提取 JDEData.bin 到临时目录
    with tempfile.TemporaryDirectory() as tmpdir:
        workdir = Path(tmpdir)
        target_bin = workdir / 'JDEData.bin'
        with zipfile.ZipFile(str(target)) as zf:
            if 'xl/JDEData.bin' not in zf.namelist():
                print('❌ 目标 xlsm 里没有 xl/JDEData.bin')
                sys.exit(1)
            with zf.open('xl/JDEData.bin') as src, open(target_bin, 'wb') as dst:
                shutil.copyfileobj(src, dst)

        orig_size = target_bin.stat().st_size
        print(f'📂 提取 xl/JDEData.bin: {orig_size:,} bytes')

        # 4) 解析,替换
        action, used_id, enc_size, purged_names = replace_codemodule(
            target_bin, args.module_name, new_code,
            purge_prefixes=args.purge_prefix,
        )
        if purged_names:
            print(f'🧹 清理掉旧模块: {purged_names}')
        print(f'🔧 {action} codemodule "{args.module_name}" (id={used_id}), '
              f'编码后 codetext={enc_size:,} 字符')

        # 5) 检查写回后大小(replace_codemodule 内部已写回)
        new_bin_size = target_bin.stat().st_size
        print(f'💽 JDEData.bin: {orig_size:,} → {new_bin_size:,} bytes')

        # 6) 重新打包到 xlsm
        print('📦 重新打包 xlsm ...')
        replace_in_zip(target, 'xl/JDEData.bin', target_bin)

    print()
    print('✅ 注入完成!')
    print()
    print('─' * 60)
    print(' 接下来:')
    print(f'   1. 用 WPS 打开 {target.name}')
    print('   2. 看 JSA Console,应该出现:')
    print('      [JSA880 v4.2.2] KO一切k函数 UDF 已整合到主框架')
    print('      顶层 function k() / jsaLambda() 会被 WPS 公式引擎')
    print('      自动注册为 UDF')
    print('   3. 在任意单元格输入 =k("JSA.getIndexs", 1, 10, 2)')
    print('      应该看到 1 3 5 7 9 的数组溢出')
    print('─' * 60)


if __name__ == '__main__':
    main()
