#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inject_jsa880.py - 把 JSA880 框架 + KO k 函数 UDF shim 注入到 xlsm 文件
═════════════════════════════════════════════════════════════════════
让任何 xlsm 工作簿打开后,单元格中直接可用 =k(...) 公式

WPS JSA 代码存储结构
─────────────────────
WPS JSA 把所有自定义代码存放在 xlsm 内的 `xl/JDEData.bin` 文件里,
这个文件是 XML 格式(WPS 专有),结构如下:

<document version="2.0">
  <name>Project</name>
  <property desc="" lock="false" password="" />
  <activemodule>8</activemodule>
  <codemodule name="m调息" id="21">
    <window cursorpos="0" actived="true" visible="false" />
    <codetext>/**&#x0A; * ...</codetext>   ← JavaScript 源码(HTML entity 编码)
  </codemodule>
  <codemodule name="mUndoManager" id="25">...</codemodule>
  ...
  <functionsdata />
</document>

注入策略
────────
- v4.8.5 是纯数据型 xlsm,没有 JDEData.bin
- v4.8.51 是带 JSA 的 xlsm,JDEData.bin 含完整 JSA880 框架
- 工具两种模式:
    ① copy-mode  : 完整复制 v4.8.51 的 JDEData.bin 到 v4.8.5
                   (含全部 JSA880 框架 + 业务模块)
    ② inject-mode: 读目标 JDEData.bin,在末尾追加 KO k 函数 shim 模块

使用方法
────────
    python3 inject_jsa880.py <target.xlsm> [--source <src.xlsm>] [--mode copy|inject] [--shim <shim.js>]

示例
────
    # 把 v4.8.51 的 JDEData.bin 完整复制到 v4.8.5 + 追加 KO k 函数 shim
    python3 inject_jsa880.py 收益测算表开发V4.8.5-wps版本.xlsm \
        --source 收益测算表开发V4.8.51-wps版本.xlsm \
        --shim KO一切的k函数_UDF模块.js
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

# 提前用 chr() 拿到实体字符,避免 Python 字符串语法坑
_AMP = chr(38)            # &
_LT = chr(60)             # <
_GT = chr(62)             # >
_QUOT = chr(34)           # "
_APOS = chr(39)           # '
_NEWLINE = chr(10)        # \n
_CR = chr(13)             # \r
_TAB = chr(9)             # \t

# ─────────────────────────────────────────────
#  XML 实体转义(逆向)
# ─────────────────────────────────────────────
_ENTITY_RE = re.compile(
    r'&(?:#x([0-9a-fA-F]+)|#(\d+)|(amp|lt|gt|quot|apos));'
)

def _decode_entity(m):
    if m.group(1):
        return chr(int(m.group(1), 16))
    if m.group(2):
        return chr(int(m.group(2)))
    name = m.group(3)
    return {'amp': _AMP, 'lt': _LT, 'gt': _GT, 'quot': _QUOT, 'apos': _APOS}[name]

def decode_entities(s):
    return _ENTITY_RE.sub(_decode_entity, s)


# 编码(写入 codetext 前)
def encode_for_xml(s):
    out = []
    for ch in s:
        if ch == _AMP:   out.append(_AMP + 'amp;')
        elif ch == _LT:  out.append(_LT + 'lt;')
        elif ch == _GT:  out.append(_GT + 'gt;')
        elif ch == _QUOT: out.append(_QUOT + 'quot;')
        elif ch == _APOS: out.append(_APOS + 'apos;')
        elif ch == _NEWLINE: out.append('&#x0A;')
        elif ch == _CR:  out.append('&#x0D;')
        elif ch == _TAB: out.append('&#x09;')
        else:
            cp = ord(ch)
            if cp < 0x20:
                out.append('&#x' + format(cp, 'x') + ';')
            else:
                out.append(ch)
    return ''.join(out)


# ─────────────────────────────────────────────
#  JDEData.bin 读取 / 写回
# ─────────────────────────────────────────────
def load_jdedata_bin(bin_path):
    tree = ET.parse(str(bin_path))
    root = tree.getroot()
    modules = root.findall('codemodule')
    max_id = 0
    for cm in modules:
        cid = int(cm.get('id', '0'))
        if cid > max_id: max_id = cid
    return tree, root, modules, max_id


def make_codemodule(name, cid, code):
    cm = ET.Element('codemodule', {'name': name, 'id': str(cid)})
    ET.SubElement(cm, 'window', {
        'cursorpos': '0', 'actived': 'true', 'visible': 'false'
    })
    text = ET.SubElement(cm, 'codetext')
    text.text = encode_for_xml(code)
    return cm


def write_jdedata_bin(root, out_path):
    tree = ET.ElementTree(root)
    ET.indent(tree, space='    ', level=0)
    tree.write(str(out_path), encoding='UTF-8',
               xml_declaration=True, short_empty_elements=True)


# ─────────────────────────────────────────────
#  核心:追加 shim codemodule
# ─────────────────────────────────────────────
def inject_shim(target_bin, shim_js_path, target_module_name='KO_k函数'):
    tree, root, modules, max_id = load_jdedata_bin(target_bin)
    print('  [JDEData.bin] 已有 {} 个 codemodule,最大 id={}'.format(len(modules), max_id))

    target_name = target_module_name
    existing = None
    for cm in modules:
        if cm.get('name') == target_name:
            existing = cm
            break

    shim_code = shim_js_path.read_text(encoding='utf-8')

    new_id = (existing.get('id') if existing is not None else None) or str(max_id + 1)
    new_cm = make_codemodule(target_name, int(new_id), shim_code)

    if existing is not None:
        idx = list(root).index(existing)
        root.remove(existing)
        root.insert(idx, new_cm)
        print('  [替换] codemodule {} (id={})'.format(target_name, new_id))
    else:
        funcdata = root.find('functionsdata')
        if funcdata is not None:
            idx = list(root).index(funcdata)
            root.insert(idx, new_cm)
        else:
            root.append(new_cm)
        print('  [追加] codemodule {} (id={})'.format(target_name, new_id))

    tmp = target_bin.with_suffix('.tmp')
    write_jdedata_bin(root, tmp)
    shutil.move(str(tmp), str(target_bin))
    print('  [写回] {} ({} bytes)'.format(target_bin, target_bin.stat().st_size))


# ─────────────────────────────────────────────
#  完整复制 JDEData.bin (copy mode)
# ─────────────────────────────────────────────
def copy_jdedata_bin(source_xlsm, target_xlsm):
    src_bin = Path('/tmp/_src_xlsm_extract') / 'xl' / 'JDEData.bin'
    if src_bin.exists():
        shutil.rmtree('/tmp/_src_xlsm_extract', ignore_errors=True)
    src_bin.parent.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(str(source_xlsm)) as zf:
        if 'xl/JDEData.bin' not in zf.namelist():
            raise RuntimeError('源 xlsm 没有 xl/JDEData.bin: ' + str(source_xlsm))
        src_bin.parent.mkdir(parents=True, exist_ok=True)
        with zf.open('xl/JDEData.bin') as src, open(src_bin, 'wb') as dst:
            shutil.copyfileobj(src, dst)
    print('  [copy] 从 {} 提取 JDEData.bin: {} bytes'.format(
        source_xlsm.name, src_bin.stat().st_size))

    with zipfile.ZipFile(str(target_xlsm)) as zf:
        has_bin = 'xl/JDEData.bin' in zf.namelist()
        members = zf.namelist()
    print('  [check] {}: 现有 {} 个文件,{} JDEData.bin'.format(
        target_xlsm.name, len(members), '已有' if has_bin else '没有'))

    return src_bin


# ─────────────────────────────────────────────
#  把文件注入到 xlsm zip
# ─────────────────────────────────────────────
def add_file_to_xlsm(xlsm_path, internal_path, source_file):
    tmp_out = xlsm_path.with_suffix('.inject.tmp')
    with zipfile.ZipFile(str(xlsm_path), 'r') as zin:
        with zipfile.ZipFile(str(tmp_out), 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == internal_path:
                    continue
                zout.writestr(item, zin.read(item.filename))
            with open(source_file, 'rb') as f:
                zout.writestr(internal_path, f.read())
    shutil.move(str(tmp_out), str(xlsm_path))
    print('  [add] {} :: {} 添加完成'.format(xlsm_path, internal_path))


# ─────────────────────────────────────────────
#  主流程
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='把 JSA880 框架 + KO k 函数 UDF shim 注入到 xlsm',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument('target', help='目标 xlsm 文件路径(要注入的那个)')
    parser.add_argument('--source', help='源 xlsm 文件路径(用来取 JDEData.bin,仅 copy 模式需要)')
    parser.add_argument('--mode', choices=['copy', 'inject', 'copy+inject'],
                        default='copy+inject',
                        help='注入模式: copy=复制整个 JDEData.bin; '
                             'inject=在已有 JDEData.bin 末尾追加 shim; '
                             'copy+inject=两个都做(默认)')
    parser.add_argument('--shim',
                        default=str(Path(__file__).parent / 'KO一切的k函数_UDF模块.js'),
                        help='KO k 函数 shim JS 文件路径(默认: 同目录下的 KO一切的k函数_UDF模块.js)')
    parser.add_argument('--module-name', default='KO_k函数',
                        help='注入到 JDEData.bin 里的 codemodule 名字(默认 KO_k函数)')
    parser.add_argument('--no-backup', action='store_true',
                        help='覆盖前不自动备份原文件')
    args = parser.parse_args()

    target = Path(args.target).resolve()
    if not target.exists():
        print('❌ 目标 xlsm 不存在: ' + str(target))
        sys.exit(1)
    shim = Path(args.shim).resolve()
    if not shim.exists():
        print('❌ shim 文件不存在: ' + str(shim))
        sys.exit(1)

    print('📦 目标 xlsm: ' + str(target))
    print('📦 Shim JS:   ' + str(shim))
    if args.source:
        src = Path(args.source).resolve()
        if not src.exists():
            print('❌ 源 xlsm 不存在: ' + str(src))
            sys.exit(1)
        print('📦 源 xlsm:   ' + str(src))
    print('🔧 模式:      ' + args.mode)
    print()

    # 1) 备份
    if not args.no_backup and not str(target).endswith('.bak'):
        backup_path = target.with_suffix(target.suffix + '.bak')
        if not backup_path.exists():
            shutil.copy2(str(target), str(backup_path))
            print('💾 备份: ' + str(backup_path))
    print()

    # 2) copy mode
    if 'copy' in args.mode:
        if not args.source:
            print('❌ --copy 模式需要指定 --source 源 xlsm')
            sys.exit(1)
        src_bin = copy_jdedata_bin(src, target)
        add_file_to_xlsm(target, 'xl/JDEData.bin', src_bin)
        print()

    # 3) inject mode
    if 'inject' in args.mode:
        with tempfile.TemporaryDirectory() as tmpdir:
            workdir = Path(tmpdir)
            with zipfile.ZipFile(str(target)) as zf:
                if 'xl/JDEData.bin' not in zf.namelist():
                    print('❌ 目标 xlsm 里没有 xl/JDEData.bin,无法 inject。请先 --mode copy+inject 或 --mode copy')
                    sys.exit(1)
                target_bin = workdir / 'JDEData.bin'
                with zf.open('xl/JDEData.bin') as src, open(target_bin, 'wb') as dst:
                    shutil.copyfileobj(src, dst)

            print('📋 inject shim 到 ' + target.name + ':JDEData.bin')
            inject_shim(target_bin, shim, args.module_name)
            add_file_to_xlsm(target, 'xl/JDEData.bin', target_bin)
        print()

    print('✅ 全部完成!')
    print('   接下来: 用 WPS 打开 ' + target.name)
    print('   Console 会看到 "✅ k() UDF 已就绪!"')


if __name__ == '__main__':
    main()
