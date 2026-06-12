#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inject_jsa880_text.py
════════════════════════════════════════════════════════════════════
v2 纯文本注入器 — 不解析 XML,只拼接字符串,完全保留 WPS 原生格式

问题背景:
  inject_jsa880_main.py(v1)用 ElementTree + ET.indent 重新格式化
  整个 JDEData.bin。WPS JSA 编辑器对格式敏感(单/双引号、缩进、
  functionsdata 自闭合等),re-format 后 WPS 加载时只把 codemodule
  显示成"1 行"(用户实测反馈),UDF 注册失败。

本工具策略:
  • 完全不解析 JDEData.bin 的 XML
  • 只在 <functionsdata /> 之前,用纯字符串拼接追加新 codemodule
  • 完全保留原 .bak 里的 XML 头、document 标签、缩进、已有模块
  • 用纯文本正则删掉旧模块(mJSA880-V1.3.8)的整段块

用法:
  python3 inject_jsa880_text.py <target.xlsm> \
      --source1 js880/JSA880.js --name1 JSA880 \
      --source2 KO_k_udf.js --name2 KO_k_udf
"""
import os
import re
import sys
import shutil
import zipfile
import tempfile
import argparse
from pathlib import Path

# ─────────────────────────────────────────────
#  XML 实体编码(写入 codetext 前)
# ─────────────────────────────────────────────
def encode_for_xml(s):
    """跟 WPS JDEData.bin 保持一致的实体编码"""
    out = []
    for ch in s:
        cp = ord(ch)
        if ch == '&':   out.append('&amp;')
        elif ch == '<':  out.append('&lt;')
        elif ch == '>':  out.append('&gt;')
        elif ch == '"':  out.append('&quot;')
        elif ch == "'":  out.append('&apos;')
        elif ch == '\n': out.append('&#x0A;')
        elif ch == '\r': out.append('&#x0D;')
        elif ch == '\t': out.append('&#x09;')
        elif cp < 0x20:  out.append('&#x' + format(cp, 'x') + ';')
        else:            out.append(ch)
    return ''.join(out)


def make_codemodule(name, cid, code):
    """生成一个 codemodule XML 块(格式跟 .bak 原生一致:2 空格缩进)"""
    enc = encode_for_xml(code)
    return (
        f'    <codemodule name="{name}" id="{cid}">\n'
        f'        <window cursorpos="0" actived="true" visible="true" />\n'
        f'        <codetext>{enc}</codetext>\n'
        f'    </codemodule>\n'
    )


# ─────────────────────────────────────────────
#  纯文本注入:删 + 追加
# ─────────────────────────────────────────────
def inject_modules(bin_text, modules_to_inject, purge_prefixes=None):
    """
    在 JDEData.bin 文本上:
      1. 删掉所有 name 以 purge_prefixes 中任一前缀开头的旧 codemodule
         (含 </codemodule> 闭标签前后的换行)
      2. 在 <functionsdata /> 之前追加新 codemodule 块

    返回 (new_text, purged_names)。
    """
    if purge_prefixes is None:
        purge_prefixes = ['mJSA880', 'JSA880']

    purged = []

    # 1) 删旧 codemodule 块(从 <codemodule ...> 到 </codemodule> 整段 + 前后换行)
    def delete_old(text):
        # 多次扫描,直到没有匹配为止
        while True:
            matches = list(re.finditer(r'<codemodule\s+name="([^"]+)"', text))
            matched = False
            for m in matches:
                nm = m.group(1)
                if not any(nm == pfx or nm.startswith(pfx) for pfx in purge_prefixes):
                    continue
                # 找匹配的 </codemodule>
                end_m = re.search(r'</codemodule>', text[m.start():])
                if not end_m:
                    continue
                block_start = m.start()
                # 向前回溯一个换行
                if block_start > 0 and text[block_start - 1] == '\n':
                    block_start -= 1
                block_end = m.start() + end_m.end()
                # 向后吃一个换行
                if block_end < len(text) and text[block_end] == '\n':
                    block_end += 1
                purged.append(nm)
                text = text[:block_start] + text[block_end:]
                matched = True
                break
            if not matched:
                break
        return text

    bin_text = delete_old(bin_text)

    # 2) 计算新 id
    existing_ids = [int(m) for m in re.findall(r'<codemodule\s+name="[^"]+"\s+id="(\d+)"', bin_text)]
    next_id = max(existing_ids, default=0) + 1

    # 3) 在 <functionsdata /> 之前插入新模块
    insertion = ''
    for (name, code) in modules_to_inject:
        insertion += make_codemodule(name, next_id, code)
        next_id += 1

    funcdata_match = re.search(r'<functionsdata\s*/>', bin_text)
    if funcdata_match:
        bin_text = bin_text[:funcdata_match.start()] + insertion + bin_text[funcdata_match.start():]
    else:
        # 没有 <functionsdata />,在 </document> 之前插入
        bin_text = bin_text.replace('</document>', insertion + '</document>')

    return bin_text, purged


# ─────────────────────────────────────────────
#  Zip 重打包
# ─────────────────────────────────────────────
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


# ─────────────────────────────────────────────
#  主流程
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='纯文本方式注入 codemodule 到 JDEData.bin,不重排已有格式',
        formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('target', help='目标 xlsm')
    parser.add_argument('--source1', required=True, help='第一个源 .js(主框架,如 JSA880.js)')
    parser.add_argument('--name1', required=True, help='第一个模块名(默认 JSA880)')
    parser.add_argument('--source2', help='第二个源 .js(可选,兜底层)')
    parser.add_argument('--name2', help='第二个模块名(默认 KO_k_udf)')
    parser.add_argument('--no-backup', action='store_true')
    args = parser.parse_args()

    target = Path(args.target).resolve()
    if not target.exists():
        print(f'❌ 目标不存在: {target}'); sys.exit(1)

    print('═' * 60)
    print(' inject_jsa880_text.py(v2 纯文本注入)')
    print('═' * 60)
    print(f'  目标:    {target}')

    # 1) 备份
    if not args.no_backup:
        backup = target.with_suffix(target.suffix + '.bak')
        if not backup.exists():
            shutil.copy2(str(target), str(backup))
            print(f'💾 备份: {backup}')
        else:
            print(f'💾 备份已存在: {backup}')

    # 2) 读源
    modules_to_inject = []
    src1 = Path(args.source1)
    if not src1.exists():
        print(f'❌ 源 1 不存在: {src1}'); sys.exit(1)
    code1 = src1.read_text(encoding='utf-8')
    modules_to_inject.append((args.name1, code1))
    print(f'📖 源 1: {src1.name} ({len(code1):,} 字符) → 模块名 "{args.name1}"')

    if args.source2:
        src2 = Path(args.source2)
        if not src2.exists():
            print(f'❌ 源 2 不存在: {src2}'); sys.exit(1)
        code2 = src2.read_text(encoding='utf-8')
        name2 = args.name2 or 'KO_k_udf'
        modules_to_inject.append((name2, code2))
        print(f'📖 源 2: {src2.name} ({len(code2):,} 字符) → 模块名 "{name2}"')

    # 3) 提取 JDEData.bin 到临时文件,纯文本修改,再写回
    with tempfile.TemporaryDirectory() as tmpdir:
        workdir = Path(tmpdir)
        bin_path = workdir / 'JDEData.bin'
        with zipfile.ZipFile(str(target)) as zf:
            if 'xl/JDEData.bin' not in zf.namelist():
                print('❌ 目标 xlsm 里没有 xl/JDEData.bin'); sys.exit(1)
            with zf.open('xl/JDEData.bin') as src, open(bin_path, 'wb') as dst:
                shutil.copyfileobj(src, dst)

        orig_size = bin_path.stat().st_size
        orig_text = bin_path.read_text(encoding='utf-8')
        print(f'📂 提取 JDEData.bin: {orig_size:,} bytes')

        # 注入
        new_text, purged = inject_modules(orig_text, modules_to_inject)
        if purged:
            print(f'🧹 清理掉旧模块: {purged}')

        # 写回
        bin_path.write_text(new_text, encoding='utf-8')
        new_size = bin_path.stat().st_size
        print(f'💽 JDEData.bin: {orig_size:,} → {new_size:,} bytes')

        # 重新打包
        print('📦 重新打包 xlsm ...')
        replace_in_zip(target, 'xl/JDEData.bin', bin_path)

    print()
    print('✅ 注入完成!')
    print()
    print('─' * 60)
    print(' 接下来:')
    print(f'   1. 完全关闭 WPS(包括右下角托盘)')
    print(f'   2. 重新打开 {target.name}')
    print('   3. 公式栏输入 =k("JSA.getIndexs", 1, 10, 2)')
    print('   4. 应该看到 1 3 5 7 9 数组溢出')
    print('─' * 60)


if __name__ == '__main__':
    main()
