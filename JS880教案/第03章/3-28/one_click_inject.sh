#!/bin/bash
# ============================================================
# 一键注入:KO_k_udf(带 $$ 修复 + 错误诊断)到 KO一切的k函数.xlsm
# 用法:bash one_click_inject.sh
# 前置:必须先关掉 WPS(包括右下角托盘)
# ============================================================
set -e
cd "$(dirname "$0")"
XL="KO一切的k函数.xlsm"
SRC="KO_k_udf.js"

echo "━━━ 步骤 1/4:检查文件 ━━━"
ls -la "$XL" "$SRC" 2>&1 | tail -2

echo ""
echo "━━━ 步骤 2/4:Python 注入(只动 JDEData.bin,保留其他) ━━━"
python3 - <<'PY'
import os, re, shutil, zipfile, tempfile
from pathlib import Path

def encode(s):
    o = []
    for c in s:
        cp = ord(c)
        if   c == '&':  o.append('&')
        elif c == '<':  o.append('<')
        elif c == '>':  o.append('>')
        elif cp == 34:  o.append('"')
        elif cp == 39:  o.append(''')
        elif c == '\n': o.append('&#x0A;')
        elif c == '\r': o.append('&#x0D;')
        elif c == '\t': o.append('&#x09;')
        elif cp < 0x20: o.append(f'&#x{cp:x};')
        else:           o.append(c)
    return ''.join(o)

def make(name, cid, code):
    return f'    <codemodule name="{name}" id="{cid}">\n' \
           f'        <window cursorpos="0" actived="true" visible="true" />\n' \
           f'        <codetext>{encode(code)}</codetext>\n' \
           f'    </codemodule>\n'

target = Path('KO一切的k函数.xlsm').resolve()
src    = Path('KO_k_udf.js').read_text(encoding='utf-8')
print(f'  KO_k_udf.js 字符数: {len(src):,}')

with tempfile.TemporaryDirectory() as tmp:
    b = Path(tmp) / 'JDEData.bin'
    with zipfile.ZipFile(target) as z:
        with z.open('xl/JDEData.bin') as i, open(b, 'wb') as o:
            shutil.copyfileobj(i, o)
    txt = b.read_text(encoding='utf-8')
    print(f'  提取 JDEData.bin: {b.stat().st_size:,} bytes')

    # 清理所有 KO_k_udf 同名模块
    while True:
        m = re.search(r'<codemodule\s+name="KO_k_udf".*?</codemodule>', txt, re.DOTALL)
        if not m: break
        a, z_ = m.start(), m.end()
        if a > 0 and txt[a-1] == '\n': a -= 1
        if z_ < len(txt) and txt[z_] == '\n': z_ += 1
        txt = txt[:a] + txt[z_:]
    print('  ✓ 清理旧 KO_k_udf')

    # 算下一个 id
    ids = [int(m) for m in re.findall(r'<codemodule[^>]+id="(\d+)"', txt)]
    nxt = max(ids, default=0) + 1
    print(f'  下一个 id: {nxt}')

    # 插入新 KO_k_udf
    ins = make('KO_k_udf', nxt, src)
    fd = re.search(r'<functionsdata\s*/>', txt)
    if fd:
        txt = txt[:fd.start()] + ins + txt[fd.start():]
    else:
        txt = txt.replace('</document>', ins + '</document>')

    b.write_text(txt, encoding='utf-8')
    print(f'  写入 JDEData.bin: {b.stat().st_size:,} bytes')

    # 重打包 xlsm
    tmp_out = target.with_suffix('.tmp.xlsm')
    with zipfile.ZipFile(target, 'r') as zin, zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'xl/JDEData.bin': continue
            zout.writestr(item, zin.read(item.filename))
        zout.writestr('xl/JDEData.bin', b.read_bytes())
    shutil.move(str(tmp_out), str(target))
    print('  ✓ 重打包完成')

print(f'  目标: {target}')
print('  大小:', os.path.getsize(target), 'bytes')
PY

echo ""
echo "━━━ 步骤 3/4:验证注入 ━━━"
unzip -p "$XL" "xl/JDEData.bin" 2>/dev/null | grep -oE '<codemodule name="[^"]+"' | head -10
echo "---"
echo "KO_k_udf 错误诊断增强版检测:"
unzip -p "$XL" "xl/JDEData.bin" 2>/dev/null | grep -oE 'jsaLambda 返回 null' | head -1 && echo "  ✓ 带错误诊断版" || echo "  ✗ 错误诊断版未注入"
unzip -p "$XL" "xl/JDEData.bin" 2>/dev/null | grep -oE 'v4.2.2 bug' | head -1 && echo "  ✓ 带 \$\$ 修复" || echo "  ✗ \$\$ 修复未注入"

echo ""
echo "━━━ 步骤 4/4:完成 ━━━"
echo "  ✅ 注入完成。请:"
echo "     1. 完全关闭 WPS(托盘也要关)"
echo "     2. 重新打开 $XL"
echo "     3. 在 Sheet1!N1 输入: =k(123)"
echo "     4. 在 Sheet1!N2 输入: =k(\"JSA.getIndexs\", 1, 10, 2)"
echo "        → 应该看到 1 3 5 7 9 数组溢出"
echo "     5. 在 Sheet1!N3 输入: =k(\"\$\$.superPivot\", A1:H40, \"f3,f2\", \"f6\", \"sum(\\\"f4*f5\\\")\")"
echo "        → 这次会用双引号不用反引号,应该看到透视表"
echo ""
echo "  如果还有问题,看 JSA Console(开发工具→JSA 编辑器→Console)的输出告诉我"
