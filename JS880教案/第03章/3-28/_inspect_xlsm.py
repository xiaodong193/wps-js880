#!/usr/bin/env python3
"""列出 xlsm 里所有 k() / jsaLambda() 公式"""
import zipfile
from xml.etree import ElementTree as ET
import sys

z = zipfile.ZipFile('KO一切的k函数.xlsm')
ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

# 读 sharedStrings
ss = ET.fromstring(z.read('xl/sharedStrings.xml').decode('utf-8'))
shared = []
for si in ss.findall('main:si', ns):
    t = si.find('main:t', ns)
    shared.append(t.text if t is not None and t.text is not None else '')

# 解析每个 sheet 的公式
for i in range(1, 11):
    name = f'xl/worksheets/sheet{i}.xml'
    try:
        data = z.read(name).decode('utf-8')
    except KeyError:
        continue
    # 截断长 sheet 防止解析卡死
    if len(data) > 200000:
        data = data[:200000]
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        print(f"sheet{i}: parse error, skip")
        continue
    formulas = []
    for c in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
        f = c.find('main:f', ns)
        if f is not None and f.text:
            ref = c.get('r')
            ft = f.text
            ft_low = ft.lower()
            if 'k(' in ft_low or 'jsa' in ft_low or 'lambda' in ft_low:
                formulas.append((ref, ft[:400]))
    print(f'\n===== sheet{i}.xml ({len(formulas)} k/jsa 公式) =====')
    for ref, ft in formulas:
        print(f'  {ref}: {ft}')
