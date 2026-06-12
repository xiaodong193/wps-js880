#!/usr/bin/env python3
"""列出 xlsm 里 sharedStrings 关键内容（标题、表头等）"""
import zipfile
from xml.etree import ElementTree as ET

z = zipfile.ZipFile('KO一切的k函数.xlsm')
ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

ss = ET.fromstring(z.read('xl/sharedStrings.xml').decode('utf-8'))
for idx, si in enumerate(ss.findall('main:si', ns)):
    t = si.find('main:t', ns)
    if t is not None and t.text:
        txt = t.text[:100]
        if txt.strip():
            print(f"  [{idx:3d}] {txt!r}")
