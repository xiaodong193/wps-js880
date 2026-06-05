#!/usr/bin/env python3
"""分析WPS Excel文件的结构和VBA/JSA代码"""
import zipfile
import xml.etree.ElementTree as ET
import os

def analyze_xlsm(file_path):
    """分析.xlsm文件结构"""
    print(f"=== 分析文件: {os.path.basename(file_path)} ===\n")

    if not os.path.exists(file_path):
        print(f"文件不存在: {file_path}")
        return

    with zipfile.ZipFile(file_path, 'r') as zf:
        print("【1. 文件内部结构】")
        all_files = zf.namelist()
        vba_files = [f for f in all_files if 'vba' in f.lower() or 'module' in f.lower() or 'vbaProject' in f.lower()]
        sheet_files = [f for f in all_files if f.startswith('xl/worksheets/')]
        print(f"  - 总文件数: {len(all_files)}")
        print(f"  - VBA相关文件: {len(vba_files)}")
        print(f"  - 工作表文件: {len(sheet_files)}")

        print("\n【2. VBA代码模块列表】")
        # 查找vbaProject.bin中的模块信息
        if 'xl/vbaProject.bin' in all_files:
            print("  - 包含VBA项目 (vbaProject.bin)")
            # 列出相关配置文件
            for f in all_files:
                if 'vba' in f.lower() or ' VBA' in f:
                    print(f"    • {f}")
        else:
            print("  - 未找到vbaProject.bin")

        print("\n【3. 工作表列表】")
        # 读取workbook.xml获取工作表名
        try:
            workbook_xml = zf.read('xl/workbook.xml')
            root = ET.fromstring(workbook_xml)
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheets = root.findall('.//ns:sheet', ns)
            for i, sheet in enumerate(sheets, 1):
                name = sheet.get('name', 'Unknown')
                sheet_id = sheet.get('sheetId', 'N/A')
                state = sheet.get('state', 'visible')
                print(f"  {i}. {name} (ID:{sheet_id}, 状态:{state})")
        except Exception as e:
            print(f"  读取工作表列表失败: {e}")

        print("\n【4. VBA代码提取 (如有)】")
        vba_code = extract_vba_code(zf)
        if vba_code:
            print(f"  找到 {len(vba_code)} 个VBA模块:")
            for mod_name, code in vba_code.items():
                lines = code.strip().split('\n')
                print(f"\n  --- 模块: {mod_name} ({len(lines)}行) ---")
                # 显示前50行代码
                display_code = '\n'.join(lines[:50])
                print(display_code)
                if len(lines) > 50:
                    print(f"    ... (还有 {len(lines)-50} 行)")
        else:
            print("  未检测到VBA代码")

        return vba_code

def extract_vba_code(zf):
    """从vbaProject.bin中提取VBA代码"""
    vba_modules = {}

    if 'xl/vbaProject.bin' not in zf.namelist():
        return vba_modules

    try:
        # 尝试读取相关文件
        dir_file = None
        for name in zf.namelist():
            if name.endswith('.bin') and 'dir' in name.lower():
                dir_file = name
                break

        if dir_file:
            content = zf.read(dir_file)
            # 简单解析Module信息
            try:
                text = content.decode('utf-8', errors='ignore')
                # 查找模块引用
                import re
                modules = re.findall(r'Module=(\d+)', text)
                print(f"  检测到模块数量: {len(modules)}")
            except:
                pass
    except Exception as e:
        print(f"  提取VBA代码时出错: {e}")

    return vba_modules

if __name__ == "__main__":
    import sys
    file_path = sys.argv[1] if len(sys.argv) > 1 else "收益测算表开发V4.8.5-wps版本.xlsm"
    analyze_xlsm(file_path)