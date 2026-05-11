#!/usr/bin/env python3
"""
JSA880 自动同步工具 (多模块版本)
功能：自动将多个 JS 模块同步到 xlsm 文件中的 JSA 代码模块
"""

import os
import sys
import zipfile
import shutil
import xml.etree.ElementTree as ET
from xml.dom import minidom
import re
from datetime import datetime


def escape_xml(text):
    """转义 XML 特殊字符"""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&apos;')
            .replace('\n', '&#x0A;')
            .replace('\r', '&#x0D;'))


def read_js_file(js_path):
    """读取 JS 文件内容"""
    with open(js_path, 'r', encoding='utf-8') as f:
        return f.read()


def create_jde_data_bin(modules):
    """
    创建 JDEData.bin XML 内容 (支持多模块)

    Args:
        modules: 模块列表，格式: [{'name': 'JSA880', 'id': 1, 'code': '...'}, ...]
    """
    module_xmls = []
    for i, module in enumerate(modules):
        name = module['name']
        module_id = module['id']
        code = module['code']
        escaped_code = escape_xml(code)

        module_xml = f'''    <codemodule name="{name}" id="{module_id}">
        <window cursorpos="0" actived="{'true' if i == 0 else 'false'}" visible="true" />
        <codetext>{escaped_code}</codetext>
    </codemodule>'''
        module_xmls.append(module_xml)

    all_modules_xml = '\n'.join(module_xmls)

    xml_content = f'''<?xml version="1.0" encoding="UTF-8" ?>
<document version="2.0">
    <name>Project</name>
    <property desc="" lock="false" password="" />
    <activemodule>1</activemodule>
{all_modules_xml}
</document>'''
    return xml_content.encode('utf-8')


def backup_file(file_path):
    """备份文件"""
    if os.path.exists(file_path):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = f"{file_path}.backup_{timestamp}"
        shutil.copy2(file_path, backup_path)
        print(f"✅ 已备份原文件: {backup_path}")
        return backup_path
    return None


def sync_modules_to_xlsm(source_files, xlsm_path, backup=True):
    """
    将多个 JS 模块同步到 xlsm 文件

    Args:
        source_files: 源文件列表，格式: [
            {'name': 'JSA880', 'path': '/path/to/JSA880.js', 'id': 1},
            {'name': 'TestHelper', 'path': '/path/to/test_helper.js', 'id': 2}
        ]
        xlsm_path: xlsm 文件路径
        backup: 是否备份原文件
    """
    print(f"🔄 开始同步多模块...")
    print(f"   目标文件: {xlsm_path}")
    print(f"   模块数量: {len(source_files)}")

    # 检查文件是否存在
    for sf in source_files:
        if not os.path.exists(sf['path']):
            print(f"❌ 错误: 文件不存在: {sf['path']}")
            return False

    if not os.path.exists(xlsm_path):
        print(f"❌ 错误: xlsm 文件不存在: {xlsm_path}")
        return False

    # 备份
    if backup:
        backup_file(xlsm_path)

    # 读取所有模块代码
    modules = []
    total_size = 0
    for sf in source_files:
        print(f"📖 读取模块: {sf['name']}")
        code = read_js_file(sf['path'])
        modules.append({
            'name': sf['name'],
            'id': sf['id'],
            'code': code
        })
        print(f"   代码行数: {len(code.splitlines())}")
        print(f"   代码大小: {len(code)} 字节")
        total_size += len(code)

    # 创建临时目录
    temp_dir = f"{xlsm_path}.temp"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    try:
        # 解压 xlsm 文件
        print(f"📦 解压 xlsm 文件...")
        with zipfile.ZipFile(xlsm_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 创建新的 JDEData.bin
        print(f"✏️  创建 JDEData.bin (包含 {len(modules)} 个模块)...")
        jde_data = create_jde_data_bin(modules)

        # 写入 JDEData.bin
        jde_path = os.path.join(temp_dir, 'xl', 'JDEData.bin')
        with open(jde_path, 'wb') as f:
            f.write(jde_data)
        print(f"   JDEData.bin 大小: {len(jde_data)} 字节")
        print(f"   总代码大小: {total_size} 字节")

        # 重新打包为 xlsm
        print(f"📦 重新打包 xlsm 文件...")
        output_xlsm = f"{xlsm_path}.new"

        with zipfile.ZipFile(output_xlsm, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)

        # 替换原文件
        shutil.move(output_xlsm, xlsm_path)
        print(f"✅ 同步完成!")
        for m in modules:
            print(f"   ✓ {m['name']} (ID: {m['id']})")
        return True

    except Exception as e:
        print(f"❌ 同步失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # 清理临时目录
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


def main():
    """主函数"""
    # 默认路径
    work_dir = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880"
    default_xlsm = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm"

    # 定义要同步的模块
    # 可用模块配置 (仅包含实际存在的文件)
    default_modules = [
        {'name': 'JSA880', 'path': os.path.join(work_dir, 'JSA880.js'), 'id': 1},
        {'name': 'SuperPivotWPS', 'path': os.path.join(work_dir, 'src/modules/superPivot_WPS_测试.js'), 'id': 2},
        {'name': 'TestDataGenerator', 'path': os.path.join(work_dir, 'src/modules/cls生成测试数据.js'), 'id': 3},
    ]
    # 注意：仅同步 src/modules 目录下实际存在的文件

    # 解析命令行参数
    quiet = '--quiet' in sys.argv or '-q' in sys.argv
    xlsm_path = default_xlsm
    modules = default_modules

    for arg in sys.argv[1:]:
        if not arg.startswith('-'):
            # 自定义 xlsm 路径
            xlsm_path = arg

    # 执行同步
    success = sync_modules_to_xlsm(modules, xlsm_path)

    if quiet:
        # 安静模式，只输出关键信息
        if success:
            print("✅ 同步完成")
        else:
            print("❌ 同步失败")

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
