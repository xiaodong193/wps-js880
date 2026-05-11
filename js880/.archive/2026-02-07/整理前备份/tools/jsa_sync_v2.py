#!/usr/bin/env python3
"""
JSA880 同步工具 V2
直接修改xlsm中的JDEData.bin内容，保持原始格式
"""

import os
import sys
import zipfile
import shutil
import re
from datetime import datetime


def escape_xml(text):
    """转义 XML 特殊字符 - 完全匹配原始格式"""
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


def update_jde_data_bin(xlsm_path, module_name, new_code):
    """
    直接更新 JDEData.bin 中指定模块的代码

    Args:
        xlsm_path: xlsm 文件路径
        module_name: 要更新的模块名称
        new_code: 新的代码内容
    """
    temp_dir = f"{xlsm_path}.temp"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    try:
        # 解压 xlsm 文件
        with zipfile.ZipFile(xlsm_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 读取现有的 JDEData.bin
        jde_path = os.path.join(temp_dir, 'xl', 'JDEData.bin')
        with open(jde_path, 'rb') as f:
            jde_content = f.read().decode('utf-8')

        # 替换指定模块的代码
        pattern = rf'(<codemodule name="{module_name}" id="\d+">.*?<codetext>)(.*?)(</codetext>)'
        escaped_new_code = escape_xml(new_code)

        def replacer(match):
            return match.group(1) + escaped_new_code + match.group(3)

        new_jde_content = re.sub(pattern, replacer, jde_content, flags=re.DOTALL)

        # 检查是否成功替换
        if new_jde_content == jde_content:
            print(f"❌ 未找到模块: {module_name}")
            return False

        # 写回文件
        with open(jde_path, 'wb') as f:
            f.write(new_jde_content.encode('utf-8'))

        print(f"✅ 已更新模块: {module_name}")
        print(f"   原代码大小: {len(match.group(2)) if 'match' in locals() else 'N/A'} 字节")
        print(f"   新代码大小: {len(escaped_new_code)} 字节")

        # 重新打包
        output_xlsm = f"{xlsm_path}.new"
        with zipfile.ZipFile(output_xlsm, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)

        # 替换原文件
        shutil.move(output_xlsm, xlsm_path)
        return True

    except Exception as e:
        print(f"❌ 更新失败: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


def main():
    """主函数"""
    work_dir = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880"
    xlsm_path = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm"

    # 备份
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = f"{xlsm_path}.backup_{timestamp}"
    shutil.copy2(xlsm_path, backup_path)
    print(f"✅ 已备份: {backup_path}")

    # 更新 JSA880 模块
    jsa880_path = os.path.join(work_dir, 'JSA880.js')
    print(f"📖 读取: {jsa880_path}")
    new_code = read_js_file(jsa880_path)
    print(f"   代码行数: {len(new_code.splitlines())}")

    success = update_jde_data_bin(xlsm_path, 'JSA880', new_code)

    if success:
        print("\n✅ 同步完成!")
        print(f"版本: v3.8.4")
    else:
        print("\n❌ 同步失败")

    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
