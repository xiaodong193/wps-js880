#!/usr/bin/env python3
"""
同步 wps-cesuan 模块到 xlsm 文件
"""

import os
import sys
import zipfile
import shutil
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


def update_jde_data_bin(xlsm_path, module_name, new_code):
    """更新 JDEData.bin 中指定模块的代码"""
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

        match = re.search(pattern, jde_content, flags=re.DOTALL)
        if not match:
            return False, "模块未找到"

        old_code_size = len(match.group(2))

        def replacer(m):
            return m.group(1) + escaped_new_code + m.group(3)

        new_jde_content = re.sub(pattern, replacer, jde_content, flags=re.DOTALL)

        # 写回文件
        with open(jde_path, 'wb') as f:
            f.write(new_jde_content.encode('utf-8'))

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
        return True, f"原:{old_code_size} → 新:{len(escaped_new_code)} 字节"

    except Exception as e:
        return False, str(e)

    finally:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


def main():
    """主函数"""
    # 配置路径
    source_dir = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/wps-cesuan"
    xlsm_path = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/外贸租赁工作文档/外贸金租工作文档/模板/收益测算表设计/收益测算表开发V4.8.22-wps版本.xlsm"

    # 模块映射: (源文件名, 目标模块名)
    modules = [
        ("mShared_constants_v2.js", "mShared_constants"),
        ("mParameterManager_v2.js", "mParameterManager"),
        ("mInitialization_v2.js", "mInitialization"),
        ("mCashFlowGenerator_v2.js", "mCashFlowGenerator"),
        ("mRentalCalculation_v2.js", "mRentalCalculation"),
        ("mMain_v2.js", "mMain"),
        ("m货币网利率更新.js", "m货币网利率更新"),
        ("m加载项.js", "m加载项"),
        ("m银行承兑汇票模块.js", "m银行承兑汇票模块"),
        ("m测试.js", "m测试"),
    ]

    print("=" * 60)
    print("JSA 模块同步工具")
    print("=" * 60)
    print(f"源目录: {source_dir}")
    print(f"目标文件: {os.path.basename(xlsm_path)}")
    print()

    # 备份
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = f"{xlsm_path}.backup_{timestamp}"
    shutil.copy2(xlsm_path, backup_path)
    print(f"✅ 已备份: {os.path.basename(backup_path)}")
    print()

    # 同步每个模块
    success_count = 0
    for js_file, module_name in modules:
        js_path = os.path.join(source_dir, js_file)

        if not os.path.exists(js_path):
            print(f"⚠️  跳过 {module_name}: 源文件不存在 ({js_file})")
            continue

        print(f"📦 同步: {js_file} → {module_name}")

        code = read_js_file(js_path)
        success, msg = update_jde_data_bin(xlsm_path, module_name, code)

        if success:
            print(f"   ✅ 成功 ({msg})")
            success_count += 1
        else:
            print(f"   ❌ 失败: {msg}")

    print()
    print("=" * 60)
    print(f"同步完成: {success_count}/{len(modules)} 个模块成功")
    print("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(main())
