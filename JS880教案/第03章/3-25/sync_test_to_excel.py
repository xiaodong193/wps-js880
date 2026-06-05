#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
SuperPivot 测试代码同步到 Excel
将测试代码同步到 xlsm 文件的 VBA 模块中
"""

import os
import sys
from datetime import datetime

# 文件路径配置
EXCEL_FILE = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm"
TEST_FILE = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/test_to_paste.js"
JSA880_FILE = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/src/JSA880.js"

# 输出文件（准备好的粘贴文件）
OUTPUT_DIR = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25"
TEST_OUTPUT = os.path.join(OUTPUT_DIR, "test_to_paste.js")
JSA880_OUTPUT = os.path.join(OUTPUT_DIR, "JSA880_to_paste.js")

def print_header(title):
    """打印标题"""
    print("\n" + "=" * 70)
    print(f"  {title}")
    print("=" * 70 + "\n")

def print_step(step_num, description):
    """打印步骤"""
    print(f"\n{'='*70}")
    print(f"  步骤 {step_num}: {description}")
    print(f"{'='*70}\n")

def check_file_exists(filepath, description):
    """检查文件是否存在"""
    if os.path.exists(filepath):
        size = os.path.getsize(filepath)
        print(f"✅ {description} 存在")
        print(f"   文件: {os.path.basename(filepath)}")
        print(f"   大小: {size:,} bytes ({size/1024:.2f} KB)")
        return True
    else:
        print(f"❌ {description} 不存在！")
        print(f"   路径: {filepath}")
        return False

def copy_file(source, dest, description):
    """复制文件"""
    try:
        import shutil
        shutil.copy2(source, dest)
        size = os.path.getsize(dest)
        print(f"✅ {description} 已准备好")
        print(f"   源文件: {os.path.basename(source)}")
        print(f"   目标文件: {os.path.basename(dest)}")
        print(f"   文件大小: {size:,} bytes ({size/1024:.2f} KB)")
        return True
    except Exception as e:
        print(f"❌ 复制失败: {str(e)}")
        return False

def main():
    """主函数"""
    print_header("SuperPivot 测试代码同步工具")
    
    print(f"开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    all_ok = True
    
    # 步骤 1: 检查源文件
    print_step(1, "检查源文件")
    
    if not check_file_exists(TEST_FILE, "测试文件"):
        all_ok = False
    else:
        # 读取测试文件并显示部分内容
        with open(TEST_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
            print(f"\n   测试文件包含 {len(content.splitlines())} 行代码")
            print(f"   包含 {content.count('test')} 个测试函数")
    
    print()
    
    if not check_file_exists(JSA880_FILE, "JSA880.js 文件"):
        print("\n⚠️  警告: JSA880.js 不存在，但测试仍然可以运行（如果 Excel 中已有）")
    else:
        # 读取 JSA880.js 并显示信息
        with open(JSA880_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
            print(f"\n   JSA880.js 包含 {len(content.splitlines())} 行代码")
            # 检查版本
            if 'version:' in content or 'VERSION' in content:
                print(f"   包含版本信息")
    
    # 步骤 2: 准备输出文件
    print_step(2, "准备输出文件")
    
    # 复制测试文件
    if not copy_file(TEST_FILE, TEST_OUTPUT, "测试代码"):
        all_ok = False
    
    print()
    
    # 复制 JSA880 文件（如果存在）
    if os.path.exists(JSA880_FILE):
        copy_file(JSA880_FILE, JSA880_OUTPUT, "JSA880 代码")
    
    # 步骤 3: 检查 Excel 文件
    print_step(3, "检查 Excel 文件")
    
    if check_file_exists(EXCEL_FILE, "Excel 测试文件"):
        print(f"\n   Excel 文件已准备就绪！")
    
    # 步骤 4: 生成操作指南
    print_step(4, "生成操作指南")
    
    print("📋 下一步操作指南:\n")
    
    print("1️⃣  打开 Excel 文件:")
    print(f"   文件名: {os.path.basename(EXCEL_FILE)}")
    print(f"   位置: {EXCEL_FILE}")
    
    print("\n2️⃣  打开 VBA 编辑器:")
    print("   - Windows: 按 Alt + F11")
    print("   - Mac: 按 Fn + Alt + F11")
    
    print("\n3️⃣  加载 JSA880.js（如果还没有）:")
    print("   a) 在 VBA 编辑器中，右键点击 → 插入 → 模块")
    print("   b) 打开文件: JSA880_to_paste.js")
    print("   c) 全选复制（Ctrl+A），复制（Ctrl+C）")
    print("   d) 粘贴到模块中（Ctrl+V）")
    print("   e) 保存（Ctrl+S）")
    
    print("\n4️⃣  加载测试代码:")
    print("   a) 在 VBA 编辑器中，右键点击 → 插入 → 模块")
    print("   b) 打开文件: test_to_paste.js")
    print("   c) 全选复制（Ctrl+A），复制（Ctrl+C）")
    print("   d) 粘贴到模块中（Ctrl+V）")
    print("   e) 保存（Ctrl+S）")
    
    print("\n5️⃣  运行测试:")
    print("   a) 打开立即窗口: 按 Ctrl + G")
    print("   b) 输入: runQuickTests")
    print("   c) 按回车运行")
    print("   d) 查看结果")
    
    print("\n6️⃣  运行完整测试（可选）:")
    print("   a) 在立即窗口中输入: runAllTests")
    print("   b) 按回车运行")
    print("   c) 查看完整结果")
    
    # 总结
    print_step("总结", "同步完成")
    
    print(f"✅ 文件已准备好！")
    print(f"\n📁 准备好的文件:")
    print(f"   1. {TEST_OUTPUT}")
    print(f"   2. {JSA880_OUTPUT}")
    print(f"\n📖 详细文档:")
    print(f"   - README_测试套件完整指南.md")
    print(f"   - 快速参考卡.md")
    print(f"   - 测试调试分步指南.md")
    
    print(f"\n⏰ 完成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    if all_ok:
        print("\n✅ 所有检查通过！现在可以开始测试了！")
        print("\n🚀 快速开始命令:")
        print("   runQuickTests")
        return 0
    else:
        print("\n⚠️  部分检查失败，请检查文件路径")
        return 1

if __name__ == '__main__':
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 错误: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
