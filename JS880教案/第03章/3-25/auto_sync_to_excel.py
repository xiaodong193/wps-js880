#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动将测试代码同步到 Excel xlsm 文件
使用 xlwings 库直接操作 Excel
"""

import os
import sys
import subprocess
from datetime import datetime

# 配置
EXCEL_FILE = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm"
TEST_CODE_FILE = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/test_to_paste.js"
JSA880_CODE_FILE = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/JSA880_to_paste.js"

def print_header(title):
    print("\n" + "=" * 80)
    print(f"  {title}")
    print("=" * 80 + "\n")

def check_xlwings():
    """检查 xlwings 是否安装"""
    try:
        import xlwings
        print(f"✅ xlwings 已安装 (版本: {xlwings.__version__})")
        return True
    except ImportError:
        print("❌ xlwings 未安装")
        print("\n📦 安装 xlwings:")
        print("   pip3 install xlwings")
        print("\n   或使用:")
        print("   python3 -m pip install xlwings")
        return False

def check_excel_app():
    """检查 Excel 应用是否可用"""
    try:
        import xlwings as xw
        # 尝试连接到 Excel
        app = xw.App(visible=False)
        app.quit()
        print("✅ Excel 应用可用")
        return True
    except Exception as e:
        print(f"❌ Excel 应用不可用: {str(e)}")
        print("\n💡 提示:")
        print("   - 确保 Excel/WPS 已安装")
        print("   - 确保 Excel/WPS 可以正常打开")
        return False

def sync_to_excel_auto():
    """自动同步到 Excel"""
    print_header("自动同步测试代码到 Excel")
    
    # 检查依赖
    if not check_xlwings():
        return False
    
    if not check_excel_app():
        return False
    
    print(f"开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    try:
        import xlwings as xw
        
        # 检查文件存在
        if not os.path.exists(EXCEL_FILE):
            print(f"❌ Excel 文件不存在: {EXCEL_FILE}")
            return False
        
        if not os.path.exists(TEST_CODE_FILE):
            print(f"❌ 测试代码文件不存在: {TEST_CODE_FILE}")
            return False
        
        print("✅ 所有文件检查通过\n")
        
        # 读取代码
        print("📖 读取测试代码...")
        with open(TEST_CODE_FILE, 'r', encoding='utf-8') as f:
            test_code = f.read()
        print(f"   测试代码: {len(test_code.splitlines())} 行\n")
        
        print("📖 读取 JSA880 代码...")
        with open(JSA880_CODE_FILE, 'r', encoding='utf-8') as f:
            jsa880_code = f.read()
        print(f"   JSA880 代码: {len(jsa880_code.splitlines())} 行\n")
        
        # 打开 Excel
        print("📂 打开 Excel 文件...")
        app = xw.App(visible=True)
        wb = app.books.open(EXCEL_FILE)
        print(f"   已打开: {os.path.basename(EXCEL_FILE)}\n")
        
        # 检查是否已有模块
        print("🔍 检查现有模块...")
        existing_modules = []
        try:
            # 尝试访问 VBA 项目
            vba = wb.api.VBProject
            for component in vba.VBComponents:
                if component.Type == 1:  # vbext_ct_StdModule
                    existing_modules.append(component.Name)
                    print(f"   现有模块: {component.Name}")
        except Exception as e:
            print(f"   ⚠️  无法访问 VBA 项目: {str(e)}")
            print("   💡 提示: 确保 Excel 宏设置允许访问 VBA 项目")
            print("   Excel → 选项 → 信任中心 → 信任对 VBA 工程对象模型的访问")
        
        print()
        
        # 添加或更新测试模块
        print("➕ 添加/更新测试模块...")
        if 'SuperPivotTests' in existing_modules:
            print("   删除旧模块: SuperPivotTests")
            try:
                wb.api.VBProject.VBComponents.Remove(wb.api.VBProject.VBComponents("SuperPivotTests"))
            except:
                pass
        
        # 添加新模块
        try:
            module = wb.api.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
            module.Name = "SuperPivotTests"
            module.CodeModule.AddFromString(test_code)
            print(f"   ✅ 测试模块已添加/更新: SuperPivotTests")
            print(f"   代码行数: {len(test_code.splitlines())}")
        except Exception as e:
            print(f"   ❌ 添加测试模块失败: {str(e)}")
            print("\n💡 可能的解决方案:")
            print("   1. 手动复制粘贴（推荐）")
            print("   2. 检查 Excel 安全设置")
            print("   3. 以管理员身份运行")
            wb.close()
            app.quit()
            return False
        
        print()
        
        # 添加或更新 JSA880 模块（如果需要）
        if 'JSA880' not in existing_modules:
            print("➕ 添加 JSA880 模块...")
            try:
                module = wb.api.VBProject.VBComponents.Add(1)
                module.Name = "JSA880"
                module.CodeModule.AddFromString(jsa880_code)
                print(f"   ✅ JSA880 模块已添加")
                print(f"   代码行数: {len(jsa880_code.splitlines())}")
            except Exception as e:
                print(f"   ⚠️  添加 JSA880 模块失败: {str(e)}")
                print("   (如果已有 JSA880，可以忽略)")
        else:
            print("ℹ️  JSA880 模块已存在，跳过")
        
        print()
        
        # 保存
        print("💾 保存文件...")
        wb.save()
        print("   ✅ 文件已保存\n")
        
        # 显示下一步
        print_header("同步完成！")
        
        print("✅ 测试代码已成功添加到 Excel！\n")
        
        print("📋 下一步操作:\n")
        print("1️⃣  Excel 文件现在已打开")
        print("2️⃣  按 Alt + F11 打开 VBA 编辑器")
        print("3️⃣  你会看到两个新模块:")
        print("     - SuperPivotTests (测试代码)")
        print("     - JSA880 (框架代码)")
        print("4️⃣  按 Ctrl + G 打开立即窗口")
        print("5️⃣  运行测试:")
        print("     runQuickTests")
        print()
        
        # 不关闭 Excel，让用户可以立即使用
        print("💡 提示:")
        print("   Excel 文件保持打开状态")
        print("   你现在可以立即运行测试！")
        print("   测试完成后记得保存文件 (Ctrl + S)")
        print()
        
        return True
        
    except Exception as e:
        print(f"\n❌ 错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主函数"""
    try:
        success = sync_to_excel_auto()
        
        if success:
            print(f"\n⏰ 完成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print("\n🎉 同步成功！现在可以在 Excel 中运行测试了！\n")
            print("🚀 快速开始:")
            print("   1. 在 Excel 中按 Alt + F11 打开 VBA 编辑器")
            print("   2. 按 Ctrl + G 打开立即窗口")
            print("   3. 输入: runQuickTests")
            print("   4. 按回车运行\n")
            return 0
        else:
            print("\n⚠️  自动同步失败")
            print("\n📖 请使用手动方法:")
            print("   1. 打开 Excel 文件")
            print("   2. 打开 VBA 编辑器 (Alt + F11)")
            print("   3. 创建新模块")
            print("   4. 复制粘贴 test_to_paste.js 的内容")
            print("   5. 运行测试: runQuickTests\n")
            return 1
            
    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断")
        return 1
    except Exception as e:
        print(f"\n❌ 错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == '__main__':
    sys.exit(main())
