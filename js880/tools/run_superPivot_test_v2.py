#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
superPivot 测试执行脚本
在 WPS 中运行 superPivot 功能测试
"""

import os
import sys
import subprocess

# 配置路径
XLSM_FILE = "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm"

def open_wps_file():
    """打开 WPS 文件"""
    print(f"📂 打开文件: {os.path.basename(XLSM_FILE)}")
    try:
        subprocess.run(["open", XLSM_FILE], check=True)
        print("✅ 文件已打开")
        return True
    except Exception as e:
        print(f"❌ 打开失败: {e}")
        return False

def show_manual_instructions():
    """显示手动测试说明"""
    print("""
╔═══════════════════════════════════════════════════════════════╗
║            superPivot 测试执行说明                              ║
╚═══════════════════════════════════════════════════════════════╝

📋 测试已同步到 Excel 文件，包含以下模块:

   ✓ JSA880 (ID: 1) - 主框架 (12695行)
   ✓ TestHelper (ID: 2) - 测试辅助函数
   ✓ TestDataGenerator (ID: 3) - 测试数据生成
   ✓ SuperPivotWPS (ID: 4) - SuperPivot 测试套件 (480行)
   ✓ PerformanceTest (ID: 5) - 性能测试
   ✓ SuperPivotFast (ID: 6) - 极速版

📝 手动运行测试步骤:

   1. WPS 表格应该已经打开

   2. 打开 JSA 编辑器:
      - 按 Alt+F11 (Windows) 或 Option+F11 (Mac)
      - 或点击: 工具 → 开发工具 → JSA 宏编辑器

   3. 在 JSA 编辑器中找到模块 "SuperPivotWPS" (ID: 4)

   4. 在立即窗口中运行测试:

      运行单个测试:
        测试12_组织3行1列()

      运行所有测试:
        运行所有测试()

   5. 查看结果:
      - 切换到 Excel 主窗口
      - 查看 "测试输出" 工作表
      - 查看 JSA 编辑器的控制台输出

══════════════════════════════════════════════════════════════
""")

def main():
    """主函数"""
    print("""
╔═══════════════════════════════════════════════════════════════╗
║     superPivot 测试执行工具 v1.0                               ║
║     JSA880 Framework v3.8.3                                    ║
╚═══════════════════════════════════════════════════════════════╝
""")

    if not os.path.exists(XLSM_FILE):
        print(f"❌ 文件不存在: {XLSM_FILE}")
        return 1

    file_size = os.path.getsize(XLSM_FILE)
    print(f"📄 测试文件: {os.path.basename(XLSM_FILE)}")
    print(f"   文件大小: {file_size:,} 字节")

    open_wps_file()
    show_manual_instructions()

    print(f"\n✅ 测试准备完成!")
    print(f"   请按照上述说明在 WPS 中运行测试")

    return 0

if __name__ == "__main__":
    sys.exit(main())
