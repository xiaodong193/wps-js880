#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JSA880 统一工具脚本
整合代码同步、测试执行、版本管理等功能
"""

import os
import sys
import zipfile
import shutil
import subprocess
import argparse
import json
import platform
import fcntl
from datetime import datetime
from pathlib import Path

# ==================== 配置 ====================

class Config:
    """配置管理"""
    # 工具脚本位于 tools/ 目录，项目根目录是其父目录
    BASE_DIR = Path(__file__).parent.parent
    XLSM_FILE = Path("/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm")
    BACKUP_DIR = BASE_DIR / ".backups"
    TEMP_DIR = BASE_DIR / ".temp"

    # 模块配置 (仅包含实际存在的文件)
    # 注意: xlsm 文件中原始模块名是 mJSA880，需要保持一致
    MODULES = [
        {'name': 'mJSA880', 'file': 'JSA880.js', 'id': 1},
        {'name': 'TestDataGenerator', 'file': 'src/modules/cls生成测试数据.js', 'id': 3},
        {'name': 'SuperPivotV390', 'file': 'src/modules/superPivot_v390.js', 'id': 4},
        {'name': 'SuperPivotV391', 'file': 'src/modules/superPivot_v391.js', 'id': 5},
        {'name': 'ErrorHandler', 'file': 'src/modules/clsErrorHandler.js', 'id': 15},
        {'name': 'Logger', 'file': 'src/modules/clsLogger.js', 'id': 16},
        {'name': 'ParameterValidator', 'file': 'src/modules/clsParameterValidator.js', 'id': 17},
    ]

    # 版本信息
    VERSION_FILE = BASE_DIR / ".version.json"
    CURRENT_VERSION = "3.8.3"

    @classmethod
    def ensure_dirs(cls):
        """确保必要的目录存在"""
        cls.BACKUP_DIR.mkdir(exist_ok=True)
        cls.TEMP_DIR.mkdir(exist_ok=True)

# ==================== 文件锁检测 ====================

class FileLockDetector:
    """文件占用检测器"""

    @staticmethod
    def is_file_locked(filepath):
        """检测文件是否被占用"""
        try:
            # 尝试以写入模式打开文件（但不修改）
            with open(filepath, 'r+b') as f:
                if platform.system() == 'Linux' or platform.system() == 'Darwin':
                    # Unix: 尝试获取排他锁
                    try:
                        fcntl.lockf(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                        fcntl.lockf(f.fileno(), fcntl.LOCK_UN)  # 释放锁
                        return False
                    except IOError:
                        return True
                else:
                    # Windows: 文件打开成功则未被占用
                    return False
        except (IOError, OSError, PermissionError):
            return True

    @staticmethod
    def is_wps_running():
        """检测 WPS 是否正在运行"""
        system = platform.system()
        try:
            if system == 'Darwin':  # macOS
                # 检查多种可能的 WPS 进程名
                wps_names = ['wpsoffice', 'wps', 'wpsoffice']
                for name in wps_names:
                    result = subprocess.run(
                        ['pgrep', '-x', name],
                        capture_output=True,
                        text=True
                    )
                    if result.returncode == 0:
                        return True
                return False
            elif system == 'Windows':
                result = subprocess.run(
                    ['tasklist', '/FI', 'IMAGENAME eq wps.exe'],
                    capture_output=True,
                    text=True
                )
                return 'wps.exe' in result.stdout.lower()
            else:  # Linux
                wps_names = ['wpsoffice', 'wps', 'wpsoffice']
                for name in wps_names:
                    result = subprocess.run(
                        ['pgrep', '-x', name],
                        capture_output=True,
                        text=True
                    )
                    if result.returncode == 0:
                        return True
                return False
        except Exception:
            return False

    @staticmethod
    def close_wps():
        """关闭 WPS（仅 macOS）"""
        system = platform.system()
        if system == 'Darwin':
            try:
                # 使用 AppleScript 关闭 WPS
                script = '''
                tell application "WPS Office"
                    quit
                end tell
                '''
                subprocess.run(['osascript', '-e', script], check=True, capture_output=True)
                return True
            except Exception as e:
                print(f"  ⚠️  无法自动关闭 WPS: {e}")
                return False
        elif system == 'Windows':
            try:
                subprocess.run(['taskkill', '/IM', 'wps.exe', '/F'], check=True, capture_output=True)
                return True
            except Exception as e:
                print(f"  ⚠️  无法自动关闭 WPS: {e}")
                return False
        else:
            print(f"  ⚠️  自动关闭 WPS 功能仅支持 macOS 和 Windows")
            return False

# ==================== WPS 日志读取 ====================

class WPSLogReader:
    """WPS 立即窗口日志读取器"""

    def __init__(self):
        self.config = Config
        self.log_file = self.config.TEMP_DIR / "wps_immediate_log.txt"

    def read_immediate_window(self, auto_mode=True):
        """读取 WPS 立即窗口内容

        Args:
            auto_mode: 是否尝试自动读取（True），或仅返回提示（False）
        """
        system = platform.system()

        if system == 'Darwin':
            return self._read_macos_auto() if auto_mode else self._read_macos_manual()
        elif system == 'Windows':
            return self._read_windows()
        else:
            return "❌ 不支持的平台: 需要手动复制立即窗口内容"

    def _read_macos_auto(self):
        """macOS: 尝试自动读取"""
        try:
            # 尝试调用自动读取脚本
            auto_reader_path = self.config.BASE_DIR / 'tools' / 'wps_auto_log_reader.py'
            if auto_reader_path.exists():
                result = subprocess.run(
                    ['python3', str(auto_reader_path), '--method', 'hybrid'],
                    capture_output=True,
                    text=True,
                    timeout=30,
                    cwd=str(self.config.BASE_DIR)
                )
                if result.returncode == 0 and "✅" in result.stdout:
                    # 从输出中提取成功信息
                    for line in result.stdout.split('\n'):
                        if "✅" in line and "行日志" in line:
                            return line
                    return "✅ 自动读取成功"
                else:
                    # 自动读取失败，返回详细错误
                    return f"⚠️  自动读取失败，请使用手动模式\n{result.stdout}"

            # 如果脚本不存在，使用内置方法
            return self._read_macos_builtin()

        except subprocess.TimeoutExpired:
            return "⚠️  自动读取超时，请使用手动模式: ./jsa log --paste"
        except Exception as e:
            return f"⚠️  自动读取出错: {e}\n💡 请使用手动模式: ./jsa log --paste"

    def _read_macos_builtin(self):
        """macOS: 内置的自动读取方法"""
        # 步骤 1: 检查 WPS 是否运行
        check_script = '''
tell application "System Events"
    set isRunning to (name of processes) contains "wpsoffice"
end tell
if isRunning then
    return "RUNNING"
else
    return "NOT_RUNNING"
end if
'''

        check_result = subprocess.run(['osascript', '-e', check_script],
                                      capture_output=True, text=True, timeout=5)

        wps_running = "RUNNING" in check_result.stdout

        if not wps_running:
            return "❌ WPS 未运行，请先打开 WPS"

        # 步骤 2: 尝试自动复制立即窗口内容
        auto_copy_script = '''
tell application "System Events"
    tell process "wpsoffice"
        -- 打开宏编辑器 (Option+F11)
        keystroke "f11" using {option down}
        delay 2

        -- 尝试聚焦立即窗口
        keystroke "g" using {command down}
        delay 0.5

        -- 全选并复制
        keystroke "a" using {command down}
        delay 0.3
        keystroke "c" using {command down}
        delay 0.5
    end tell
end tell
return "COPIED"
'''

        # 执行自动复制
        subprocess.run(['osascript', '-e', auto_copy_script],
                      capture_output=True, text=True, timeout=10)

        # 步骤 3: 从剪贴板读取内容
        get_clipboard_script = '''
set theContent to the clipboard
return theContent
'''

        result = subprocess.run(['osascript', '-e', get_clipboard_script],
                              capture_output=True, text=True, timeout=5)

        if result.returncode == 0:
            content = result.stdout

            if content and content.strip():
                # 保存到文件
                self.config.ensure_dirs()
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                with open(self.log_file, 'w', encoding='utf-8') as f:
                    f.write(f"{'='*80}\n")
                    f.write(f"📅 自动读取的日志 - {timestamp}\n")
                    f.write(f"{'='*80}\n")
                    f.write(content)
                    if not content.endswith('\n'):
                        f.write("\n")

                lines = content.count('\n') + 1
                return f"✅ 成功读取 {lines} 行日志\n💾 已保存到: {self.log_file.name}"
            else:
                return "⚠️  剪贴板为空，可能复制失败\n💡 建议: ./jsa log --paste"
        else:
            return f"⚠️  读取剪贴板失败\n💡 建议: ./jsa log --paste"

    def _read_macos_manual(self):
        """macOS: 返回手动操作提示"""
        return """⚠️  自动读取已禁用

💡 推荐使用粘贴模式:
   1. 在 WPS 宏编辑器立即窗口中按 Cmd+A (全选)
   2. 按 Cmd+C (复制)
   3. 运行: ./jsa log --paste
   4. 粘贴内容后按 Ctrl+D 结束

📖 或者使用:
   ./jsa log --help    # 查看更多选项"""

    def _read_windows(self):
        """Windows: 使用快捷键复制"""
        try:
            # 使用 PowerShell 获取剪贴板
            ps_script = '''
Add-Type -AssemblyName System.Windows.Forms

# 激活 WPS
$wps = Get-Process wpsoffice -ErrorAction SilentlyContinue
if ($wps) {
    $wps.MainWindowHandle | Out-Null
    # 发送快捷键到 WPS 立即窗口
    [System.Windows.Forms.SendKeys]::WaitSend("^a", 100)  # Ctrl+A
    [System.Windows.Forms.SendKeys]::WaitSend("^c", 100)  # Ctrl+C
    Start-Sleep -Milliseconds 300

    # 获取剪贴板内容
    $content = [System.Windows.Forms.Clipboard]::GetText()
    Write-Output $content
} else {
    Write-Output "ERROR: WPS 未运行"
}
'''

            result = subprocess.run(
                ['powershell', '-Command', ps_script],
                capture_output=True,
                text=True,
                timeout=10
            )

            content = result.stdout.strip()

            if content.startswith("ERROR:"):
                return f"❌ {content[6:]}"

            if content:
                self.config.ensure_dirs()
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(f"\n{'='*80}\n")
                    f.write(f"📅 WPS 立即窗口日志 - {timestamp}\n")
                    f.write(f"{'='*80}\n")
                    f.write(content)
                    f.write("\n")

                lines = content.count('\n') + 1
                return f"✅ 成功读取 {lines} 行内容\n💾 已保存到: {self.log_file.name}"
            return "ℹ️  立即窗口为空"

        except subprocess.TimeoutExpired:
            return "❌ 操作超时"
        except Exception as e:
            return f"❌ 读取失败: {e}"

    def get_saved_log(self, lines=None):
        """获取已保存的日志"""
        if not self.log_file.exists():
            return "ℹ️  暂无已保存的日志\n   提示: 先运行 log 命令读取 WPS 立即窗口内容"

        try:
            with open(self.log_file, 'r', encoding='utf-8') as f:
                content = f.read()

            if lines:
                # 获取最后 N 行
                all_lines = content.split('\n')
                last_lines = all_lines[-lines:]
                return '\n'.join(last_lines)

            return content
        except Exception as e:
            return f"❌ 读取日志文件失败: {e}"

    def clear_log(self):
        """清空日志文件"""
        if self.log_file.exists():
            self.log_file.unlink()
            return "✅ 日志文件已清空"
        return "ℹ️  没有需要清空的日志"

# ==================== WPS 自动化 ====================

class WPSAutomation:
    """WPS 自动化操作 - 打开文件、运行测试、收集日志"""

    def __init__(self, xlsm_path=None):
        self.config = Config
        self.xlsm_path = Path(xlsm_path) if xlsm_path else self.config.XLSM_FILE
        self.system = platform.system()
        self.log_reader = WPSLogReader()

    def open_wps(self):
        """打开 WPS 文件"""
        if not self.xlsm_path.exists():
            return f"❌ 文件不存在: {self.xlsm_path}"

        try:
            if self.system == 'Darwin':  # macOS
                subprocess.run(['open', str(self.xlsm_path)], check=True)
            elif self.system == 'Windows':
                os.startfile(str(self.xlsm_path))
            else:
                subprocess.run(['xdg-open', str(self.xlsm_path)], check=True)
            return "✅ 已打开 WPS 文件"
        except Exception as e:
            return f"❌ 打开文件失败: {e}"

    def run_test_by_applescript(self, test_function="运行所有测试"):
        """通过 AppleScript 运行测试 (仅 macOS)"""
        if self.system != 'Darwin':
            return "⚠️  自动运行测试仅支持 macOS"

        script = f'''
tell application "System Events"
    set isRunning to (name of processes) contains "wpsoffice"
end tell

if isRunning then
    tell application "WPS Office"
        activate
    end tell

    delay 1

    tell application "System Events"
        tell process "wpsoffice"
            -- 打开宏编辑器 (Alt+F11)
            keystroke "f11" using {{option down}}
            delay 1

            -- 聚焦到立即窗口 (Cmd+4 或 Ctrl+G)
            keystroke "4" using {{command down}}
            delay 0.5

            -- 在立即窗口输入测试命令
            keystroke "{test_function}()"
            keystroke return

            -- 等待测试执行
            delay 2
        end tell
    end tell

    return "SUCCESS: 测试已触发"
else
    return "ERROR: WPS 未运行"
end if
'''

        try:
            result = subprocess.run(
                ['osascript', '-e', script],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode == 0:
                content = result.stdout.strip()
                if content.startswith("SUCCESS"):
                    return "✅ 测试已触发，正在执行..."
                elif content.startswith("ERROR"):
                    return f"❌ {content[6:]}"
                else:
                    return f"ℹ️  {content}"
            else:
                return f"⚠️  AppleScript 执行问题: {result.stderr}"
        except subprocess.TimeoutExpired:
            return "⚠️  操作超时"
        except Exception as e:
            return f"⚠️  自动触发失败: {e}"

    def collect_logs_from_clipboard(self):
        """从剪贴板收集日志 (macOS/Linux)"""
        try:
            if self.system == 'Darwin':  # macOS
                result = subprocess.run(
                    ['pbpaste'],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                content = result.stdout
            elif self.system == 'Windows':
                result = subprocess.run(
                    ['powershell', '-Command', 'Get-Clipboard'],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                content = result.stdout
            else:
                # Linux (xclip)
                result = subprocess.run(
                    ['xclip', '-selection', 'clipboard', '-o'],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                content = result.stdout

            if content and content.strip():
                # 保存日志
                self.config.ensure_dirs()
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                with open(self.log_reader.log_file, 'w', encoding='utf-8') as f:
                    f.write(f"{'='*80}\n")
                    f.write(f"📅 自动收集的日志 - {timestamp}\n")
                    f.write(f"{'='*80}\n")
                    f.write(content)
                    f.write("\n" + "="*80 + "\n")

                lines = content.count('\n') + 1
                return f"✅ 已收集 {lines} 行日志"
            return "ℹ️  剪贴板为空或无日志内容"
        except FileNotFoundError:
            return "⚠️  需要剪贴板工具 (macOS: 内置, Windows: PowerShell, Linux: xclip)"
        except Exception as e:
            return f"⚠️  收集失败: {e}"

    def wait_for_test(self, seconds=10):
        """等待测试完成"""
        import time
        print(f"  ⏱️  等待测试执行 ({seconds} 秒)...")
        for i in range(seconds):
            remaining = seconds - i
            print(f"  ⏳  {remaining} 秒...", end='\r')
            time.sleep(1)
        print("  ✅ 等待完成")

# ==================== 工作流管理 ====================

class WorkflowManager:
    """测试自动化工作流管理器"""

    def __init__(self, xlsm_path=None, selected_modules=None):
        self.config = Config
        self.xlsm_path = Path(xlsm_path) if xlsm_path else self.config.XLSM_FILE
        self.selected_modules = selected_modules
        self.automation = WPSAutomation(self.xlsm_path)
        self.log_reader = WPSLogReader()

    def run_workflow(self, mode='auto', test_function="运行所有测试", wait_time=10):
        """执行完整工作流

        Args:
            mode: 工作流模式
                - 'auto': 自动模式（同步 → 打开 → 运行测试 → 收集日志）
                - 'sync': 仅同步
                - 'test': 仅运行测试
                - 'collect': 仅收集日志
                - 'manual': 手动模式（同步 → 打开 → 等待手动操作）
            test_function: 要运行的测试函数名
            wait_time: 等待测试完成的秒数
        """
        print("🚀 JSA880 测试自动化工作流")
        print("="*60)

        # 步骤 1: 同步代码
        if mode in ['auto', 'sync', 'manual']:
            print("\n📦 步骤 1/4: 同步代码")
            print("-"*60)

            # 检查 WPS 是否运行
            if FileLockDetector.is_wps_running():
                print("  ⚠️  WPS 正在运行")
                print("  💡 提示: 使用 --auto-close 自动关闭 WPS 后同步")
                return False

            sync_manager = SyncManager(
                self.xlsm_path,
                self.selected_modules,
                force=False,
                auto_close=False
            )
            success = sync_manager.sync(backup=True)
            if not success:
                print("  ❌ 同步失败")
                return False
            print("  ✅ 同步完成")

        # 步骤 2: 打开 WPS
        if mode in ['auto', 'test', 'manual']:
            print("\n📂 步骤 2/4: 打开 WPS 文件")
            print("-"*60)
            result = self.automation.open_wps()
            print(f"  {result}")

            if "❌" in result:
                return False

            # 等待 WPS 完全加载
            import time
            time.sleep(2)

        # 步骤 3: 运行测试
        if mode in ['auto', 'test']:
            print("\n🧪 步骤 3/4: 运行测试")
            print("-"*60)

            result = self.automation.run_test_by_applescript(test_function)
            print(f"  {result}")

            if "✅" in result or "⏱️" not in result:
                # 等待测试完成
                self.automation.wait_for_test(wait_time)

        elif mode == 'manual':
            print("\n🧪 步骤 3/4: 手动运行测试")
            print("-"*60)
            print("  请在 WPS 中手动运行测试:")
            print(f"    1. 按 Alt+F11 打开宏编辑器")
            print(f"    2. 在立即窗口输入: {test_function}()")
            print(f"    3. 按 Enter 运行")
            print()
            input("  完成后按 Enter 继续...")

        # 步骤 4: 收集日志
        if mode in ['auto', 'test', 'collect']:
            print("\n📋 步骤 4/4: 收集日志")
            print("-"*60)

            # 尝试从剪贴板收集
            result = self.automation.collect_logs_from_clipboard()
            print(f"  {result}")

            if "剪贴板为空" in result or "收集失败" in result:
                print("\n  💡 请手动复制日志:")
                print("     1. 在 WPS 宏编辑器立即窗口按 Cmd+A (全选)")
                print("     2. 按 Cmd+C (复制)")
                print("     3. 运行: ./jsa log --paste")

        # 显示日志分析
        if mode in ['auto', 'test', 'collect']:
            print("\n📊 日志分析")
            print("-"*60)
            content = self.log_reader.get_saved_log()
            if "暂无已保存的日志" not in content:
                _analyze_log(content)
            else:
                print("  ℹ️  暂无日志可供分析")

        print("\n" + "="*60)
        print("✅ 工作流完成")
        return True

# ==================== 版本管理 ====================

class VersionManager:
    """版本管理器"""

    def __init__(self):
        self.config = Config
        self.version_file = self.config.VERSION_FILE

    def get_version(self):
        """获取当前版本"""
        if self.version_file.exists():
            with open(self.version_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('version', 'unknown')
        return self.config.CURRENT_VERSION

    def set_version(self, version):
        """设置版本"""
        with open(self.version_file, 'w', encoding='utf-8') as f:
            json.dump({
                'version': version,
                'updated_at': datetime.now().isoformat(),
                'updated_by': 'jsa_tools.py'
            }, f, indent=2, ensure_ascii=False)

    def get_changelog(self):
        """获取更新日志"""
        # 从 JSA880.js 文件头部提取
        jsa_file = self.config.BASE_DIR / 'JSA880.js'
        if jsa_file.exists():
            with open(jsa_file, 'r', encoding='utf-8') as f:
                lines = []
                for line in f:
                    if '更新日志' in line or '=====' in line or line.strip().startswith('v3.'):
                        lines.append(line)
                    if len(lines) > 50:  # 限制行数
                        break
                return ''.join(lines)
        return "无法读取更新日志"

# ==================== 错误处理 ====================

class ErrorHandler:
    """错误处理和回滚管理"""

    def __init__(self, xlsm_path=None):
        self.config = Config
        self.xlsm_path = Path(xlsm_path) if xlsm_path else self.config.XLSM_FILE
        self.backup_path = None
        self.original_data = None

    def backup_xlsm(self):
        """备份 xlsm 文件"""
        if not self.xlsm_path.exists():
            return None

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        self.backup_path = self.config.BACKUP_DIR / f"{self.xlsm_path.stem}_backup_{timestamp}{self.xlsm_path.suffix}"

        shutil.copy2(self.xlsm_path, self.backup_path)
        print(f"  💾 已备份到: {self.backup_path.name}")
        return self.backup_path

    def rollback(self):
        """回滚到备份版本"""
        if self.backup_path and self.backup_path.exists():
            shutil.copy2(self.backup_path, self.xlsm_path)
            print(f"  ↩️  已回滚到备份版本")
            return True
        return False

# ==================== 代码同步 ====================

class SyncManager:
    """代码同步管理器"""

    def __init__(self, xlsm_path=None, selected_modules=None, force=False, auto_close=False):
        self.config = Config
        self.xlsm_path = Path(xlsm_path) if xlsm_path else self.config.XLSM_FILE
        self.error_handler = ErrorHandler(self.xlsm_path)
        self.selected_modules = selected_modules  # None = 同步所有模块
        self.force = force
        self.auto_close = auto_close

    def _backup_xlsm(self):
        """备份 xlsm 文件"""
        return self.error_handler.backup_xlsm()

    def _rollback(self):
        """回滚到备份版本"""
        return self.error_handler.rollback()

    def escape_xml(self, text):
        """转义 XML 特殊字符"""
        return (text
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&apos;')
                .replace('\n', '&#x0A;')
                .replace('\r', '&#x0D;'))

    def read_js_file(self, js_path):
        """读取 JS 文件内容"""
        with open(js_path, 'r', encoding='utf-8') as f:
            return f.read()

    def create_jde_data_bin(self, modules):
        """创建 JDEData.bin XML 内容"""
        module_xmls = []
        for i, module in enumerate(modules):
            name = module['name']
            module_id = module['id']
            code = module['code']
            escaped_code = self.escape_xml(code)

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

    def sync(self, backup=True):
        """同步代码到 xlsm 文件"""
        self.config.ensure_dirs()

        # 确定要同步的模块
        modules_to_sync = self.config.MODULES
        if self.selected_modules:
            modules_to_sync = [m for m in self.config.MODULES if m['id'] in self.selected_modules]
            print(f"🔄 同步指定模块: {[m['name'] for m in modules_to_sync]}")

        print(f"🔄 开始同步代码...")
        print(f"   目标文件: {self.xlsm_path.name}")
        if self.xlsm_path != self.config.XLSM_FILE:
            print(f"   原路径: {self.config.XLSM_FILE}")

        # 检查文件
        if not self.xlsm_path.exists():
            print(f"  ❌ 文件不存在: {self.xlsm_path}")
            return False

        # 热同步：检测文件是否被占用
        is_locked = FileLockDetector.is_file_locked(self.xlsm_path)
        wps_running = FileLockDetector.is_wps_running()

        if is_locked:
            print(f"  ⚠️  文件已被占用（可能被 WPS 打开）")

            if self.auto_close:
                print(f"  🔧 尝试自动关闭 WPS...")
                if FileLockDetector.close_wps():
                    import time
                    time.sleep(1)  # 等待 WPS 完全关闭
                    print(f"  ✅ WPS 已关闭")
                else:
                    print(f"  ❌ 无法关闭 WPS，请手动关闭后重试")
                    return False
            elif not self.force:
                print(f"""
  ═══════════════════════════════════════════════════════════════
  ⚠️  无法同步：目标文件正在被使用

  解决方案:
    1. 关闭 WPS 后重新运行同步命令
    2. 使用 --force 强制同步（可能丢失未保存的更改）
    3. 使用 --auto-close 自动关闭 WPS

  示例:
    jsa_tools.py sync --force
    jsa_tools.py sync --auto-close
  ═══════════════════════════════════════════════════════════════
""")
                return False
            else:
                print(f"  🔥 强制同步模式：尝试继续...")
                print(f"  ⚠️  警告：如果文件被占用，同步可能失败或导致数据丢失")
        elif wps_running:
            print(f"  ℹ️  WPS 正在运行，但目标文件未被占用")
            print(f"  💡 提示：如果 WPS 打开了此文件，请先关闭或使用 --auto-close")

        # 备份
        if backup:
            self._backup_xlsm()

        try:
            # 读取模块
            modules = []
            total_size = 0
            for mod_config in modules_to_sync:
                js_path = self.config.BASE_DIR / mod_config['file']
                if not js_path.exists():
                    print(f"  ⚠️  跳过不存在: {mod_config['file']}")
                    continue

                print(f"  📖 读取: {mod_config['file']}")
                code = self.read_js_file(js_path)
                modules.append({
                    'name': mod_config['name'],
                    'id': mod_config['id'],
                    'code': code
                })
                print(f"     {len(code.splitlines())} 行, {len(code)} 字节")
                total_size += len(code)

            # 创建临时目录
            temp_dir = self.config.TEMP_DIR / "sync"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir()

            # 解压 xlsm
            print(f"  📦 解压 xlsm 文件...")
            with zipfile.ZipFile(self.xlsm_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # 创建 JDEData.bin
            print(f"  ✏️  创建 JDEData.bin...")
            jde_data = self.create_jde_data_bin(modules)

            # 写入 JDEData.bin
            jde_path = temp_dir / 'xl' / 'JDEData.bin'
            jde_path.parent.mkdir(exist_ok=True)
            with open(jde_path, 'wb') as f:
                f.write(jde_data)

            # 重新打包
            print(f"  📦 重新打包 xlsm...")
            output_xlsm = temp_dir / "output.xlsm"

            with zipfile.ZipFile(output_xlsm, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = Path(root) / file
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)

            # 替换原文件
            shutil.move(output_xlsm, self.xlsm_path)

            print(f"  ✅ 同步完成! 共 {len(modules)} 个模块, {total_size} 字节")
            return True

        except Exception as e:
            print(f"  ❌ 同步失败: {e}")
            if backup:
                print(f"  🔄 正在回滚...")
                self._rollback()
            return False

        finally:
            # 清理临时目录
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

# ==================== 测试管理 ====================

class TestManager:
    """测试管理器"""

    def __init__(self):
        self.config = Config

    def open_wps(self):
        """打开 WPS 文件"""
        print(f"📂 打开 WPS 文件: {self.config.XLSM_FILE.name}")
        try:
            subprocess.run(["open", str(self.config.XLSM_FILE)], check=True)
            print("  ✅ 文件已打开")
            return True
        except Exception as e:
            print(f"  ❌ 打开失败: {e}")
            return False

    def show_test_instructions(self):
        """显示测试说明"""
        print(f"""
╔═══════════════════════════════════════════════════════════════╗
║                    测试执行说明                                  ║
╚═══════════════════════════════════════════════════════════════╝

📋 已同步模块:
""")
        for mod in self.config.MODULES:
            print(f"   ✓ {mod['name']} (ID: {mod['id']})")

        print(f"""
📝 运行测试步骤:

   1. WPS 已打开文件

   2. 打开 JSA 编辑器:
      按 Alt+F11 (Windows) 或 Option+F11 (Mac)

   3. 找到模块 "SuperPivotWPS" (ID: 4)

   4. 在立即窗口运行测试:
      运行所有测试()
      或
      测试12_组织3行1列()

══════════════════════════════════════════════════════════════
""")

# ==================== 主命令 ====================

def cmd_sync(args):
    """同步命令"""
    # 解析模块 ID
    selected_modules = None
    if args.modules:
        try:
            selected_modules = [int(x.strip()) for x in args.modules.split(',')]

            # 验证模块 ID 是否有效
            valid_ids = {m['id'] for m in Config.MODULES}
            invalid_ids = set(selected_modules) - valid_ids
            if invalid_ids:
                print(f"  ❌ 无效的模块 ID: {', '.join(map(str, invalid_ids))}")
                print(f"  可用模块: {', '.join(str(m['id']) for m in Config.MODULES)}")
                print(f"  可用模块名称: {', '.join(m['name'] + '(' + str(m['id']) + ')' for m in Config.MODULES)}")
                return 1

        except ValueError:
            print(f"  ❌ 无效的模块 ID 格式: {args.modules}")
            print(f"  请使用逗号分隔的数字，如: 1,3,4")
            print(f"  可用模块: {', '.join(str(m['id']) for m in Config.MODULES)}")
            return 1

    xlsm_path = args.file if args.file else None

    sync_manager = SyncManager(
        xlsm_path=xlsm_path,
        selected_modules=selected_modules,
        force=args.force,
        auto_close=args.auto_close
    )
    success = sync_manager.sync(backup=not args.no_backup)
    return 0 if success else 1

def cmd_test(args):
    """测试命令"""
    test_manager = TestManager()
    test_manager.open_wps()
    if not args.no_info:
        test_manager.show_test_instructions()
    return 0

def cmd_version(args):
    """版本命令"""
    version_manager = VersionManager()
    version = version_manager.get_version()
    print(f"JSA880 版本: {version}")

    if args.verbose:
        print("\n" + "="*60)
        print("更新日志:")
        print("="*60)
        changelog = version_manager.get_changelog()
        print(changelog)

    return 0

def cmd_status(args):
    """状态命令"""
    print(f"╔═══════════════════════════════════════════════════════════════╗")
    print(f"║                    JSA880 状态                                   ║")
    print(f"╚═══════════════════════════════════════════════════════════════╝")
    print()

    # 版本信息
    version_manager = VersionManager()
    version = version_manager.get_version()
    print(f"📌 版本: {version}")
    print()

    # 文件状态
    print(f"📄 文件状态:")
    for mod in Config.MODULES:
        js_path = Config.BASE_DIR / mod['file']
        status = "✅" if js_path.exists() else "❌"
        print(f"   {status} {mod['name']}: {mod['file']}")
    print()

    # xlsm 文件
    xlsm_status = "✅" if Config.XLSM_FILE.exists() else "❌"
    print(f"   {xlsm_status} 目标文件: {Config.XLSM_FILE.name}")
    print()

    # 备份状态
    backups = list(Config.BACKUP_DIR.glob("*.xlsm*"))
    print(f"📦 备份文件: {len(backups)} 个")
    if backups and args.verbose:
        for backup in sorted(backups)[-3:]:
            print(f"      - {backup.name}")
    print()

    return 0

def cmd_clean(args):
    """清理命令"""
    print("🧹 清理临时文件...")

    # 清理临时目录
    if Config.TEMP_DIR.exists():
        size = sum(f.stat().st_size for f in Config.TEMP_DIR.rglob('*') if f.is_file())
        shutil.rmtree(Config.TEMP_DIR)
        print(f"  ✅ 已清理临时目录: {size} 字节")

    # 清理旧备份 (保留最近 5 个)
    backups = sorted(Config.BACKUP_DIR.glob("*.xlsm*"), reverse=True)
    for old_backup in backups[5:]:
        old_backup.unlink()
        print(f"  🗑️  已删除旧备份: {old_backup.name}")

    print(f"  ✅ 清理完成!")
    return 0

def cmd_log(args):
    """日志命令 - 读取和管理 WPS 立即窗口日志"""
    log_reader = WPSLogReader()

    if args.clear:
        # 清空日志
        result = log_reader.clear_log()
        print(result)
        return 0

    # 新增：--clipboard 从剪贴板读取（推荐）
    if args.clipboard:
        print("📋 从剪贴板读取日志")
        print("="*60)
        import subprocess
        try:
            result = subprocess.run(
                ['pbpaste'],
                capture_output=True,
                text=True,
                timeout=5
            )
            content = result.stdout

            if content and content.strip():
                # 保存日志
                log_reader.config.ensure_dirs()
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                with open(log_reader.log_file, 'w', encoding='utf-8') as f:
                    f.write('='*80 + '\n')
                    f.write(f'📅 WPS 日志 (剪贴板) - {timestamp}\n')
                    f.write('='*80 + '\n')
                    f.write(content)
                    if not content.endswith('\n'):
                        f.write('\n')
                    f.write('='*80 + '\n')

                # 分析并显示
                lines = content.count('\n') + 1
                chars = len(content)
                ok_count = content.count('✅')
                warn_count = content.count('⚠️')
                error_count = content.count('❌')

                print()
                print(f"  📊 日志统计:")
                print(f"     📏 大小: {chars} 字符, {lines} 行")
                print(f"     ✅ 成功: {ok_count}")
                print(f"     ⚠️  警告: {warn_count}")
                print(f"     ❌ 错误: {error_count}")
                print()
                print(f"  💾 已保存到: {log_reader.log_file}")
                print()
                print("  📄 日志内容 (前 20 行):")
                print('  ' + '-'*60)
                content_lines = content.split('\n')
                for line in content_lines[:20]:
                    print(f"  {line}")
                if len(content_lines) > 20:
                    print(f"  ... 还有 {len(content_lines) - 20} 行")
                print('  ' + '-'*60)
                return 0
            else:
                print("  ❌ 剪贴板为空")
                print("  💡 请在 WPS 中复制日志 (Cmd+A, Cmd+C)")
                return 1
        except Exception as e:
            print(f"  ❌ 读取失败: {e}")
            return 1

    # 新增：--listen 启动剪贴板监听器
    if args.listen:
        print("👂 启动剪贴板监听器")
        print("="*60)
        import subprocess
        listener_path = log_reader.config.BASE_DIR / 'tools' / 'clipboard_listener.py'

        if not listener_path.exists():
            print(f"  ❌ 监听器脚本不存在: {listener_path}")
            return 1

        print(f"  💡 启动: python3 {listener_path} --interval {args.interval}")
        print()
        print("  使用方法:")
        print("  1. 监听器会在后台运行")
        print("  2. 在 WPS 中复制日志 (Cmd+A, Cmd+C)")
        print("  3. 监听器会自动检测并保存")
        print("  4. 按 Ctrl+C 停止监听")
        print()

        try:
            subprocess.run(
                ['python3', str(listener_path), '--interval', str(args.interval)],
                check=False,
                cwd=str(log_reader.config.BASE_DIR)
            )
            return 0
        except KeyboardInterrupt:
            print("\n  ⏹️  监听已停止")
            return 0
        except Exception as e:
            print(f"  ❌ 启动失败: {e}")
            return 1

    # 新增：--v2 使用改进版 pyautogui
    if args.v2:
        print("🤖 改进版 pyautogui 自动读取")
        print("="*60)
        import subprocess
        v2_reader_path = log_reader.config.BASE_DIR / 'tools' / 'wps_auto_reader_v2.py'

        if not v2_reader_path.exists():
            print(f"  ❌ 脚本不存在: {v2_reader_path}")
            return 1

        try:
            result = subprocess.run(
                ['python3', str(v2_reader_path)],
                capture_output=True,
                text=True,
                timeout=60,
                cwd=str(log_reader.config.BASE_DIR)
            )
            print(result.stdout)
            if result.stderr:
                print("  错误:", result.stderr)
            return 0 if result.returncode == 0 else 1
        except subprocess.TimeoutExpired:
            print("  ❌ 操作超时")
            return 1
        except Exception as e:
            print(f"  ❌ 执行失败: {e}")
            return 1

    # 新增：--saved 读取已保存的日志
    if args.saved:
        print("📄 读取已保存的日志")
        print("="*60)

        if not log_reader.log_file.exists():
            print(f"  ❌ 日志文件不存在: {log_reader.log_file}")
            return 1

        try:
            with open(log_reader.log_file, 'r', encoding='utf-8') as f:
                content = f.read()

            lines = content.count('\n') + 1
            chars = len(content)
            ok_count = content.count('✅')
            warn_count = content.count('⚠️')
            error_count = content.count('❌')

            print()
            print(f"  📊 日志统计:")
            print(f"     📏 大小: {chars} 字符, {lines} 行")
            print(f"     ✅ 成功: {ok_count}")
            print(f"     ⚠️  警告: {warn_count}")
            print(f"     ❌ 错误: {error_count}")
            print()
            print("  📄 日志内容:")
            print('  ' + '-'*60)
            content_lines = content.split('\n')
            for line in content_lines[:30]:
                print(f"  {line}")
            if len(content_lines) > 30:
                print(f"  ... 还有 {len(content_lines) - 30} 行")
            print('  ' + '-'*60)
            return 0
        except Exception as e:
            print(f"  ❌ 读取失败: {e}")
            return 1

    if args.paste:
        # 粘贴模式：等待用户从剪贴板粘贴日志
        print("📋 粘贴模式：请粘贴 WPS 立即窗口的日志内容")
        print("")
        print("使用方法:")
        print("1. 在 WPS 宏编辑器立即窗口中按 Cmd+A (macOS) 或 Ctrl+A (Windows)")
        print("2. 按 Cmd+C 或 Ctrl+C 复制内容")
        print("3. 在下方粘贴日志内容")
        print("4. 粘贴完成后按 Ctrl+D (Unix) 或 Ctrl+Z (Windows) 结束输入")
        print("")
        print("="*60)

        # 读取多行输入直到 EOF
        import sys
        lines = []
        try:
            for line in sys.stdin:
                lines.append(line.rstrip('\n'))
        except KeyboardInterrupt:
            pass

        if lines:
            content = '\n'.join(lines)
            # 保存到日志文件
            log_reader.config.ensure_dirs()
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open(log_reader.log_file, 'w', encoding='utf-8') as f:
                f.write(f"{'='*80}\n")
                f.write(f"📅 粘贴的日志 - {timestamp}\n")
                f.write(f"{'='*80}\n")
                f.write(content)
                f.write("\n" + "="*80 + "\n")

            print("")
            print("="*60)
            print(f"✅ 已保存 {len(lines)} 行日志到: {log_reader.log_file.name}")
            print(f"   使用 --show 查看日志")
            print("="*60)
        else:
            print("⚠️  未检测到输入内容")

        return 0

    if args.show:
        # 显示已保存的日志
        lines = args.lines if args.lines else None
        content = log_reader.get_saved_log(lines)

        if args.format == 'cl':  # Claude Code 格式
            print("📖 日志内容 (供 Claude Code 分析):")
            print("")
            print("```")
            print(content)
            print("```")
        else:
            print(content)

        return 0

    if args.analyze:
        # 分析日志
        content = log_reader.get_saved_log()
        if "暂无已保存的日志" in content:
            print("ℹ️  暂无日志可供分析")
            print("   提示: 先使用 log 命令或 log --paste 获取日志")
        else:
            print("🔍 日志分析:")
            print("")
            _analyze_log(content)
        return 0

    # 曲线救国方案：--cell 从单元格读取
    if args.cell:
        print("📊 从单元格读取日志（曲线救国方案）")
        print("="*60)
        import subprocess
        cell_reader_path = log_reader.config.BASE_DIR / 'tools' / 'cell_log_reader.py'

        if not cell_reader_path.exists():
            print(f"  ❌ 读取器脚本不存在: {cell_reader_path}")
            return 1

        # 构建命令
        cmd = ['python3', str(cell_reader_path), '--cell', '--sheet', args.sheet_name]

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=30,
                cwd=str(log_reader.config.BASE_DIR)
            )
            print(result.stdout)
            if result.stderr:
                print("  错误:", result.stderr)
            return 0 if result.returncode == 0 else 1
        except subprocess.TimeoutExpired:
            print("  ❌ 操作超时")
            return 1
        except Exception as e:
            print(f"  ❌ 执行失败: {e}")
            return 1

    # 曲线救国方案：--file-log 从文件读取
    if args.file_log:
        print("📄 从文件读取日志（曲线救国方案）")
        print("="*60)
        import subprocess
        cell_reader_path = log_reader.config.BASE_DIR / 'tools' / 'cell_log_reader.py'

        if not cell_reader_path.exists():
            print(f"  ❌ 读取器脚本不存在: {cell_reader_path}")
            return 1

        # 构建命令
        cmd = ['python3', str(cell_reader_path), '--txt']

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=30,
                cwd=str(log_reader.config.BASE_DIR)
            )
            print(result.stdout)
            if result.stderr:
                print("  错误:", result.stderr)
            return 0 if result.returncode == 0 else 1
        except subprocess.TimeoutExpired:
            print("  ❌ 操作超时")
            return 1
        except Exception as e:
            print(f"  ❌ 执行失败: {e}")
            return 1

    # 默认：显示使用提示
    print("📖 JSA880 日志工具")
    print()
    print("可用命令:")
    print("  ./jsa log --clipboard   # 📋 从剪贴板读取（推荐）")
    print("  ./jsa log --listen      # 👂 启动剪贴板监听器")
    print("  ./jsa log --cell        # 📊 从单元格读取（曲线救国）")
    print("  ./jsa log --file-log    # 📄 从文件读取（曲线救国）")
    print("  ./jsa log --v2          # 🤖 改进版 pyautogui")
    print("  ./jsa log --saved       # 📄 读取已保存的日志")
    print("  ./jsa log --paste       # 📝 手动粘贴")
    print("  ./jsa log --show        # 👁️  查看日志")
    print("  ./jsa log --analyze     # 📊 分析日志")
    print()
    print("💡 提示: 使用 --help 查看所有选项")

    # 如果指定了 --pyautogui，使用 pyautogui 自动读取
    if args.pyautogui:
        print()
        print("🤖 pyautogui 自动读取模式")
        print("="*60)

        # 调用 pyautogui 脚本
        import subprocess
        auto_reader_path = log_reader.config.BASE_DIR / 'tools' / 'wps_log_reader_pyautogui.py'

        if not auto_reader_path.exists():
            print("  ❌ pyautogui 脚本不存在")
            return 1

        # 构建命令
        cmd = ['python3', str(auto_reader_path)]
        if args.retry:
            cmd.extend(['--retry', str(args.retry)])

        # 执行
        try:
            result = subprocess.run(cmd, capture_output=True, text=True,
                                   timeout=60, cwd=str(log_reader.config.BASE_DIR))
            print(result.stdout)
            if result.stderr:
                print("  错误:", result.stderr)
            return 0 if result.returncode == 0 else 1
        except subprocess.TimeoutExpired:
            print("  ❌ 操作超时")
            return 1
        except Exception as e:
            print(f"  ❌ 执行失败: {e}")
            return 1

    # 如果指定了 --auto，尝试自动读取
    if args.auto:
        print()
        print("🤖 尝试自动读取...")
        result = log_reader.read_immediate_window(auto_mode=True)
        print(result)

    return 0

def _analyze_log(log_content):
    """分析日志内容"""
    lines = log_content.split('\n')

    # 统计
    error_count = log_content.count('❌')
    warn_count = log_content.count('⚠️')
    ok_count = log_content.count('✅')
    total_lines = len([l for l in lines if l.strip()])

    print(f"   总行数: {total_lines}")
    print(f"   ✅ 成功: {ok_count}")
    print(f"   ⚠️  警告: {warn_count}")
    print(f"   ❌ 错误: {error_count}")

    # 查找错误模式
    error_lines = [l for l in lines if '❌' in l or 'ERROR' in l or '错误' in l]
    if error_lines:
        print("\n   发现的问题:")
        for line in error_lines[:5]:  # 显示前 5 个
            print(f"      {line.strip()}")
        if len(error_lines) > 5:
            print(f"      ... 还有 {len(error_lines) - 5} 个问题")

    return 0

def cmd_workflow(args):
    """工作流命令 - 测试自动化"""
    workflow_manager = WorkflowManager(
        xlsm_path=args.file,
        selected_modules=None  # 工作流默认同步所有模块
    )

    # 确定模式
    mode = args.mode if args.mode else 'auto'

    # 执行工作流
    success = workflow_manager.run_workflow(
        mode=mode,
        test_function=args.test,
        wait_time=args.wait
    )

    return 0 if success else 1

def cmd_getlog(args):
    """自动获取日志命令 - 智能获取最新日志"""
    import subprocess

    auto_getter_path = Config.BASE_DIR / 'tools' / 'auto_get_log.py'

    if not auto_getter_path.exists():
        print(f"❌ 自动获取脚本不存在: {auto_getter_path}")
        return 1

    # 构建命令
    cmd = ['python3', str(auto_getter_path)]
    if args.claude:
        cmd.append('--claude')
    if args.save:
        cmd.append('--save')
    if args.quiet:
        cmd.append('--quiet')

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30,
            cwd=str(Config.BASE_DIR)
        )
        print(result.stdout)
        if result.stderr:
            print("错误:", result.stderr)
        return 0 if result.returncode == 0 else 1
    except subprocess.TimeoutExpired:
        print("❌ 操作超时")
        return 1
    except Exception as e:
        print(f"❌ 执行失败: {e}")
        return 1

def cmd_watch(args):
    """文件监视和自动同步命令"""
    import time

    # 收集要监视的文件
    watch_files = []
    for module in Config.MODULES:
        file_path = Config.BASE_DIR / module['file']
        if file_path.exists():
            watch_files.append({
                'path': file_path,
                'name': module['name'],
                'mtime': file_path.stat().st_mtime
            })

    if not watch_files:
        print("❌ 没有找到可监视的文件")
        return 1

    print("=" * 60)
    print("👀 文件监视模式")
    print("=" * 60)
    print(f"监视 {len(watch_files)} 个文件:")
    for f in watch_files:
        print(f"  - {f['name']}: {f['path'].relative_to(Config.BASE_DIR)}")
    print(f"检查间隔: {args.interval} 秒")
    if args.auto_close:
        print("自动关闭 WPS: 启用")
    if args.reopen:
        print("同步后重新打开 WPS: 启用")
    print()
    print("提示: 按 Ctrl+C 停止监视")
    print("=" * 60)
    print()

    sync_count = 0
    try:
        while True:
            time.sleep(args.interval)
            changed = False
            changed_files = []

            # 检查文件变化
            for f in watch_files:
                current_mtime = f['path'].stat().st_mtime
                if current_mtime != f['mtime']:
                    f['mtime'] = current_mtime
                    changed = True
                    changed_files.append(f['name'])

            if changed:
                sync_count += 1
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                print()
                print(f"🔄 [{timestamp}] 检测到文件变化: {', '.join(changed_files)}")

                # 执行同步
                sync_tool = SyncManager(Config.XLSM_FILE, auto_close=args.auto_close)
                result = sync_tool.sync(backup=(sync_count == 1))  # 只在第一次同步时备份

                if result:
                    print(f"✅ 同步成功 (第 {sync_count} 次)")

                    # 如果需要重新打开 WPS
                    if args.reopen:
                        time.sleep(0.5)
                        automation = WPSAutomation(Config.XLSM_FILE)
                        open_result = automation.open_wps()
                        if open_result:
                            print("📂 已重新打开 WPS 文件")
                        else:
                            print("⚠️  重新打开 WPS 失败")
                else:
                    print(f"❌ 同步失败")

                if not args.verbose:
                    print("⏳ 继续监视...")

    except KeyboardInterrupt:
        print()
        print("=" * 60)
        print(f"✅ 监视已停止")
        print(f"总同步次数: {sync_count}")
        print("=" * 60)
        return 0

# ==================== 主函数 ====================

def main():
    parser = argparse.ArgumentParser(
        description='JSA880 统一工具脚本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例:
  jsa sync                                           # 同步所有模块到默认 xlsm
  jsa sync --file path/to/file.xlsm                  # 同步到指定文件
  jsa sync --modules 1,4                             # 只同步模块 1 和 4
  jsa sync --auto-close                              # 热同步：自动关闭 WPS 后同步
  jsa watch                                          # 监视文件变化并自动同步
  jsa watch --auto-close --reopen                    # 自动关闭 WPS → 同步 → 重新打开
  jsa watch --interval 1                             # 每 1 秒检查一次（更快响应）
  jsa workflow                                       # 完整自动化测试工作流
  jsa workflow --mode manual                         # 手动模式工作流
  jsa workflow --test 测试12_组织3行1列               # 运行指定测试
  jsa workflow --mode test --wait 15                 # 仅运行测试，等待15秒
  jsa test                                           # 打开文件并显示测试说明
  jsa log --show                                     # 显示已保存的日志
  jsa log --analyze                                  # 分析日志统计
  jsa getlog                                         # 自动获取最新日志（推荐）
  jsa getlog --claude                                # Claude Code 友好格式
  jsa version                                        # 显示版本
  jsa status                                         # 显示状态

工作流模式 (workflow):
  auto        完整自动化（同步 → 打开 → 运行测试 → 收集日志）
  manual      手动模式（同步 → 打开 → 等待手动运行测试）
  sync        仅同步代码
  test        仅运行测试（WPS 必须已打开）
  collect     仅收集日志

监视模式 (watch):
  实时监视 JS 源文件变化，自动同步到 xlsm 文件
  修改源文件后，无需手动运行 sync 命令
  配合 --auto-close 和 --reopen 可实现完全自动化

热同步选项:
  --auto-close    自动关闭 WPS 后同步（macOS/Windows 支持）
  --force        强制同步，忽略文件占用警告

可用模块 ID:
  1 - JSA880 (主框架)
  3 - TestDataGenerator (测试数据生成)
  4 - SuperPivotWPS (SuperPivot 测试套件)
  5 - PerformanceTest (性能测试)
  6 - CellLog (单元格日志模块)
  7 - FileLog (文件日志模块)
        '''
    )

    subparsers = parser.add_subparsers(dest='command', help='可用命令')

    # sync 命令
    parser_sync = subparsers.add_parser('sync', help='同步代码到 xlsm 文件')
    parser_sync.add_argument('--file', '-f', type=str, help='指定 xlsm 文件路径')
    parser_sync.add_argument('--modules', '-m', type=str, help='指定要同步的模块 ID (逗号分隔, 如: 1,3,4)')
    parser_sync.add_argument('--auto-close', action='store_true', help='热同步：自动关闭 WPS 后同步')
    parser_sync.add_argument('--force', action='store_true', help='强制同步，忽略文件占用警告')
    parser_sync.add_argument('--no-backup', action='store_true', help='不备份文件')
    parser_sync.set_defaults(func=cmd_sync)

    # test 命令
    parser_test = subparsers.add_parser('test', help='打开 WPS 并显示测试说明')
    parser_test.add_argument('--no-info', action='store_true', help='不显示测试说明')
    parser_test.set_defaults(func=cmd_test)

    # version 命令
    parser_version = subparsers.add_parser('version', help='显示版本信息')
    parser_version.add_argument('-v', '--verbose', action='store_true', help='显示更新日志')
    parser_version.set_defaults(func=cmd_version)

    # status 命令
    parser_status = subparsers.add_parser('status', help='显示状态信息')
    parser_status.add_argument('-v', '--verbose', action='store_true', help='详细信息')
    parser_status.set_defaults(func=cmd_status)

    # clean 命令
    parser_clean = subparsers.add_parser('clean', help='清理临时文件')
    parser_clean.set_defaults(func=cmd_clean)

    # log 命令
    parser_log = subparsers.add_parser('log', help='读取和管理 WPS 立即窗口日志')
    parser_log.add_argument('--show', '-s', action='store_true', help='显示已保存的日志')
    parser_log.add_argument('--lines', '-n', type=int, help='显示最近的 N 行日志 (配合 --show 使用)')
    parser_log.add_argument('--paste', '-p', action='store_true', help='粘贴模式：从剪贴板粘贴日志内容')
    parser_log.add_argument('--clear', '-c', action='store_true', help='清空已保存的日志')
    parser_log.add_argument('--analyze', '-a', action='store_true', help='分析已保存的日志')
    parser_log.add_argument('--auto', action='store_true', help='自动模式：尝试自动读取 WPS 立即窗口（需要 WPS 运行）')
    parser_log.add_argument('--pyautogui', action='store_true', help='使用 pyautogui 自动读取（最可靠的自动方案）')
    parser_log.add_argument('--retry', '-r', type=int, default=2, help='pyautogui 模式下的重试次数（默认: 2）')
    parser_log.add_argument('--format', '-f', choices=['text', 'cl'], default='text', help='输出格式 (text=纯文本, cl=Claude Code 代码块)')
    # 新增选项
    parser_log.add_argument('--clipboard', '-C', action='store_true', help='从剪贴板读取日志（推荐，最可靠）')
    parser_log.add_argument('--listen', '-l', action='store_true', help='启动剪贴板监听器（自动检测复制）')
    parser_log.add_argument('--interval', '-i', type=float, default=2.0, help='监听器检查间隔秒数（默认: 2.0）')
    parser_log.add_argument('--v2', action='store_true', help='使用改进版 pyautogui 读取器')
    parser_log.add_argument('--saved', action='store_true', help='读取已保存的日志文件')
    # 曲线救国方案
    parser_log.add_argument('--cell', action='store_true', help='从单元格读取日志（曲线救国方案）')
    parser_log.add_argument('--file-log', '-F', action='store_true', help='从文件读取日志（曲线救国方案）')
    parser_log.add_argument('--sheet-name', type=str, default='日志', help='单元格日志的工作表名（默认: 日志）')
    parser_log.set_defaults(func=cmd_log)

    # workflow 命令
    parser_workflow = subparsers.add_parser('workflow', help='测试自动化工作流')
    parser_workflow.add_argument('--file', '-f', type=str, help='指定 xlsm 文件路径')
    parser_workflow.add_argument('--mode', '-m', choices=['auto', 'sync', 'test', 'collect', 'manual'],
                                 help='工作流模式: auto=完整自动化, sync=仅同步, test=仅测试, collect=仅收集日志, manual=手动模式')
    parser_workflow.add_argument('--test', '-t', type=str, default='运行所有测试', help='测试函数名 (默认: 运行所有测试)')
    parser_workflow.add_argument('--wait', '-w', type=int, default=10, help='等待测试完成的秒数 (默认: 10)')
    parser_workflow.set_defaults(func=cmd_workflow)

    # getlog 命令 - 自动获取日志
    parser_getlog = subparsers.add_parser('getlog', help='自动获取最新的 WPS 日志（供 Claude Code 使用）')
    parser_getlog.add_argument('--claude', '-c', action='store_true', help='输出 Claude Code 友好格式（带代码块）')
    parser_getlog.add_argument('--save', '-s', action='store_true', help='保存到文件')
    parser_getlog.add_argument('--quiet', '-q', action='store_true', help='安静模式，只输出日志内容')
    parser_getlog.set_defaults(func=cmd_getlog)

    # watch 命令 - 文件监视和自动同步
    parser_watch = subparsers.add_parser('watch', help='监视文件变化并自动同步到 xlsm')
    parser_watch.add_argument('--interval', '-i', type=int, default=2, help='检查间隔（秒），默认 2 秒')
    parser_watch.add_argument('--auto-close', '-a', action='store_true', help='自动关闭 WPS 后同步')
    parser_watch.add_argument('--reopen', '-r', action='store_true', help='同步后自动重新打开 WPS')
    parser_watch.add_argument('--verbose', '-v', action='store_true', help='显示详细输出')
    parser_watch.set_defaults(func=cmd_watch)

    # 解析参数
    args = parser.parse_args()

    # 执行命令
    if hasattr(args, 'func'):
        sys.exit(args.func(args))
    else:
        parser.print_help()
        sys.exit(1)

if __name__ == "__main__":
    main()
