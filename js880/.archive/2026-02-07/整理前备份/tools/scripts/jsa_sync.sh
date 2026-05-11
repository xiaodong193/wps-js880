#!/bin/bash
# JSA880 自动同步脚本 (Bash 版本)
# 功能：将 JSA880.js 自动同步到 xlsm 文件

set -e

# 配置路径
SOURCE_JS="/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/JSA880.js"
TARGET_XLSM="/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm"
WORK_DIR="/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880"

echo "🔄 JSA880 自动同步工具"
echo "========================="
echo "源文件: $SOURCE_JS"
echo "目标文件: $TARGET_XLSM"
echo ""

# 检查 Python 脚本是否存在
PYTHON_SCRIPT="$WORK_DIR/jsa_sync.py"
if [ ! -f "$PYTHON_SCRIPT" ]; then
    echo "❌ 错误: Python 脚本不存在: $PYTHON_SCRIPT"
    exit 1
fi

# 执行 Python 同步脚本
python3 "$PYTHON_SCRIPT" "$SOURCE_JS" "$TARGET_XLSM"

echo ""
echo "✅ 同步完成!"
