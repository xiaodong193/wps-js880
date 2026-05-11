#!/bin/bash
# JSA 自动化测试 - Claude 一键测试
# 功能：同步代码 -> 打开 WPS -> 读取测试结果

WORK_DIR="/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880"
TEST_RESULT_FILE="/tmp/jsa_test_result.txt"
TEST_ERROR_FILE="/tmp/jsa_test_error.txt"

echo "========================================"
echo "  JSA 自动化测试 - Claude 一键测试"
echo "========================================"
echo ""

# 检查是否有测试结果
if [ -f "$TEST_RESULT_FILE" ] && [ -s "$TEST_RESULT_FILE" ]; then
    echo "📄 发现已有的测试结果"
    echo ""
    echo "========================================"
    echo "  测试结果"
    echo "========================================"
    echo ""
    cat "$TEST_RESULT_FILE"
    echo ""

    if [ -f "$TEST_ERROR_FILE" ] && [ -s "$TEST_ERROR_FILE" ]; then
        echo "========================================"
        echo "  错误信息"
        echo "========================================"
        echo ""
        cat "$TEST_ERROR_FILE"
        echo ""
    fi

    # 清空结果文件（下次重新测试）
    > "$TEST_RESULT_FILE"
    > "$TEST_ERROR_FILE"

    echo "💡 提示: 结果文件已清空，下次将运行新测试"
    exit 0
fi

# 没有结果，执行测试流程
echo "步骤1: 同步代码到 xlsm..."
python3 "$WORK_DIR/jsa_sync.py"
if [ $? -ne 0 ]; then
    echo "❌ 同步失败"
    exit 1
fi
echo "✅ 同步完成"
echo ""

echo "步骤2: 打开 WPS..."
open -a "wpsoffice" "/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/JS880教案/第03章/3-25/3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm"
echo "✅ WPS 已打开"
echo ""

echo "========================================"
echo "  请在 WPS 中执行测试"
echo "========================================"
echo ""
echo "在宏编辑器中运行以下函数之一:"
echo ""
echo "  ClaudeAutoTest()      - 快速测试"
echo "  自动测试_完整版()      - 完整测试"
echo ""
echo "然后再次运行此脚本读取结果"
echo ""
echo "========================================"
