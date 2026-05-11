#!/bin/bash
# =============================================================================
# JSA880 项目文件整理脚本
# =============================================================================
# 使用方法:
#   1. 确保在 js880 项目根目录运行此脚本
#   2. chmod +x 整理项目文件.sh
#   3. ./整理项目文件.sh
# =============================================================================

echo "╔════════════════════════════════════════════════════════════╗"
echo "║     JSA880 项目文件整理脚本                                ║"
echo "║     版本: 1.0                                              ║"
echo "╚════════════════════════════════════════════════════════════╝"
echo ""

# 设置项目根目录
PROJECT_ROOT="/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880"
cd "$PROJECT_ROOT" || exit 1

# 检查是否在正确目录
if [ ! -f "JSA880.js" ]; then
    echo "❌ 错误: 未找到 JSA880.js，请在项目根目录运行此脚本"
    exit 1
fi

echo "✓ 检测到项目根目录"
echo ""

# =============================================================================
# 步骤1: 创建备份
# =============================================================================
echo "【步骤1】创建完整备份..."
BACKUP_DIR=".archive/$(date +%Y-%m-%d)/整理前备份"
mkdir -p "$BACKUP_DIR"
cp -r * "$BACKUP_DIR/" 2>/dev/null
echo "✓ 备份已创建: $BACKUP_DIR"
echo ""

# =============================================================================
# 步骤2: 创建新目录结构
# =============================================================================
echo "【步骤2】创建新目录结构..."

# 文档目录
mkdir -p docs/guides
mkdir -p docs/api
mkdir -p docs/development
mkdir -p docs/archive/$(date +%Y-%m)

# 源码目录
mkdir -p src/core
mkdir -p src/modules
mkdir -p src/tests
mkdir -p src/examples

# 分发目录
mkdir -p dist

echo "✓ 目录结构创建完成"
echo ""

# =============================================================================
# 步骤3: 移动根目录文档
# =============================================================================
echo "【步骤3】整理根目录文档..."

# 使用指南类
[ -f "superPivot_完整使用指南.md" ] && \
    mv "superPivot_完整使用指南.md" "docs/guides/03_superPivot使用指南.md" && \
    echo "  ✓ superPivot_完整使用指南.md → docs/guides/"

[ -f "superPivot_快速参考卡.md" ] && \
    mv "superPivot_快速参考卡.md" "docs/guides/快速参考卡.md" && \
    echo "  ✓ superPivot_快速参考卡.md → docs/guides/"

[ -f "superPivot_视频教程脚本.md" ] && \
    mv "superPivot_视频教程脚本.md" "docs/guides/视频教程脚本.md" && \
    echo "  ✓ superPivot_视频教程脚本.md → docs/guides/"

# 索引文件
[ -f "使用文档索引.md" ] && \
    cp "使用文档索引.md" "docs/index.md" && \
    echo "  ✓ 使用文档索引.md → docs/index.md"

# 归档文件
[ -f "代码检查报告与模块清单.md" ] && \
    mv "代码检查报告与模块清单.md" "docs/archive/$(date +%Y-%m)/代码检查报告_$(date +%Y%m%d).md" && \
    echo "  ✓ 代码检查报告与模块清单.md → docs/archive/"

[ -f "代码检查总结报告.md" ] && \
    mv "代码检查总结报告.md" "docs/archive/$(date +%Y-%m)/代码检查总结_$(date +%Y%m%d).md" && \
    echo "  ✓ 代码检查总结报告.md → docs/archive/"

echo ""

# =============================================================================
# 步骤4: 整理源码模块
# =============================================================================
echo "【步骤4】整理源码模块..."

# superPivot模块整理
if [ -f "src/modules/superPivot_v390规范版.js" ]; then
    cp "src/modules/superPivot_v390规范版.js" "src/modules/superPivot_v390.js"
    echo "  ✓ superPivot_v390规范版.js → superPivot_v390.js"
fi

# 移动测试文件
[ -f "src/modules/superPivot_一键测试.js" ] && \
    mv "src/modules/superPivot_一键测试.js" "src/examples/示例_一键测试.js" && \
    echo "  ✓ superPivot_一键测试.js → src/examples/"

[ -f "src/modules/superPivot_示例测试.js" ] && \
    mv "src/modules/superPivot_示例测试.js" "src/examples/示例_完整演示.js" && \
    echo "  ✓ superPivot_示例测试.js → src/examples/"

# 移动测试套件
[ -f "src/modules/superPivot_测试套件.js" ] && \
    mv "src/modules/superPivot_测试套件.js" "src/tests/test_superPivot.js" && \
    echo "  ✓ superPivot_测试套件.js → src/tests/"

# 归档旧版本
[ -f "src/modules/superPivot_WPS_测试.js" ] && \
    mv "src/modules/superPivot_WPS_测试.js" "docs/archive/$(date +%Y-%m)/" && \
    echo "  ✓ superPivot_WPS_测试.js → docs/archive/"

echo ""

# =============================================================================
# 步骤5: 创建版本分发
# =============================================================================
echo "【步骤5】创建版本分发..."

cp "JSA880.js" "dist/JSA880_v3.9.1.js"
echo "✓ 创建 dist/JSA880_v3.9.1.js"
echo ""

# =============================================================================
# 步骤6: 清理临时文件
# =============================================================================
echo "【步骤6】清理临时文件..."

# 删除重复文档说明文件
[ -f "src/modules/superPivot_测试说明.md" ] && rm "src/modules/superPivot_测试说明.md"
[ -f "src/modules/测试文件说明.md" ] && rm "src/modules/测试文件说明.md"
[ -f "src/modules/测试文件索引.md" ] && rm "src/modules/测试文件索引.md"

echo "✓ 清理完成"
echo ""

# =============================================================================
# 完成报告
# =============================================================================
echo "╔════════════════════════════════════════════════════════════╗"
echo "║     整理完成!                                              ║"
echo "╚════════════════════════════════════════════════════════════╝"
echo ""
echo "📁 新目录结构:"
echo "  docs/           - 文档中心"
echo "  ├── guides/     - 使用指南"
echo "  ├── api/        - API参考 (待创建)"
echo "  ├── examples/   - 示例代码"
echo "  └── archive/    - 归档文档"
echo ""
echo "  src/            - 源码目录"
echo "  ├── modules/    - 功能模块"
echo "  ├── tests/      - 测试套件"
echo "  └── examples/   - 示例代码"
echo ""
echo "  dist/           - 分发版本"
echo ""
echo "⚠️  请手动完成以下任务:"
echo "  1. 更新 README.md 中的链接"
echo "  2. 检查 docs/index.md 导航"
echo "  3. 验证代码运行正常"
echo "  4. 提交 git commit (如果有git)"
echo ""
echo "📦 备份位置: $BACKUP_DIR"
echo ""
