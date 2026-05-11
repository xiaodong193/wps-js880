# JSA880 同步工具使用与维护指南

> **版本**: 1.0
> **更新日期**: 2026年2月2日
> **工具名称**: jsa_tools.py
> **支持平台**: macOS, Windows, Linux

---

## 📋 目录

1. [功能概述](#功能概述)
2. [快速开始](#快速开始)
3. [命令详解](#命令详解)
4. [热同步功能](#热同步功能)
5. [模块管理](#模块管理)
6. [故障排除](#故障排除)
7. [最佳实践](#最佳实践)
8. [维护指南](#维护指南)

---

## 功能概述

### 核心功能

`jsa_tools.py` 是 JSA880 框架的统一工具脚本，提供以下功能：

| 功能 | 命令 | 说明 |
|------|------|------|
| **代码同步** | `sync` | 将 JS 模块同步到 xlsm 文件 |
| **测试执行** | `test` | 打开 WPS 并显示测试说明 |
| **版本管理** | `version` | 查看版本和更新日志 |
| **状态查询** | `status` | 显示文件和模块状态 |
| **清理维护** | `clean` | 清理临时文件和旧备份 |

### 主要特性

- ✅ **多模块支持**: 可选择同步特定模块
- ✅ **文件选择**: 支持指定目标 xlsm 文件
- ✅ **热同步**: 支持在 WPS 打开时自动关闭并同步
- ✅ **自动备份**: 同步前自动备份原文件
- ✅ **错误回滚**: 同步失败时自动回滚
- ✅ **跨平台**: 支持 macOS、Windows、Linux

---

## 快速开始

### 环境要求

```bash
# Python 版本
Python 3.6+

# 依赖库（均为标准库，无需额外安装）
- argparse
- zipfile
- shutil
- subprocess
- pathlib
- platform
- fcntl (Unix)
```

### 基本使用

```bash
# 进入工作目录
cd /Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880

# 查看帮助
python3 jsa_tools.py --help

# 查看同步命令帮助
python3 jsa_tools.py sync --help

# 同步所有模块到默认文件
python3 jsa_tools.py sync

# 同步后打开测试
python3 jsa_tools.py test
```

### 典型工作流程

```bash
# 1. 修改代码后同步
vim JSA880.js
python3 jsa_tools.py sync

# 2. 只同步主框架（快速更新）
python3 jsa_tools.py sync --modules 1

# 3. 同步并测试
python3 jsa_tools.py sync
python3 jsa_tools.py test

# 4. 在 WPS 打开时热同步
python3 jsa_tools.py sync --auto-close
```

---

## 命令详解

### sync - 同步命令

#### 语法

```bash
python3 jsa_tools.py sync [选项]
```

#### 选项

| 选项 | 简写 | 说明 | 示例 |
|------|------|------|------|
| `--file` | `-f` | 指定 xlsm 文件路径 | `-f /path/to/file.xlsm` |
| `--modules` | `-m` | 指定要同步的模块 ID | `-m 1,3,4` |
| `--auto-close` | | 自动关闭 WPS 后同步 | `--auto-close` |
| `--force` | | 强制同步，忽略占用警告 | `--force` |
| `--no-backup` | | 不备份文件 | `--no-backup` |

#### 使用场景

##### 场景 1: 同步所有模块

```bash
# 同步所有配置的模块到默认文件
python3 jsa_tools.py sync

# 输出:
# 🔄 开始同步代码...
#    目标文件: 3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm
#   📖 读取: JSA880.js
#      12695 行, 365792 字节
#   📖 读取: cls生成测试数据.js
#      466 行, 13078 字节
#   ...
#   ✅ 同步完成! 共 4 个模块, 399656 字节
```

##### 场景 2: 只同步特定模块

```bash
# 只同步主框架（开发时常用）
python3 jsa_tools.py sync --modules 1

# 只同步测试模块
python3 jsa_tools.py sync --modules 3,4,5

# 组合使用：指定文件和模块
python3 jsa_tools.py sync --file test.xlsm --modules 1,4
```

##### 场景 3: 同步到不同的 xlsm 文件

```bash
# 同步到项目文件
python3 jsa_tools.py sync --file /path/to/project.xlsm

# 同步到测试文件
python3 jsa_tools.py sync -f 测试副本.xlsm
```

##### 场景 4: 快速同步（不备份）

```bash
# 频繁开发时不备份，提高速度
python3 jsa_tools.py sync --modules 1 --no-backup
```

---

### test - 测试命令

#### 语法

```bash
python3 jsa_tools.py test [选项]
```

#### 选项

| 选项 | 说明 |
|------|------|
| `--no-info` | 不显示测试说明 |

#### 使用示例

```bash
# 打开 WPS 并显示测试说明
python3 jsa_tools.py test

# 只打开 WPS，不显示说明
python3 jsa_tools.py test --no-info
```

---

### version - 版本命令

#### 语法

```bash
python3 jsa_tools.py version [选项]
```

#### 选项

| 选项 | 简写 | 说明 |
|------|------|------|
| `--verbose` | `-v` | 显示更新日志 |

#### 使用示例

```bash
# 查看版本
python3 jsa_tools.py version
# 输出: JSA880 版本: 3.8.3

# 查看版本和更新日志
python3 jsa_tools.py version -v
```

---

### status - 状态命令

#### 语法

```bash
python3 jsa_tools.py status [选项]
```

#### 选项

| 选项 | 简写 | 说明 |
|------|------|------|
| `--verbose` | `-v` | 显示详细信息（包括备份列表） |

#### 使用示例

```bash
# 查看状态
python3 jsa_tools.py status

# 查看详细状态（包括备份文件）
python3 jsa_tools.py status -v
```

#### 输出示例

```
╔═══════════════════════════════════════════════════════════════╗
║                    JSA880 状态                                   ║
╚═══════════════════════════════════════════════════════════════╝

📌 版本: 3.8.3

📄 文件状态:
   ✅ JSA880: JSA880.js
   ✅ TestDataGenerator: cls生成测试数据.js
   ✅ SuperPivotWPS: superPivot_WPS_测试.js
   ✅ PerformanceTest: superPivot_性能测试.js

   ✅ 目标文件: 3.25 superPivot降维打击 实现多层表头交叉透视汇总_副本.xlsm

📦 备份文件: 2 个
```

---

### clean - 清理命令

#### 语法

```bash
python3 jsa_tools.py clean
```

#### 功能

- 清理临时目录 `.temp/`
- 删除旧备份文件（保留最近 5 个）

#### 使用示例

```bash
# 清理临时文件
python3 jsa_tools.py clean

# 输出:
# 🧹 清理临时文件...
#   ✅ 已清理临时目录: 0 字节
#   ✅ 清理完成!
```

---

## 热同步功能

### 功能说明

热同步允许在 WPS 打开目标文件时进行代码同步，无需手动关闭 WPS。

### 工作原理

```
┌─────────────────────────────────────────────────────────────┐
│                    热同步流程                                │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  1. 检测文件占用状态                                        │
│     ├─ 文件未被占用 → 直接同步                             │
│     └─ 文件被占用 → 进入处理流程                            │
│                                                             │
│  2. 文件被占用时的处理选项                                  │
│     ├─ --auto-close: 自动关闭 WPS                          │
│     ├─ --force: 强制尝试同步                               │
│     └─ 默认: 显示提示并退出                                 │
│                                                             │
│  3. 执行同步                                                │
│     ├─ 备份原文件                                          │
│     ├─ 解压 xlsm                                           │
│     ├─ 更新 JDEData.bin                                    │
│     └─ 重新打包                                            │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### 使用场景

#### 场景 1: 开发中的快速迭代

```bash
# WPS 打开文件时，自动关闭并同步最新代码
python3 jsa_tools.py sync --auto-close --no-backup
```

#### 场景 2: 多项目切换

```bash
# 为不同项目同步代码
python3 jsa_tools.py sync --auto-close --file 项目A.xlsm
# 测试项目A...
python3 jsa_tools.py sync --auto-close --file 项目B.xlsm
# 测试项目B...
```

#### 场景 3: 自动化脚本

```bash
# 在脚本中使用自动关闭
#!/bin/bash
while true; do
    python3 jsa_tools.py sync --auto-close --no-backup
    sleep 30  # 每30秒同步一次
done
```

### 热同步选项对比

| 选项 | 行为 | 适用场景 | 风险 |
|------|------|----------|------|
| 默认 | 文件被占用时退出 | 正常开发，手动关闭 WPS | 无 |
| `--auto-close` | 自动关闭 WPS | 快速迭代，自动化 | 可能丢失未保存更改 |
| `--force` | 强制尝试同步 | 确定文件未被实际占用 | 可能同步失败 |

### 平台支持

| 平台 | 自动关闭 WPS | 文件锁检测 | 备注 |
|------|-------------|-----------|------|
| macOS | ✅ AppleScript | ✅ fcntl | 完全支持 |
| Windows | ✅ taskkill | ✅ 文件打开测试 | 完全支持 |
| Linux | ❌ 不支持 | ✅ fcntl | 需手动关闭 WPS |

---

## 模块管理

### 可用模块

| ID | 名称 | 文件 | 行数 | 说明 |
|----|------|------|------|------|
| 1 | JSA880 | JSA880.js | 12,695 | 主框架，包含 Array2D 等核心类 |
| 3 | TestDataGenerator | cls生成测试数据.js | 466 | 测试数据生成器 |
| 4 | SuperPivotWPS | superPivot_WPS_测试.js | 480 | SuperPivot 测试套件 |
| 5 | PerformanceTest | superPivot_性能测试.js | 422 | 性能测试模块 |

### 模块配置

模块配置在 `jsa_tools.py` 的 `Config` 类中：

```python
class Config:
    # 模块配置 (仅包含实际存在的文件)
    MODULES = [
        {'name': 'JSA880', 'file': 'JSA880.js', 'id': 1},
        {'name': 'TestDataGenerator', 'file': 'cls生成测试数据.js', 'id': 3},
        {'name': 'SuperPivotWPS', 'file': 'superPivot_WPS_测试.js', 'id': 4},
        {'name': 'PerformanceTest', 'file': 'superPivot_性能测试.js', 'id': 5},
    ]
```

### 添加新模块

#### 步骤 1: 准备模块文件

```bash
# 创建新的 JS 模块文件
vim my_new_module.js
```

#### 步骤 2: 更新配置

编辑 `jsa_tools.py`，添加模块配置：

```python
MODULES = [
    # ... 现有模块
    {'name': 'MyNewModule', 'file': 'my_new_module.js', 'id': 6},
]
```

#### 步骤 3: 测试同步

```bash
# 验证新模块可以同步
python3 jsa_tools.py sync --modules 6
```

### 删除模块

```python
# 从配置中移除模块
MODULES = [
    {'name': 'JSA880', 'file': 'JSA880.js', 'id': 1},
    # 移除不需要的模块
]
```

---

## 故障排除

### 常见问题

#### 问题 1: 文件被占用

**症状**:
```
⚠️  文件已被占用（可能被 WPS 打开）
```

**解决方案**:

```bash
# 方案 A: 手动关闭 WPS 后同步
# 1. 关闭 WPS
# 2. 运行同步
python3 jsa_tools.py sync

# 方案 B: 使用自动关闭
python3 jsa_tools.py sync --auto-close

# 方案 C: 强制同步（不推荐）
python3 jsa_tools.py sync --force
```

---

#### 问题 2: 无效的模块 ID

**症状**:
```
❌ 无效的模块 ID: 99
可用模块: 1, 3, 4, 5
```

**解决方案**:

```bash
# 查看可用模块
python3 jsa_tools.py status

# 使用正确的模块 ID
python3 jsa_tools.py sync --modules 1,4
```

---

#### 问题 3: xlsm 文件不存在

**症状**:
```
❌ 文件不存在: /path/to/file.xlsm
```

**解决方案**:

```bash
# 方案 A: 使用默认文件（不指定 --file）
python3 jsa_tools.py sync

# 方案 B: 检查文件路径
ls -la /path/to/file.xlsm

# 方案 C: 更新默认文件路径
# 编辑 jsa_tools.py 中的 Config.XLSM_FILE
```

---

#### 问题 4: 同步失败

**症状**:
```
❌ 同步失败: [错误信息]
🔄 正在回滚...
↩️  已回滚到备份版本
```

**解决方案**:

```bash
# 1. 查看详细错误
python3 jsa_tools.py sync 2>&1 | tee sync_error.log

# 2. 检查模块文件是否存在
ls -la JSA880.js cls生成测试数据.js

# 3. 检查文件权限
chmod 644 *.js

# 4. 检查备份目录
mkdir -p .backups
```

---

#### 问题 5: WPS 无法自动关闭（macOS）

**症状**:
```
⚠️  无法自动关闭 WPS: ...
```

**解决方案**:

```bash
# 方案 A: 授予终端访问权限
# 系统偏好设置 -> 安全性与隐私 -> 隐私 -> 自动化
# 添加终端或 iTerm 的权限

# 方案 B: 手动关闭 WPS
# 使用 Cmd+Q 或 菜单 -> 退出 WPS

# 方案 C: 使用命令行强制关闭
pkill -9 'WPS Office'
```

---

### 调试技巧

#### 启用详细输出

```bash
# 查看详细状态
python3 jsa_tools.py status -v

# 查看版本和日志
python3 jsa_tools.py version -v
```

#### 检查模块文件

```bash
# 验证模块文件
ls -lh *.js

# 检查文件编码
file JSA880.js

# 检查文件语法
node --check JSA880.js  # 如果有 Node.js
```

#### 检查备份

```bash
# 查看备份文件
ls -la .backups/

# 恢复备份
cp .backups/file_backup_20260202.xlsm file.xlsm
```

---

## 最佳实践

### 开发工作流

#### 推荐流程

```bash
# 1. 编辑代码
vim JSA880.js

# 2. 快速同步（只同步修改的模块）
python3 jsa_tools.py sync --modules 1 --no-backup

# 3. 在 WPS 中测试
# （WPS 已打开，使用热同步）
vim JSA880.js  # 再次修改
python3 jsa_tools.py sync --modules 1 --auto-close

# 4. 完成后完整同步（带备份）
python3 jsa_tools.py sync
```

#### Git 工作流

```bash
# 1. 同步前提交
git add JSA880.js
git commit -m "更新: 修复 bug"

# 2. 同步到 xlsm
python3 jsa_tools.py sync --no-backup

# 3. 测试通过后更新版本
# 编辑 JSA880.js 中的版本号
git add JSA880.js jsa_tools.py
git commit -m "版本: v3.8.4"
```

### 备份策略

#### 备份文件命名

```
原始文件名_backup_YYYYMMDD_HHMMSS.xlsm
```

示例：
```
3.25..._副本_backup_20260202_221512.xlsm
```

#### 备份管理

```bash
# 定期清理（保留最近 5 个）
python3 jsa_tools.py clean

# 手动清理旧备份
cd .backups
ls -t | tail -n +6 | xargs rm -f

# 备份到外部位置
cp .backups/* ~/Documents/jsa_backups/
```

### 性能优化

#### 快速迭代

```bash
# 只同步主框架（90% 的情况只需更新这个）
python3 jsa_tools.py sync --modules 1 --no-backup
```

#### 批量同步

```bash
# 为多个项目同步
for file in 项目A.xlsm 项目B.xlsm 项目C.xlsm; do
    python3 jsa_tools.py sync --file "$file" --no-backup
done
```

---

## 维护指南

### 版本管理

#### 更新版本号

```python
# 编辑 jsa_tools.py
class Config:
    CURRENT_VERSION = "3.8.4"  # 更新版本号

# 编辑 JSA880.js
/**
 * 版本: 3.8.4 (2026年2月3日)
 * 更新: 新增功能说明
 */
```

#### 版本信息文件

版本信息存储在 `.version.json`：

```json
{
  "version": "3.8.3",
  "updated_at": "2026-02-02T22:30:00",
  "updated_by": "jsa_tools.py"
}
```

### 目录结构

```
js880/
├── jsa_tools.py              # 主工具脚本
├── JSA880.js                 # 主框架
├── cls生成测试数据.js
├── superPivot_WPS_测试.js
├── superPivot_性能测试.js
├── .backups/                 # 备份目录
│   └── *_backup_*.xlsm
├── .temp/                    # 临时目录（自动清理）
├── .version.json             # 版本信息
└── jsa_tools使用指南.md      # 本文档
```

### 配置维护

#### 更新目标文件路径

编辑 `jsa_tools.py` 中的 `Config.XLSM_FILE`：

```python
class Config:
    BASE_DIR = Path(__file__).parent
    XLSM_FILE = Path("/new/path/to/your/file.xlsm")
```

#### 添加环境变量支持（可选）

```python
import os

class Config:
    # 支持环境变量
    XLSM_FILE = Path(os.getenv(
        'JSA_XLSM_FILE',
        '/default/path/to/file.xlsm'
    ))
```

### 日志和监控

#### 启用日志

```bash
# 保存同步日志
python3 jsa_tools.py sync 2>&1 | tee sync_$(date +%Y%m%d).log
```

#### 监控备份大小

```bash
# 检查备份目录大小
du -sh .backups/

# 查看备份文件数量
ls -1 .backups/*.xlsm | wc -l
```

### 安全建议

#### 文件权限

```bash
# 设置适当的权限
chmod 644 *.js              # JS 文件可读
chmod 755 jsa_tools.py      # 工具脚本可执行
chmod 600 .version.json     # 版本文件仅所有者可写
```

#### 备份加密（可选）

```bash
# 加密备份
gpg --encrypt .backups/file_backup.xlsm

# 解密备份
gpg --decrypt file_backup.xlsm.gpg > file.xlsm
```

---

## 附录

### 完整命令参考

```bash
# 同步命令
python3 jsa_tools.py sync                                    # 同步所有模块
python3 jsa_tools.py sync --file path/to/file.xlsm           # 指定文件
python3 jsa_tools.py sync --modules 1,4                      # 指定模块
python3 jsa_tools.py sync --auto-close                       # 热同步
python3 jsa_tools.py sync --force                            # 强制同步
python3 jsa_tools.py sync --no-backup                        # 不备份
python3 jsa_tools.py sync -f file.xlsm -m 1,4 --auto-close   # 组合使用

# 其他命令
python3 jsa_tools.py test                                    # 打开测试
python3 jsa_tools.py test --no-info                          # 打开文件（无说明）
python3 jsa_tools.py version                                 # 查看版本
python3 jsa_tools.py version -v                              # 查看详细版本
python3 jsa_tools.py status                                  # 查看状态
python3 jsa_tools.py status -v                               # 查看详细状态
python3 jsa_tools.py clean                                   # 清理临时文件
python3 jsa_tools.py --help                                  # 查看帮助
```

### 错误代码

| 退出代码 | 含义 |
|---------|------|
| 0 | 成功 |
| 1 | 失败（无效参数、文件错误等） |

### 相关文档

- [JSA880维护指南.md](./JSA880维护指南.md) - JSA880 框架维护指南
- [SuperPivot测试报告_20260202.md](./SuperPivot测试报告_20260202.md) - 测试报告
- [Array2D维护指南.md](./Array2D维护指南.md) - Array2D 类文档

---

**文档维护**: 本文档应随 `jsa_tools.py` 的更新同步维护。

**问题反馈**: 如有问题，请检查故障排除章节或查看相关文档。
