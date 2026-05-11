# Claude Code 自动测试 WPS JSA 完整指南

> 版本: 1.0
> 创建日期: 2026年2月1日
> 作者: Claude Code

---

## 快速开始

### 一键测试流程

```bash
# 在 Claude Code 中运行
bash js880/jsa_claude_test.sh
```

这个脚本会：
1. 检查是否有测试结果 → 如果有，直接显示
2. 如果没有 → 同步代码并打开 WPS
3. 提示你在 WPS 中运行测试
4. 再次运行脚本读取结果

---

## 工作原理

```
┌─────────────┐     同步      ┌─────────────┐
│ Claude Code │ ───────────> │   xlsm 文件  │
└─────────────┘             └─────────────┘
                                │
                                │ 用户运行测试
                                ▼
┌─────────────┐   读取结果   ┌─────────────┐
│ Claude Code │ <────────── │ /tmp/jsa_*.txt│
└─────────────┘             └─────────────┘
```

### 文件通信机制

| 文件 | 用途 |
|------|------|
| `/tmp/jsa_test_result.txt` | 测试结果 |
| `/tmp/jsa_test_error.txt` | 错误信息 |

---

## 在 WPS 中可用的测试函数

| 函数 | 说明 | 输出 |
|------|------|------|
| `ClaudeAutoTest()` | 快速测试 | 文件 |
| `自动测试_完整版()` | 完整测试 | 文件 |
| `RunQuickTest()` | 快速测试 | 工作表 |
| `RunAllTests()` | 完整测试 | 工作表 |

---

## 使用场景

### 场景1: 修改代码后测试

```bash
# 1. Claude 修改 JSA880.js
# 2. 运行一键测试
bash js880/jsa_claude_test.sh

# 3. 在 WPS 中运行: ClaudeAutoTest()
# 4. 再次运行脚本读取结果
bash js880/jsa_claude_test.sh
```

### 场景2: Claude 主动测试

```javascript
// Claude 可以通过 Bash 工具执行
bash js880/jsa_sync.py  // 同步代码
bash js880/jsa_claude_test.sh  // 打开 WPS

// 然后提示用户运行测试
// 用户在 WPS 中运行 ClaudeAutoTest()

// Claude 再次执行读取
bash js880/jsa_claude_test.sh  // 读取结果
```

### 场景3: 读取现有结果

```bash
# 如果已经有测试结果，直接读取
cat /tmp/jsa_test_result.txt
```

---

## 测试结果格式

### 成功示例

```
========================================
JSA 自动化测试结果
========================================

测试时间: 2026/2/1 13:19:10
耗时: 0.15 秒
状态: 成功

========================================
详细日志
========================================

[时间] [INFO] ========== 测试开始 ==========
[时间] [INFO] ---------- 自动化测试 ----------
[时间] [PASS] ✅ 通过 - 求和测试: 21 = 21
[时间] [INFO] 表格: 透视结果 (3 行 x 3 列)
[时间] [INFO] ========== 测试完成 ==========
```

### 失败示例

```
========================================
测试失败
========================================

错误位置: xxx
错误消息: xxx is not defined
```

---

## Claude Code 工作流示例

### 示例1: 完整自动化流程

```javascript
// Claude 执行

// 1. 同步代码
Bash: python3 js880/jsa_sync.py

// 2. 打开 WPS
Bash: open -a wpsoffice /path/to/test.xlsm

// 3. 提示用户
"请在 WPS 宏编辑器中运行: ClaudeAutoTest()"

// 4. 等待用户确认后读取结果
Bash: cat /tmp/jsa_test_result.txt

// 5. 分析结果
// 根据测试结果反馈用户
```

### 示例2: 快速迭代

```bash
# 用户: 修改 z求和 方法
# Claude: 同步并测试

bash js880/jsa_sync.py
bash js880/jsa_claude_test.sh

# 输出: WPS 已打开，请运行 ClaudeAutoTest()

# [用户在 WPS 中运行测试]

# Claude: 读取结果
bash js880/jsa_claude_test.sh

# 输出: 测试结果...
# ✅ 通过 - 求和测试: 21 = 21
```

---

## 高级用法

### 自定义测试

在 `jsa_auto_test.js` 中添加自定义测试：

```javascript
function Claude自定义测试() {
    var test = TestHelper.capture();

    // 你的测试代码
    test.section('我的测试');
    test.assert(条件, '描述');

    test.finish();

    // 导出到文件
    写入测试结果(test.logBuffer.join('\n'));
}
```

### 添加更多测试函数

```javascript
function Claude测试SuperPivot() {
    var result = Array2D.z超级透视(...);
    写入测试结果(JSON.stringify(result.val()));
}
```

---

## 故障排除

### Q: 测试结果文件不存在？

A: 请确保在 WPS 中运行了测试函数：
```
ClaudeAutoTest()  // 或 自动测试_完整版()
```

### Q: 文件写入失败？

A: 检查 `/tmp` 目录权限，或确保 WPS 有文件系统访问权限。

### Q: 如何清空测试结果？

A: 运行此命令：
```bash
> /tmp/jsa_test_result.txt
```

---

## 文件清单

| 文件 | 说明 |
|------|------|
| `jsa_sync.py` | 多模块同步脚本 |
| `jsa_auto_test.js` | 自动化测试模块 |
| `jsa_claude_test.sh` | 一键测试脚本 |
| `jsa_auto_read.py` | 读取测试结果 |

---

## 总结

**现在 Claude Code 可以：**

1. ✅ 自动同步代码到 xlsm
2. ✅ 自动打开 WPS
3. ✅ 通过文件读取测试结果
4. ✅ 分析错误并修复

**工作流程：**

```
修改代码 → 同步 → 运行测试 → 读取结果 → 分析修复
   ↑                                        ↓
   └──────────────── Claude Code ◄─────────┘
```

---

*文档由 Claude Code 自动生成*
