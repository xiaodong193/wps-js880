---
name: jsa880-code-structure
description: JSA880.js 文件结构布局，包含各 namespace 的位置要求
metadata:
  type: project
---

JSA880.js 文件结构（约 17000 行），需要自动补全的 namespace 必须在前部：

**前部（IntelliSense 可识别，约第 1-1000 行）：**
- 第 502 行：`Array2D` - 二维数组操作（@constructor + @class + prototype 模式）
- 第 599 行：`ShtUtils` - 工作表操作工具（同上模式）
- 第 864 行：`ShtUtils_ctor` - ShtUtils 构造器别名

**后部（IntelliSense 不可识别）：**
- 第 6868 行：`_STATIC_METHOD_CONFIG` - Array2D 静态方法批量创建
- 第 8656 行：`RngUtils` - Range 操作工具
- 第 10221 行：`DateUtils` - 日期工具
- 第 10503 行：`JSA` - 全局工具函数
- 第 11238 行：`IO` - 文件操作

**Why:** WPS JSA IntelliSense 有解析行数限制，见 [[wps-jsa-intellisense]]

**How to apply:** 新增需要自动补全的 namespace 时，插入到第 599-864 行区域（ShtUtils 之后、Array2D prototype 方法之前）
