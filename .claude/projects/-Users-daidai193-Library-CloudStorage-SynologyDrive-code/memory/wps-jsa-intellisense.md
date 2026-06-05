---
name: wps-jsa-intellisense
description: WPS JSA IntelliSense 引擎有文件解析行数限制，namespace 定义必须在文件前部才能被识别
metadata:
  type: feedback
---

WPS JSA 的 IntelliSense 引擎对 JS 文件有解析行数限制（大约前 1000 行以内）。只有定义在文件前部的 namespace/class 才能获得自动补全支持。

**Why:** 实测发现 Array2D（第 502 行）能补全，而 JSA（第 10506 行）和 ShtUtils（原第 9985 行）不能补全。代码模式（对象字面量/静态方法/prototype）都不是根本原因，位置才是。

**How to apply:**
- 新增 namespace/class 时，必须放在 JSA880.js 的前部（Array2D 基础架构之后，约第 599 行附近）
- 需要自动补全的命名空间不能放在文件后半部分
- 正确的模式：`function Xxx() {}` + `@constructor` + `@class` JSDoc + `Xxx.prototype = Object.create(...)` + `Xxx.prototype.constructor = Xxx` + 方法定义在 `Xxx.prototype.xxx`
- 参考 [[jsa880-code-structure]] 了解文件结构布局
