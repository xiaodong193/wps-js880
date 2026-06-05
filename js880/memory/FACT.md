# JSA880框架与Agent知识库

## 框架版本
- **JSA880框架**: v4.0.0 (完整代码位于 `/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/jsa880-framework.js`)
- **JSA880 Agent技能**: 已创建并注册到CherryStudio

## JSA880框架组成

### 核心模块
| 模块 | 说明 |
|------|------|
| Array2D | 二维数组处理类，支持链式调用 |
| JSA | 全局工具函数库 |
| $ | Range快捷访问对象 |
| agg | 聚合函数 |
| IO | 文件操作 |
| DateUtils | 日期处理 |
| RngUtils | 单元格工具 |
| SuperMap | 调试字典 |
| PicUtils | 图片操作 |
| ShtUtils | 工作表工具 |
| FormUtils | 窗体工具 |

### 安全约束
- 禁止使用 ES2020+ 语法：`?.`、`??`、`??=`、`BigInt`
- 必须使用 `var` 声明变量
- 比较使用 `===` 而非 `=`
- 日期操作使用 `cdate()` 或 `DateUtils.fromExcelDate()`

## 文件路径
- 完整框架代码: `/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/jsa880-framework.js`
- Agent技能目录: `/Users/daidai193/Library/Application Support/CherryStudio/Data/Skills/jsa880-agent/`
- 知识库文档: `/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/jsa880-docs/`

## 使用方式
将 `jsa880-framework.js` 的内容完整复制到WPS宏编辑器中即可使用全部功能。

## 核心语法
- f模式: `f1`, `f2` 代表第1、2列
- Lambda字符串: `'f2 > 100'`
- 链式调用: `arr.z筛选('f1>0').z排序('f2+').toRange('A1')`