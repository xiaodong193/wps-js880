# JSA880 - WPS Office JavaScript API 快速开发框架

[![版本](https://img.shields.io/badge/版本-v3.9.1-blue.svg)]()
[![WPS](https://img.shields.io/badge/WPS-JSA-green.svg)]()

> 🚀 让 WPS Office 宏开发更简单、更高效！

---

## 📋 项目简介

JSA880 是一个专为 WPS Office JavaScript API (JSA) 设计的快速开发框架，提供了类似 pandas DataFrame 的二维数组处理能力，以及丰富的办公自动化工具函数。

### 核心特性

- 📊 **superPivot** - Excel 风格的多层透视表
- 🔄 **Array2D** - 高性能二维数组处理
- 📁 **IO** - 文件操作工具
- 📅 **DateUtils** - 日期时间工具
- 🔧 **RngUtils/ShtUtils** - Range/工作表快捷操作

---

## 🚀 快速开始

### 方法一：直接导入 (推荐)

1. 下载 `JSA880.js` 或 `dist/JSA880_v3.9.1.js`
2. 在 WPS 表格中按 `Alt + F11` 打开宏编辑器
3. 选择「工具」→「导入文件」→ 选择 `JSA880.js`

### 方法二：复制代码

1. 打开 [JSA880.js](JSA880.js)
2. 复制全部代码
3. 粘贴到 WPS 宏编辑器

### 第一个示例

```javascript
// 准备数据
var data = [
    ['产品', '年份', '销售额'],
    ['手机', '2023', 1000],
    ['电脑', '2024', 2000]
];

// 创建透视表
var result = Array2D.z超级透视(
    data,
    ['f1', '产品'],           // 行字段
    ['f2', '年份'],           // 列字段
    ['sum("f3")', '销售额']   // 数据字段
);

// 输出到工作表
result.toRange("A1");
```

更多示例见 [src/examples/](src/examples/)

---

## 📚 文档导航

| 文档 | 说明 |
|------|------|
| [快速开始](docs/guides/01_快速开始.md) | 5分钟上手教程 |
| [superPivot指南](docs/guides/03_superPivot使用指南.md) | 透视表功能详解 |
| [快速参考卡](docs/guides/快速参考卡.md) | 常用代码速查 |
| [视频教程](docs/guides/视频教程脚本.md) | 跟着做教程 |
| [API参考](docs/api/) | 完整API文档 |
| [示例代码](src/examples/) | 可运行示例 |

📖 **完整文档索引**: [docs/index.md](docs/index.md)

---

## 💡 功能示例

### superPivot 多层透视表

```javascript
var result = Array2D.z超级透视(
    data,
    ['f1,f2', '类别,产品'],      // 多级行字段
    ['f3,f4', '年份,季度'],      // 多级列字段
    ['sum("f5"),count()', '销售额,订单数'],
    1, 1, '@^@',
    {
        cornerTitle: '销售分析报表',
        rowSubtotals: { enabled: true, label: '小计' },
        colSubtotals: { enabled: true, label: '小计' },
        grandTotals: { row: true, column: true },
        displayAs: { mode: 'percentOfGrandTotal', decimals: 2 }
    }
);

result.toRange("A1", true);  // 自动合并标题
```

### Array2D 数据处理

```javascript
// 选择列
var cols = arr.z选择Column(['f1', 'f2']);

// 筛选
var filtered = arr.z筛选(row => row[0] > 100);

// 排序
var sorted = arr.z排序(1, true);  // 按第2列升序

// 统计
var sum = arr.z求和('f3');
var avg = arr.z平均值('f3');
```

---

## 📦 版本历史

### v3.9.1 (2026-02-07)
- 🐛 修复多层表头前导空白列计算
- 🐛 修复单列字段表头行数
- 📝 优化文档结构

### v3.9.0 (2026-02-06)
- ✨ 新增 superPivot 高级功能
  - 行/列小计与总计
  - 百分比显示
  - 多种布局模式
  - 层级缩进
  - 角标题

[查看更多版本历史](CHANGELOG.md)

---

## 📁 项目结构

```
js880/
├── JSA880.js              # 核心框架
├── README.md              # 本文件
├── docs/                  # 文档中心
│   ├── guides/            # 使用指南
│   ├── api/               # API参考
│   └── examples/          # 示例代码
├── src/                   # 源码目录
│   ├── modules/           # 功能模块
│   ├── tests/             # 测试套件
│   └── examples/          # 示例代码
└── dist/                  # 分发版本
```

---

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request！

### 开发规范
- 遵循 ES5 语法规范（使用 `var`）
- 方法调用必须加括号 `()`
- 添加 JSDoc 注释
- 保持命名一致性

详见 [docs/development/贡献指南.md](docs/development/贡献指南.md)

---

## 📞 联系方式

- **原作者**: 郑广学 (EXCEL880)
- **维护者**: 徐晓冬
- **官方网站**: https://vbayyds.com

---

## 📄 许可证

本项目基于 MIT 许可证开源。

---

**免责声明**: 使用本框架产生的任何数据丢失或文件损坏，开发者不承担责任。建议在使用前备份重要数据。

---

**当前版本**: v3.9.1  
**最后更新**: 2026-02-07
