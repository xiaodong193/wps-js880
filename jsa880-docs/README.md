# JSA880 文档网站

郑广学 JSA880 快速开发框架的 API 文档网站

## 基于 VitePress 构建

### 目录结构

```
jsa880-docs/
├── docs/
│   ├── .vitepress/
│   │   └── config.mjs       # VitePress 配置
│   ├── api/
│   │   ├── index.md         # API 概述
│   │   ├── global-functions.md  # JSA 全局函数
│   │   └── array2d-class.md      # Array2D 类
│   ├── guide/
│   │   ├── getting-started.md  # 快速开始
│   │   └── lambda.md            # Lambda 表达式
│   └── index.md             # 首页
├── scripts/
│   └── extract-api.js      # API 提取脚本
├── package.json
└── .github/workflows/
    └── deploy.yml          # GitHub Actions 部署
```

## 开发

### 安装依赖

```bash
cd jsa880-docs
npm install
```

### 本地预览

```bash
npm run docs:dev
```

访问 http://localhost:5173

### 构建

```bash
npm run docs:build
```

### 提取 API 文档

```bash
npm run extract-api
```

此脚本会从 `../js880/JSA880.js` 提取 API 信息并生成 JSON 文件。

## 部署

本项目配置了 GitHub Actions，会在推送到 `main` 分支时自动：

1. 安装依赖
2. 提取 API 文档
3. 构建 VitePress 站点
4. 部署到 GitHub Pages

需要先在 GitHub 仓库设置中启用 GitHub Pages，选择 `gh-pages` 分支。

## 文档规范

### 函数文档格式

```markdown
### z函数名 / aliasName

描述...

```javascript
// 示例代码
```

**参数:**
| 参数 | 类型 | 说明 |
|------|------|------|
| `name` | String | 参数说明 |

**返回:** `Type` - 返回值说明
```

### 中英文对照

所有函数都需要标注中英文名称：
- 中文名：`z筛选`、`z排序`、`z超级透视`
- 英文别名：`filter`、`sort`、`superPivot`

## 许可证

MIT License