# JSA880 分发版本说明

## 📦 文件清单

| 文件 | 大小 | 说明 | 适用场景 |
|------|------|------|----------|
| JSA880_v3.9.1.js | 424KB | 完整版 | 需要兼容多种环境 |
| JSA880_WPS.js | 423KB | WPS专用版 | 仅在WPS中使用 |

---

## 🚀 推荐使用

### WPS用户 → 使用 JSA880_WPS.js
```
✅ 已移除Node.js兼容代码
✅ 已移除浏览器检测代码  
✅ 体积略小，加载更快
✅ 代码更简洁，易于调试
```

### 需要Node.js支持 → 使用 JSA880_v3.9.1.js
```
✅ 包含完整环境兼容
✅ 支持Node.js文件操作
✅ 支持浏览器环境检测
```

---

## 📥 使用方法

### 在WPS中使用

#### 方法1: 导入文件
1. 打开WPS表格
2. 按 `Alt + F11` 打开宏编辑器
3. 选择「工具」→「导入文件」
4. 选择 `JSA880_WPS.js`

#### 方法2: 复制粘贴
1. 用文本编辑器打开 `JSA880_WPS.js`
2. 复制全部代码
3. 粘贴到WPS宏编辑器中

### 快速测试
```javascript
// 测试Array2D是否可用
Console.log(typeof Array2D);  // 应输出: function

// 测试superPivot
var data = [['A',1],['B',2]];
var result = Array2D.z超级透视(data, 'f1', 'f1', 'count()');
Console.log(result.length);  // 应输出: 3
```

---

## 🔍 版本差异

### JSA880_WPS.js 移除的内容

| 内容 | 移除原因 |
|------|---------|
| `module.exports` | WPS不使用CommonJS |
| `require('fs')` | WPS使用ActiveXObject |
| `typeof window` | WPS不是浏览器环境 |
| `isNodeJS` 判断 | WPS专用，无需判断 |
| `isBrowser` 判断 | WPS专用，无需判断 |

### 保留的核心功能

- ✅ Array2D 全部功能
- ✅ superPivot 全部功能
- ✅ IO 模块 (WPS方式)
- ✅ DateUtils 全部功能
- ✅ RngUtils/ShtUtils 全部功能
- ✅ $ 快捷工具

---

## ⚠️ 注意事项

### WPS专用版限制
1. **不能使用 Node.js 文件操作**
   - ❌ `IO.z读文件()` 在Node路径下不可用
   - ✅ WPS环境下使用 ActiveXObject

2. **不能使用浏览器对象**
   - ❌ `window`, `document` 等不存在
   - ✅ 使用 `Application` 对象

### 版本选择建议

| 场景 | 推荐版本 |
|------|---------|
| 只在WPS中使用 | JSA880_WPS.js |
| 需要在Node.js测试 | JSA880_v3.9.1.js |
| 代码需要在多环境运行 | JSA880_v3.9.1.js |

---

## 📋 VERSION.txt 说明

```
v3.9.1 - 2026-02-07
- 版本号: v3.9.1
- 发布日期: 2026-02-07
- 更新内容: 新增superPivot v3.9.0功能，优化文档结构
```

---

## 📞 技术支持

- **项目主页**: https://vbayyds.com
- **问题反馈**: 请提交Issue
- **使用文档**: 查看 ../docs/guides/

---

**最后更新**: 2026-02-07
