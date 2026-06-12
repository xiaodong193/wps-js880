# soul.md — JSA880 开发经验

> 从实际调试中积累的经验，每次踩坑后更新。

---

## xlsm 注入：JDEData.bin 编码规则

xlsm 的 `xl/JDEData.bin` 是 XML 格式，JSA 代码存在 `<codemodule><codetext>` 里。

### 正确注入方式

```python
from xml.etree import ElementTree as ET

root = ET.fromstring(bin_text)
for cm in root.findall('codemodule'):
    if cm.get('name') == 'JSA880':
        cm.find('codetext').text = raw_js_code  # 直接赋值原始 JS，不要预编码
        break
new_bin = ET.tostring(root, encoding='unicode')
```

### 禁止的预编码操作

| 错误做法 | 后果 | 原因 |
|---------|------|------|
| `'` → `&apos;` | JS 语法错误 | ET 二次编码 `&`→`&amp;`，WPS 解码后 JS 看到 `&apos;` |
| `\n` → `&#x0A;` | JS 收到字面量 | 同上 |
| `"` → `&quot;` | 同理 | 同上 |

**原则：`ct.text = raw_code`，让 ET 自动处理 `&` `<` `>` 的 XML 转义，换行和引号不做任何额外处理。**

### 模块顺序必须保持

`ThisWorkbook` 调 `JSA.k`，所以 `JSA880` 必须在 `ThisWorkbook` 之前。
正确顺序：`Module2 → JSA880 → ThisWorkbook`

---

## superPivot corner 空列修复（v4.0.42）

### 问题

单列字段表头，corner 硬编码为空串：

```javascript
if (!hideRowTitles) {
    headerRows[0].push('');  // BUG: 前导空列
}
headerRows[0].push(getFieldTitle(colConfig.fields[0], 0, 'col'));
```

### 修复（最终版）

```javascript
if (!hideRowTitles) {
    headerRows[0].push(
        numRowFieldLevels === 1
            ? (cornerTitle || getFieldTitle(colConfig.fields[0], 0, 'col'))
            : ''
    );
    if (numRowFieldLevels > 1) recordMerge(0, 0, numRowFieldLevels, 1);
    for (var rfIdx = 0; rfIdx < numRowFieldLevels; rfIdx++)
        headerRows[1].push(getFieldTitle(rowConfig.fields[rfIdx], rfIdx, 'row'));
}
if (numRowFieldLevels > 1)
    headerRows[0].push(getFieldTitle(colConfig.fields[0], 0, 'col'));
```

### 关键：区分单行字段 vs 多行字段

| 行字段数 | 示例 | corner | 列字段标题 |
|---------|------|--------|----------|
| 1 | sheet2 `f2` | 填"经手人" | corner `[0][0]` |
| ≥2 | 多层透视 `f3,f2` | `''` merged | `[0][numRowFieldLevels]` |

第一次修复只覆盖单行字段，多行字段（多层透视 J1）corner 被错误填入"年"导致列错位。

### 不影响范围

多层列字段、无列字段、hideRowTitles 模式 — 均为独立分支，未修改。

---

## 版本兼容性

- 嵌入版本 vs 加载项版本可能不同，要确认实际加载路径
- 注入新版本时从同结构基座 xlsm 出发，只替换 codetext
- Node 测试通过 ≠ WPS 能跑，必须在 WPS 验证

## superPivot corner 不合并修复（v4.0.37 XXD-74）

### 问题

单列字段 + 多行字段时，corner `recordMerge(0, 0, numRowFieldLevels, 1)` 合并了角标区 (1,1)(2,1)，不应合并。

### 修复

删除 `recordMerge`，行 0 改为遍历 `rfIdx` 推 `numRowFieldLevels` 个空 cell：

```javascript
if (numRowFieldLevels === 1) {
    headerRows[0].push(cornerTitle || getFieldTitle(colConfig.fields[0], 0, 'col'));
} else {
    for (var rfIdx = 0; rfIdx < numRowFieldLevels; rfIdx++) {
        headerRows[0].push('');
    }
}
// 不再调用 recordMerge(0, 0, numRowFieldLevels, 1)
```

### 影响

表头列数从 14 → 15（2 row fields 时），行 0/行 1 列对齐，corner 不再合并。

不被影响的路径：单行字段（`numRowFieldLevels===1` 走独立分支）、多列字段（`numColFieldLevels>1` 走独立分支）、`hideRowTitles`。

## superPivot corner 不合并 — XXD-74 验证（v4.0.37）

### 问题

Bug C (XXD-69) 引入的 `recordMerge(0, rfIdx, headerRowCount, 1)` 在**多列字段分支**合并了角标 (1,1)(2,1)，用户期望 4 个 corner cell 独立。

### 修复

多列字段分支（`numColFieldLevels > 1`）中删除：
```javascript
// 旧: recordMerge(0, rfIdx, headerRowCount, 1);
// 新: 仅保留注释，不合并
```

单列字段分支的 `recordMerge(0, 0, numRowFieldLevels, 1)` **保留**（Bug A，不同场景）。

### 验证

- `rg recordMerge` 确认多列字段角标合并已删除
- 单列字段 `recordMerge` 完好保留
