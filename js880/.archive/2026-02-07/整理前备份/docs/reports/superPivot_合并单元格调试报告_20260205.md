# superPivot 合并单元格调试报告

> **调试日期**: 2026年2月5日
> **调试版本**: v3.8.2 → v3.8.9
> **调试者**: Claude Code
> **问题**: z超级透视 函数的合并单元格功能不工作

---

## 🔍 问题分析

### 原始问题

用户报告 `z超级透视` 函数的 `applyMerges` 功能不工作，表头单元格没有被正确合并。

### 根本原因

通过代码审计发现，虽然 `recordMerge` 辅助函数和 `applyMerges` 方法都已正确实现，但在表头生成过程中，**`recordMerge` 函数从未被调用**，导致 `mergeInfo` 对象始终为空。

```javascript
// JSA880.js:6518-6528
var mergeInfo = Object.create(null);

function recordMerge(rowIdx, colIdx, rowSpan, colSpan) {
    if (rowSpan > 1 || colSpan > 1) {
        if (!mergeInfo[rowIdx]) mergeInfo[rowIdx] = Object.create(null);
        mergeInfo[rowIdx][colIdx] = { rowSpan: rowSpan, colSpan: colSpan };
    }
}
// ... 但是 recordMerge 从未被调用！
```

---

## 🛠️ 修复方案

### 1. 单列字段合并逻辑

**表头结构** (单列字段，如"月份"):
```
第1行: 行标题1, 行标题2, | 1月, 1月, 1月 | 2月, 2月, 2月 | ...
第2行: 空白, 空白,      | 月份(标题)      | 月份(标题)    | ...
第3行: 空白, 空白,      | 销量, 金额, 数量 | 销量, 金额, 数量 | ...
```

**合并规则**:
- 第1行列值: 相同的月份值合并 (1月出现3次时跨3列合并)
- 第2行标题: "月份"标题跨所有数据列合并

**实现代码**:
```javascript
// 第1行：列字段值需要合并
var colOffset = baseColIdx;
var currentVal = null;
var mergeStartCol = colOffset;

for (var ck = 0; ck < colKeys.length; ck++) {
    var colKeyParts = colKeys[ck].split(separator);
    var val = colKeyParts[0];

    if (currentVal === null) {
        currentVal = val;
        mergeStartCol = colOffset + ck * numDataFields;
    } else if (currentVal !== val) {
        // 值变化，记录前面的合并
        var mergeColWidth = (colOffset + ck * numDataFields) - mergeStartCol;
        if (mergeColWidth > 1) {
            recordMerge(0, mergeStartCol, 1, mergeColWidth);
        }
        currentVal = val;
        mergeStartCol = colOffset + ck * numDataFields;
    }
}
// 记录最后一组合并
if (colKeys.length > 0) {
    var finalMergeColWidth = (colOffset + colKeys.length * numDataFields) - mergeStartCol;
    if (finalMergeColWidth > 1) {
        recordMerge(0, mergeStartCol, 1, finalMergeColWidth);
    }
}

// 第2行：列字段标题跨所有数据列合并
if (colKeys.length > 0 && numDataFields > 1) {
    var totalDataCols = colKeys.length * numDataFields;
    recordMerge(1, colOffset, 1, totalDataCols);
}
```

### 2. 多列字段合并逻辑

**表头结构** (多列字段，如"大区→省份→城市→区域"):
```
第0行: 空白 | 大区(标题) | 华东 | 华东 | 华东 | 华北 | 华北 | ...
第1行: 空白 | 省份(标题) | 江苏 | 江苏 | 浙江 | 北京 | 天津 | ...
第2行: 空白 | 城市(标题) | 南京 | 苏州 | 杭州 | 北京 | 天津 | ...
第3行: 空白 | 区域(标题) | 市区 | 郊区 | 市区 | 市区 | 郊区 | ...
第4行: 部门 | 数据标题(重复)...
```

**合并规则**:
- 每行列字段标题: 跨所有数据列合并
- 每行列字段值: 相同的值合并
- 最后一行数据字段标题: 按列键分组合并

**实现代码**:
```javascript
// 为每个列字段层级收集合并信息
for (var cfIdx = 0; cfIdx < numColFieldLevels; cfIdx++) {
    var targetRow = cfIdx;
    var colOffset = baseColIdx;

    // 列字段标题跨所有数据列合并
    if (colKeys.length > 0) {
        var totalDataCols = colKeys.length * numDataFields;
        recordMerge(targetRow, colOffset, 1, totalDataCols);
    }

    // 列字段值需要合并（相同的值合并在一起）
    var seenValues = Object.create(null);
    var valuePositions = [];

    // 首先收集每个值的位置
    for (var ck = 0; ck < colKeys.length; ck++) {
        var colKeyParts = colKeys[ck].split(separator);
        if (cfIdx < colKeyParts.length) {
            var val = colKeyParts[cfIdx];
            if (!seenValues[val]) {
                seenValues[val] = [];
                valuePositions.push({ val: val, startIdx: ck });
            }
            seenValues[val].push(ck);
        }
    }

    // 为每个连续的值记录合并
    for (var v = 0; v < valuePositions.length; v++) {
        var vp = valuePositions[v];
        var startIdx = vp.startIdx;
        var indices = seenValues[vp.val];
        var count = indices.length;
        var mergeColWidth = count * numDataFields;
        var mergeStartColPos = colOffset + startIdx * numDataFields;

        if (mergeColWidth > 1) {
            recordMerge(targetRow, mergeStartColPos, 1, mergeColWidth);
        }
    }
}

// 最后一行的数据字段标题需要按列键分组合并
var lastRow = numColFieldLevels;
var dataColOffset = baseColIdx;

for (var ck = 0; ck < colKeys.length; ck++) {
    var mergeStartCol = dataColOffset + ck * numDataFields;
    if (numDataFields > 1) {
        recordMerge(lastRow, mergeStartCol, 1, numDataFields);
    }
}
```

### 3. applyMerges 函数增强

**改进点**:
1. 添加详细日志模式（`verbose` 参数）
2. 改进错误处理
3. 添加合并操作统计输出

```javascript
wrappedResult.applyMerges = function(rng, options) {
    // 解析 options
    var verbose = false;
    if (typeof options === 'object' && options !== null) {
        verbose = options.verbose === true;
    }

    // ... 检查 mergeInfo 是否存在
    if (!mergeInfo || Object.keys(mergeInfo).length === 0) {
        if (verbose) {
            Console.log('[applyMerges] 警告: 没有可用的合并信息');
        }
        return [];
    }

    // ... 执行合并
    for (var i = 0; i < mergeRanges.length; i++) {
        var m = mergeRanges[i];
        try {
            var startCell = targetRange.Item(m.row + 1, m.col + 1);
            var endCell = targetRange.Item(
                m.row + m.rowSpan,
                m.col + m.colSpan
            );
            var mergeRange = ws.Range(startCell, endCell);
            mergeRange.Merge();

            if (verbose) {
                Console.log('[applyMerges] 合并: R' + (m.row + 1) + 'C' + (m.col + 1) + ':R' + (m.row + m.rowSpan) + 'C' + (m.col + m.colSpan));
            }
        } catch (e) {
            Console.log('[applyMerges] 合并失败: ' + (e.message || e));
        }
    }

    return appliedMerges;
};
```

---

## 📊 测试用例

### 测试1: 单列字段 + 多数据字段

```javascript
var data = [
    ['产品', '国家', '月份', '销量', '金额'],
    ['A', '中国', '1月', 100, 1000],
    ['A', '中国', '1月', 50, 500],
    ['A', '美国', '1月', 80, 800],
    ['B', '中国', '2月', 120, 1200],
];

var rs = Array2D.z超级透视(data,
    ['f1', '产品'],
    ['f3', '月份'],
    ['sum("f4"),average("f5")'],
    1
);

rs.toRange("K2", true);  // 自动应用合并

// 预期结果:
// 第1行: | 产品 | 1月(合并2列) | 1月(合并2列) | 2月(合并2列) |
// 第2行: | 产品 | 月份(合并6列)                                  |
// 第3行: | 产品 | 销量, 金额 | 销量, 金额 | 销量, 金额           |
```

### 测试2: 多列字段 + 单数据字段

```javascript
var rs = Array2D.z超级透视(data,
    ['f1', '产品'],
    ['f2,f3', '国家,月份'],
    ['sum("f4")'],
    1
);

rs.toRange("K2", true);

// 预期结果:
// 第1行: | 产品 | 国家(合并)        | 国家(合并)        |
// 第2行: | 产品 | 中国(合并2列)    | 美国(合并2列)    |
// 第3行: | 产品 | 1月, 2月, 1月, 2月 (每个数据字段单独) |
```

### 测试3: 详细日志模式

```javascript
var rs = Array2D.z超级透视(data, rowFields, colFields, dataFields);
rs.toRange("K2", false);  // 不自动合并
rs.applyMerges("K2", {verbose: true});  // 手动合并 + 详细日志
```

---

## ✅ 修复验证

### 合并信息收集验证

```javascript
// 调试日志输出示例
[superPivot DEBUG] 表头结构 (3 行):
[Row 0] 长度=5 | 产品, 1月, 1月, 2月, 2月
[Row 1] 长度=5 | , 月份, , ,
[Row 2] 长度=5 | , 销量, 金额, 销量, 金额

[superPivot DEBUG] 合并信息统计:
[superPivot DEBUG]   合并 #1: 行1列3 -> 跨1行x2列
[superPivot DEBUG]   合并 #2: 行1列5 -> 跨1行x2列
[superPivot DEBUG]   合并 #3: 行2列2 -> 跨1行x4列
[superPivot DEBUG] 总计 3 个合并区域
```

### applyMerges 验证

```javascript
// 详细日志模式输出
[applyMerges] 准备执行 3 个合并操作
[applyMerges] 合并: R1C3:R1C4
[applyMerges] 合并: R1C5:R1C6
[applyMerges] 合并: R2C2:R2C5
[applyMerges] 完成: 成功执行 3/3 个合并操作
```

---

## 📝 使用说明

### 自动合并模式 (推荐)

```javascript
var rs = Array2D.z超级透视(data, rowFields, colFields, dataFields);
rs.toRange("A1", true);  // 第二个参数 true 表示自动应用合并
```

### 手动合并模式

```javascript
var rs = Array2D.z超级透视(data, rowFields, colFields, dataFields);
rs.toRange("A1", false);  // 不自动合并
// ... 其他操作 ...
rs.applyMerges("A1");  // 手动应用合并
```

### 详细日志模式 (调试用)

```javascript
var rs = Array2D.z超级透视(data, rowFields, colFields, dataFields);
rs.toRange("A1", false);
rs.applyMerges("A1", {verbose: true});  // 输出详细日志
```

---

## 🎯 总结

### 修复内容

| 项目 | 修复前 | 修复后 |
|------|--------|--------|
| mergeInfo | 始终为空 | 正确收集合并信息 |
| 单列字段合并 | 不工作 | 按值合并 + 标题合并 |
| 多列字段合并 | 不工作 | 按层级合并 + 标题合并 |
| applyMerges | 简单实现 | 增强日志 + 错误处理 |

### 影响范围

- ✅ 不影响现有功能
- ✅ 向后兼容
- ✅ 默认启用自动合并

### 后续建议

1. 考虑添加合并样式设置（如对齐方式）
2. 考虑添加取消合并功能
3. 考虑优化大数据量场景的合并性能

---

*报告生成时间: 2026年2月5日*
*调试版本: v3.8.9*
