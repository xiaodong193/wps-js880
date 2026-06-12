# 测试公式备忘录

> 数据源：Sheet2 A2:L23（f1=序号 f2=材料名称 f3=规格 f4=单位 f5=数量 f6=单价 f7=金额 f12=经手人）
> 技巧：WPS JSA 不支持 `...` 展开运算符，用 `.concat()` 替代

---

## superPivot corner 测试 ✅ 已完成

| # | 位置 | 公式 |
|---|------|------|
| ① | Sheet2 N3 | `=jsaLambda("$$.superPivot",A2:L23,{"f2"},{"f12"},{"sum(f7)","金额求和"})` |
| ② | Sheet2 N16 | `=jsaLambda("$$.superPivot",A2:L23,{"f2"},{"f12"},{"count(),sum(f7),average(f7)","计数,求和,均值"})` |
| ③ | Sheet2 N24 | `=jsaLambda("$$.superPivot",A2:L23,{"f2"},{""},{"sum(f7),count()","求和,计数"})` |
| ④ | 多层透视 J1 | `=k("$$.superPivot",A1:H40,"f3,f2","f6","count(),sum(f4),textjoin(f4,'+')")` |
| ⑤ | 多层透视 J15 | `=k("$$.superPivot",A1:H40,"f3,f2","f6","sum(f4),average(f5)")` |
| ⑥ | 多层透视 J25 | `=k("$$.superPivot",A1:H40,"f2","f6","sum(f4),average(f5)")` |

---

## 箭头函数测试

数据源：Sheet2 A2:L23

### L1: 筛选保留表头 ✅ 已通
**O32:** `=k("arr => [arr[0]].concat(arr.slice(1).filter(x => x.f7 > 100))", A2:L23)`
→ 金额>100的行（水表107.69×2、卡套150、槽钢4464、螺栓165.36）+ 表头

### L2: 筛选取列
**O44:** `=k("arr => [arr[0]].concat(arr.slice(1).filter(x => x.f5 > 10).map(x => [x.f2, x.f5]))", A2:L23)`
→ 数量>10的 材料名+数量，保留表头

### L3: 总金额
**O54:** `=k("arr => arr.slice(1).reduce((s, x) => s + x.f7, 0)", A2:L23)`
→ 单值：所有金额之和

### L4: 按人分组求和
**O56:** `=k("arr => { var m={}; arr.slice(1).forEach(x => { var p=x.f12; m[p]=(m[p]||0)+x.f7 }); return Object.entries(m) }", A2:L23)`
→ 经手人→金额汇总，2列

### L5: 排序
**O62:** `=k("arr => [arr[0]].concat(arr.slice(1).sort((a,b) => b.f7 - a.f7))", A2:L23)`
→ 按金额降序，保留表头

### L6: 材料去重列表
**O74:** `=k("arr => [...new Set(arr.slice(1).map(x => x.f2))]", A2:L23)`
→ 不重复的材料名称（单列溢出）

---
