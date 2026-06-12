# inject_jsa880.py 使用说明

> 一键把 JSA880 框架 + KO k 函数 UDF shim 注入到任何 `.xlsm` 工作簿,让打开的 xlsm 立即支持 `=k(...)` 公式

---

## 🎯 解决什么问题

WPS 把 JSA 自定义函数(UDF)存放在 xlsm 内的 `xl/JDEData.bin` 文件里(WPS 专有 XML 格式)。手动做这件事的步骤是:
1. 打开 WPS → JSA 编辑器
2. 找到 ThisWorkbook 模块
3. 粘入 `KO一切的k函数_UDF模块.js` 内容
4. Ctrl+S 保存
5. 重启 xlsm

如果 xlsm **完全没有任何 JSA 代码**(纯数据型工作簿),第 1-2 步还要先:
- WPS → 选项 → 加载项 → 加载 `JSA880.js`
- 然后让工作簿引用这个加载项

`inject_jsa880.py` 把这两步都自动化:把 JSA880 框架 + KO k 函数 shim 一起注入到 `xl/JDEData.bin`,**双击打开 xlsm 就有 `=k(...)` 公式可用**。

---

## 🚀 一行命令

```bash
# 最常用:把 v4.8.51 的 JSA 代码 + KO k 函数 shim 合并到 v4.8.5
python3 inject_jsa880.py <目标.xlsm> \
    --source <有JSA的源.xlsm> \
    --shim KO一切的k函数_UDF模块.js
```

**示例**(用户实际场景):
```bash
cd "/Users/daidai193/Library/CloudStorage/SynologyDrive-code"

python3 "JS880教案/第03章/3-28/inject_jsa880.py" \
    "wps-cesuan/收益测算表开发V4.8.5-wps版本.xlsm" \
    --source "wps-cesuan/收益测算表开发V4.8.51-wps版本.xlsm" \
    --shim "JS880教案/第03章/3-28/KO一切的k函数_UDF模块.js"
```

输出:
```
📦 目标 xlsm: /.../收益测算表开发V4.8.5-wps版本.xlsm
📦 Shim JS:   /.../KO一切的k函数_UDF模块.js
📦 源 xlsm:   /.../收益测算表开发V4.8.51-wps版本.xlsm
🔧 模式:      copy+inject

💾 备份: /.../收益测算表开发V4.8.5-wps版本.xlsm.bak

  [copy] 从 收益测算表开发V4.8.51-wps版本.xlsm 提取 JDEData.bin: 1081658 bytes
  [check] test.xlsm: 现有 16 个文件,没有 JDEData.bin
  [add] /.../test.xlsm :: xl/JDEData.bin 添加完成

📋 inject shim 到 test.xlsm:JDEData.bin
  [JDEData.bin] 已有 13 个 codemodule,最大 id=29
  [追加] codemodule KO_k函数 (id=30)
  [写回] /var/folders/.../tmp.../JDEData.bin (931302 bytes)
  [add] /.../test.xlsm :: xl/JDEData.bin 添加完成

✅ 全部完成!
   接下来: 用 WPS 打开 收益测算表开发V4.8.5-wps版本.xlsm
   Console 会看到 "✅ k() UDF 已就绪!"
```

---

## 📦 工作原理

### WPS JSA 代码存储结构

xlsm 里的 `xl/JDEData.bin` 是 WPS 专有的 XML 格式,结构:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<document version="2.0">
  <name>Project</name>
  <property desc="" lock="false" password="" />
  <activemodule>8</activemodule>
  <codemodule name="m调息" id="21">
    <window cursorpos="0" actived="true" visible="false" />
    <codetext>/**&#x0A; * ...JavaScript 源码(HTML entity 编码)...</codetext>
  </codemodule>
  <codemodule name="m调息" id="21">...</codemodule>
  ...
  <functionsdata />
</document>
```

每个 `<codemodule>` 就是一个 JSA 模块,`id` 必须唯一,`name` 唯一即可,`<codetext>` 里是源码(经 HTML entity 编码)。

### 工具工作步骤

```
copy+inject 模式(默认)
━━━━━━━━━━━━━━━━━━━━
1. 解压 源.xlsm  → 提取 xl/JDEData.bin
2. 解压 目标.xlsm → 修改 xl/JDEData.bin
3. inject_shim: 在 JDEData.bin 的 </functionsdata> 前追加 codemodule
4. 把修改后的 JDEData.bin 重新打包到 目标.xlsm

copy 模式
━━━━━━━━━
1. 提取源.xlsm 的 xl/JDEData.bin
2. 直接覆盖目标.xlsm 的 xl/JDEData.bin(若有)

inject 模式
━━━━━━━━━━
1. 解压目标.xlsm 的 xl/JDEData.bin
2. 追加 shim codemodule
3. 重新打包
```

---

## 🛡️ 安全保证

1. **自动备份** — 任何覆盖前自动生成 `.xlsm.bak`(可在 `main()` 里加 `--no-backup` 关闭)
2. **临时文件** — 改 JDEData.bin 时先写 `.tmp`,完成后再 `rename`,防止写一半挂掉损坏原文件
3. **XML 编码** — 写入 codetext 前把所有特殊字符做 HTML entity 编码(`&` `<` `>` `"` `'` `\n` 等),保证 WPS 能正确读回
4. **id 唯一** — 自动找最大 id + 1,不跟现有模块冲突
5. **同名替换** — 已有同名 codemodule 则替换而非追加,多次运行安全

---

## 🎯 常见场景

### 场景 1: 全新 xlsm,完全没 JSA 代码

```bash
python3 inject_jsa880.py 我的工作簿.xlsm \
    --source v4.8.51-wps版本.xlsm \
    --shim KO一切的k函数_UDF模块.js
```

完成后打开"我的工作簿.xlsm",JSA Console 立刻显示:
```
✅ k() UDF 已就绪!(JSA880 v4.2.2)
   测试:在任意单元格输入
   =k("JSA.getIndexs", 1, 10, 2)
   看到 1 3 5 7 9 = 成功!
```

### 场景 2: xlsm 已有自己的 JSA,但缺 KO k 函数

```bash
python3 inject_jsa880.py 我自己的项目.xlsm \
    --shim KO一切的k函数_UDF模块.js
```

工具只追加 `KO_k函数` codemodule,不影响现有模块。

### 场景 3: 只想检查 JDEData.bin 长啥样

```bash
unzip -p 我的工作簿.xlsm xl/JDEData.bin | head -50
```

或解压后用文本编辑器打开看。

---

## 🔧 命令行参数

| 参数 | 说明 | 默认值 |
|---|---|---|
| `target` | 目标 xlsm 路径(必填) | - |
| `--source` | 源 xlsm 路径(只有 copy 类模式才需要) | - |
| `--mode` | `copy` / `inject` / `copy+inject` | `copy+inject` |
| `--shim` | KO k 函数 shim JS 文件路径 | 同目录 `KO一切的k函数_UDF模块.js` |
| `--module-name` | 注入到 JDEData.bin 里的 codemodule 名字 | `KO_k函数` |
| `--no-backup` | 不生成 `.bak` 备份(默认会备份) | `False` |

---

## ❓ 常见问题

### Q1: 注入后打开 xlsm 报错"无法识别的项目格式"

可能原因:
- 源 xlsm 的 JDEData.bin 是用比当前 WPS 版本更新的格式写的
- 用更新版本的 WPS 重新打开即可

### Q2: 注入后 =k(...) 还是报 `#NAME?`

注入只对**当前 WPS 公式引擎版本**有效。如果用户的 WPS 不支持 WPS 15990+ 数组溢出,或 JSA 编辑器不识别 `KO_k函数` 名字,可能需要:
- 打开 JSA 编辑器 → 找到 "KO_k函数" 模块 → 看 Console 有没有报错
- 模块名用了非 ASCII 字符(`KO_k函数` 含中文),WPS 15990 之前可能不支持

**应急方案**:把 `--module-name KO_k函数` 改成 `--module-name KO_k_module`(纯 ASCII):
```bash
python3 inject_jsa880.py ... --module-name KO_k_module
```

### Q3: 想撤销注入

直接用备份还原:
```bash
mv 收益测算表开发V4.8.5-wps版本.xlsm.bak 收益测算表开发V4.8.5-wps版本.xlsm
```

### Q4: 想再注入一个 JSA 模块

工具目前只支持注入 KO k 函数 shim。如果想注入别的模块,可以:
```bash
# 把另一个 JS 改个名当 shim
cp 我的模块.js /tmp/我的模块.js
python3 inject_jsa880.py 目标.xlsm --shim /tmp/我的模块.js --module-name 我的模块 --mode inject
```

### Q5: 工具运行报错 "目标 xlsm 里没有 xl/JDEData.bin,无法 inject"

在 `--mode copy` 模式下 xlsm 本来就可以没有 JDEData.bin。但 `--mode inject` 必须有。`--mode copy+inject` 会先 copy 源 JDEData.bin 到目标,再 inject——所以目标是"空"也没关系。

---

## 📋 实现细节

| 关键技术点 | 处理方式 |
|---|---|
| XML 解析 | Python 标准库 `xml.etree.ElementTree` |
| HTML entity 编码 | 自定义 `encode_for_xml` 函数,跟 WPS 格式一致(`'` 包裹 `'`、`&#x0A;` 包裹 `\n` 等) |
| Zip 重打包 | 临时 zip 文件 → 逐项复制 → 写入新文件 → rename,避免损坏 |
| 跨平台路径 | `pathlib.Path`,支持 macOS / Linux / Windows |
| 编码一致性 | 整个流程 UTF-8 |

---

## 🔗 相关文件

- `inject_jsa880.py` — 工具脚本
- `KO一切的k函数_UDF模块.js` — 注入的 shim(顶层 function k / jsaLambda / Workbook_Open 等)
- `KO一切的k函数.md` — k 函数完整文档(含 6 种 fn 语法、容错、FAQ)
- `js880/JSA880.js` — 框架源代码(从源 xlsm 的 JDEData.bin 注入)
