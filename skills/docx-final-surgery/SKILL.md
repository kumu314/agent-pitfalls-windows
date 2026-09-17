---
name: docx-final-surgery
description: 对已定稿的 Word 论文/报告做定点外科手术——给 OMML 公式批量加编号（公式居中、编号 (x.y) 贴最右端）、给多子图图题去掉 (a)(b)(c)(d) 罗列只留总题、给表格加行/删行/换项（含补建超链接关系）、把附录文件清单与压缩包对齐、**输出 Ctrl+H 替换清单前做命中面盘点（短数字会误伤数据值）与连锁改动台账**，并保证"除指定位置外一字不改"。附"真源集合校验 + 逐段 diff + PDF 真渲染全文 diff"三重验证。触发词：公式编号、公式后面加序号、图题精简、图名带abcd不好看、Word定点修改、只改这两处、其他不要改、docx保守编辑、保留格式改文字、表格加一行、附录清单对不上、源码表补几行、Ctrl+H、批量替换、全字匹配、替换会误伤吗、清单执行了没有。
---

# Word 终稿定点手术（公式编号 / 图题精简 / 保真改写）

## 何时用
- 用户给一份**已定稿**的 .docx，只要求改极少数位置，并且明确说「其他一字不要改」。
- 需要给展示公式批量加编号，且要求「公式居中 + 编号 (5.1) 位于最右端」。
- 图题里带 `(a)…(b)…(c)…(d)…` 罗列，用户觉得不美观，要求只留总题（图内的 (a)(b) 小标题保留不动）。

## 铁律
1. **先侦察再动手**：写 `inspect_doc.py` 输出「段落号 | 样式 | 是否含 oMath/oMathPara | 文本」，再逐段导出 `pPr` / `runs` / `drawing` 计数。**不要凭段落文本来定位**——公式段的 `p.text` 是空串，只能靠 `el.findall(qn('m:oMathPara'))` 找。
2. **改前存快照**：`orig_texts = [p.text for p in paras]`、`orig_styles = [...]`，改完逐段比对。
3. **只动该动的**：不要"顺手"改纸张、页边距、字体、图内文字。
4. **别人手工改过的文档，必须逐字符回读，不能信"看起来改了"**：用户拿着你的清单在 Word 里手改过一轮后再交回来，很可能出现 ① 改了一半 ② 用错了字符。**真实案例**：清单让他把 `python prob1/solve.py` 补成带 `--air`，回读发现是 `python prob1/solve.py—air`——连的是**破折号 U+2014**，不是 ` --`（空格+两个连字符），照抄命令直接报错退出。**对策**：动手前先写 `probe_keys.py`（关键词落点）+ `probe_runs.py`（逐 run 打印），并用 `re.compile(r'[\u2014\u2013][A-Za-z]')` 全文扫「破折号紧跟字母」的可疑组合（注意排除正常用法：中文破折号、`Crank–Nicolson`、`40–80 °C` 数值范围）。
5. **清单会过期，文档才是真相**：清单写"删 X 行、替换 Y 项"时，实际文档可能已经删过了、或表里根本没有 Y。**不要机械执行清单**，先实测当前状态，再以「外部真源」（如支撑材料 zip、原始数据文件）反推目标，见 §C。

## 技术路径

### A. 公式编号（OMML）
`m:oMathPara` 是**独占显示行**，无法与制表位同排。做法：**拆包成行内 `m:oMath`**，前后各插一个 tab run，pPr 加制表位。

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

pPr = el.find(qn('w:pPr'))
# 1) 制表位：中心 + 右对齐（pos 是绝对 twips；A4 版心 16cm -> center 4536 / right 9072）
tabs = OxmlElement('w:tabs')
for val, pos in (('center', 4536), ('right', 9072)):
    t = OxmlElement('w:tab'); t.set(qn('w:val'), val); t.set(qn('w:pos'), str(pos)); tabs.append(t)
pStyle = pPr.find(qn('w:pStyle'))
pStyle.addnext(tabs)                     # schema 顺序：pStyle -> tabs -> spacing -> ind -> jc
# 2) 缩进清零（否则制表位不以版心左边界起算）
ind = OxmlElement('w:ind'); ind.set(qn('w:firstLine'), '0'); ind.set(qn('w:firstLineChars'), '0'); ind.set(qn('w:left'), '0')
pPr.find(qn('w:spacing')).addnext(ind)
# 3) 段落改左对齐 —— 居中交给中心制表位（保留 w:jc=center 会与 tab 打架）
jc = pPr.find(qn('w:jc'))
if jc is None: jc = OxmlElement('w:jc'); pPr.append(jc)
jc.set(qn('w:val'), 'left')
# 4) 拆 oMathPara -> 行内 oMath
omp = el.find(qn('m:oMathPara')); om = omp.find(qn('m:oMath'))
pos = list(el).index(omp); el.remove(omp)
r1 = OxmlElement('w:r'); r1.append(OxmlElement('w:tab'))
r2 = OxmlElement('w:r')
rPr = OxmlElement('w:rPr')
f = OxmlElement('w:rFonts')
for k in ('w:ascii', 'w:hAnsi', 'w:cs'): f.set(qn(k), 'Times New Roman')   # 西文/数字用 TNR
f.set(qn('w:hint'), 'eastAsia'); rPr.append(f); r2.append(rPr)
r2.append(OxmlElement('w:tab'))
wt = OxmlElement('w:t'); wt.text = '(5.1)'; r2.append(wt)
for i, e in enumerate((r1, om, r2)): el.insert(pos + i, e)
```

要点
- **pPr 子元素顺序必须遵 schema**：`pStyle → tabs → spacing → ind → jc`。用 python-docx 的 `pPr.get_or_add_tabs() / get_or_add_ind()` 会自动保序，更省事。
- **制表位 `w:pos` 是绝对 twips**：A4 版心 16.0cm → 4536/9072；US Letter 版心 16.59cm → 4703/9406。客户还没从 Letter 改到 A4 时，**按 A4 设**更好——Letter 上只是编号略内缩 0.6cm，改 A4 后精确贴右。
- 行内化**不会**把 `m:limLoc=m:undOvr` 降级，∑/∫ 的上下限照常显示（已实测）。
- 章号判定：由前文最近的 `Heading 1/2` 推。附录里的公式可编 `(附.1)`。

### B. 图题精简（保格式改文字）
不要 `p.clear()` 或重建段落——会丢「Image Caption」样式/加粗/居中。做法：**保留首个 run 的 rPr，只换文字、删其余 run**。

```python
import re
for p in caption_paras:
    runs = p.runs
    first = runs[0]; first.text = new_text          # python-docx 会自动处理 xml:space
    for r in runs[1:]: r._r.getparent().remove(r._r)

# 「图序 + 分隔符」原样保留（全角/半角空格的差异别自己改）
m = re.match(r'^(图\d+-\d+)(\s+)', old); new = m.group(1) + m.group(2) + title
```
- **警惕"一个图题被拆成多段"**：子图说明有时是 4 个独立段落，必须**整段删除**，否则留下孤立的 `(a)…`。
- 删完记得校正段落总数。

### C. 表格手术（加行 / 删行 / 换项 / 补超链接）
**铁律：绝不手工拼 `<w:tr>` XML，一律 `copy.deepcopy` 现有行再改文本**——边框、字体、行高、段落属性 100% 继承。

```python
tr = copy.deepcopy(sample_tr)                      # 样板行
tcs = tr.findall(qn('w:tc'))
set_tc(tcs[0], fname); set_tc(tcs[1], lines)       # 见下方 set_tc
tr_last.addprevious(tr)                            # ★ 插在「末行」之前，末行的粗底线自动保持
```

要点
- **选样板行**：中间行（如 `rows[1]`）带**上细线** `w:top sz=2`；末行带**下粗线** `w:bottom sz=12`。想加行就 clone 中间行、**插在末行之前**，边框逻辑自洽、无需改任何边框。
- **`set_tc(tc, text)`：把单元格压成「单段落 单 run」**——删多余 `<w:p>`，删全部 `<w:r>` 与 `<w:hyperlink>`，再 append 一个新 run。rPr 的选择顺序：① 段落里带 `w:rFonts[@w:hint="eastAsia"]` 的 run（中文才不会走西文字体）② 任意 run ③ 同列样板行的 rPr 兜底（**空单元格没有 run，必须走兜底**，否则新 run 无字体）。
- **超链接列**：用 `doc.part.relate_to(url, RT.HYPERLINK, is_external=True)` 自动分配 rId，再构造
  ```python
  h = OxmlElement('w:hyperlink'); h.set(qn('r:id'), rId)
  r = OxmlElement('w:r'); rpr = OxmlElement('w:rPr')
  rs = OxmlElement('w:rStyle'); rs.set(qn('w:val'), 'af'); rpr.append(rs)   # 'af'=超链接字符样式
  r.append(rpr); t = OxmlElement('w:t'); t.text = url; r.append(t); h.append(r)
  ```
  保存后核验 `word/_rels/document.xml.rels` 里 hyperlink 数 = 行数。
- **`w:lastRenderedPageBreak`** 是 Word 的渲染缓存，clone 出来的新行里若带它**要删掉**（`el.findall('.//'+qn('w:lastRenderedPageBreak'))`）。
- **清空/填补「原来就空着的单元格」**：先 `assert cell.text.strip() == ''` 确认，再 `set_tc` 填。若原表有空格位，**填满比留空好**（留空看起来像漏填）；但要注意左右两列的语义搭配（同类放同一行）。

### D. 清单类表格与外部真源对齐（最强校验）
当表格是「支撑材料文件清单」「源码表」这种**有外部对应物**的清单时：
1. **以外部真源为唯一基准**（如 `支撑材料.zip` 的实际文件列表），把它归一成与论文相同的相对路径口径（`coder/x.py`→`x.py`，`figs/x.png`→`../figs/x.png`，包说明 README 不计入）。
2. 改完后做 **`set` 相等断言**，而不是数个数：
   ```python
   assert set(清单项) == set(zip内容), (set(清单项)-set(zip内容), set(zip内容)-set(清单项))
   assert set(源码表第0列) == set(zip里的.py)
   ```
   差集能直接告诉你「包里有的、清单没有」和「清单有的、包里没有」。
3. **再用「行数合计」交叉验证**：源码表「行数」列 `int()` 求和应等于正文里写的总行数（本例 5652+2143=**7795**，与论文所写分毫不差，是强证据）。
4. 三项**全过才允许 `doc.save`**，任一不过 `sys.exit(1)` 不落盘。

### E. 页型与痕迹硬化（Letter → A4、清属性、清 descr）

改 `sectPr` 与 `docProps` 时**直接改 zip parts 比走 python-docx 更稳**（不会重排其余部件、体积不变），改完立即回读断言。

```python
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
sect = root.find(W+"body").find(W+"sectPr")        # ⚠️ 全文档可能不止一个，先数 len(list(root.iter(W+"sectPr")))
sect.find(W+"pgSz").set(W+"w","11906"); sect.find(W+"pgSz").set(W+"h","16838")   # A4
pgmar = sect.find(W+"pgMar")
for k in ("top","right","bottom","left"): pgmar.set(W+k,"1418")   # 2.5009cm ≥ 2.5cm
```

**⭐ 整数 twips 的数学必然（A4 + 2.5cm 边距无解）**：A4 宽 11906 twips，2.5cm = 1417.32 twips。
- 取 **1417** → 边距 2.4991cm，**不达 2.5cm 硬规**；
- 取 **1418** → 边距 2.5009cm 合规，但版心 = 11906−2836 = **9070 twips**，小于按 16.00cm 设的公式编号右制表位 **9072**。

**必须选 1418**（硬规优先）。代价是编号最多超出新右界 2 twips。**实测影响 0.218pt = 0.077mm，且 Word 会把越界制表位自动 clamp 到版心内**——渲染与打印均不可见，不要为此再去改 14 处制表位。

**痕迹清理清单**（官方"不得出现身份信息"类条款）：
| 位置 | 常见残留 |
|---|---|
| `docProps/core.xml` | `dc:creator`（本机用户名）、`cp:lastModifiedBy`（**真实姓名**）、`cp:author` |
| `docProps/app.xml` | `Company`、`Manager` |
| `word/document.xml` | 图片 `descr` 属性里的内部路径（如 `../../01_OUTBOX/figs/x.png`）|

## 验证（缺一不可）
0. **外部真源集合校验**：清单类表格必做，见 §D。这是最强的一道——它能抓出「看着对但少了/多了文件」。
1. **机检**：读回后确认 ① 编号个数/顺序/无重号；② 图题不再含 `(a)`；③ 遍历所有段落，除计划改动外 `text` 与 `style` 与快照**逐字一致**，打印不一致条数（应为 0）。
2. **真渲染**：LibreOffice 转 PDF → **全文拼接后 `difflib`**：
   ```bash
   "C:\Program Files\LibreOffice\program\soffice.exe" --headless --norestore --convert-to pdf --outdir <outdir> <doc.docx>
   ```
   忽略 `Could not find platform independent libraries <prefix>` 警告。
   ```python
   # 去页码行后拼接全部页文本，再压缩空白，然后 SequenceMatcher 比差异块
   ops = [o for o in difflib.SequenceMatcher(None, A, B, autojunk=False).get_opcodes() if o[0] != 'equal']
   ```
   **判据**：差异块数应**恰好等于**「新增编号数 + 删除的图题文本数」（表格行因分页重排产生的 delete+insert 成对块算净零）。出现任何未解释的块 = 内容被误伤，必须回滚查因。
   - ⚠️ **别做「逐页对页」**：只要改动让**总页数变了**（哪怕只多 1 页），后面所有页都会整体错位，逐页比对会刷出几十处假差异（表现为"某整段被删除"）。**永远用「全文拼接后一次 difflib」**，并同时打印页数与页锚点（摘要/正文起/参考文献/附录各在第几页）来判断改动是否越界。
   - **附录加行是"安全"的地方**：多数赛制「附录页数不限」，而正文有页数上限。若改动集中在附录，正文页界不应变动——打印页锚点即可确认。
3. 目视：把含公式/图题的页渲染成 PNG（`pymupdf`：`pg.get_pixmap(dpi=125)`）逐张看图，确认公式居中、编号贴右、图题干净、图内子图字母仍在。

## 顺手要做的痕迹审计（用户没问也要报）
```python
# 页型、身份痕迹、图片 descr 内部路径、文档属性
sp = doc.element.body.find(qn('w:sectPr')); sp.find(qn('w:pgSz')).attrib   # A4=11906x16838, Letter=12240x15840
zipfile.ZipFile(f).read('docProps/core.xml')   # dc:creator / cp:lastModifiedBy 常留本机人名
re.findall(r'descr="([^"]*)"', raw_xml)        # ⚠️ 只扫 descr 属性，别全文 count
```
- **中文论文页型必须是 A4**；`w:pgSz` 是 Letter 是 pandoc/Word 常见坑。
- **⭐ 校验时别用全文 `count("内部目录名")`**：正文/附录里的 **GitHub 超链接 URL 会合法包含** `01_OUTBOX/...` 这类仓库路径（本例 26 条），全文 count 会误报 FAIL。**必须只统计 `descr` 属性**里的路径。
- GitHub 链接、内部目录名出现在正文/附录 = 匿名风险（很多赛制明令禁止）。
- **只报告，不擅自改**——用户说了"其他不要改"就照办，把发现列成"待拍板清单"。**唯一例外：违反官方硬性项且属纯形式项（页型、页边距、身份痕迹），这类是"必做"而非"可选"**，可直接执行并报告。

### 排版边界测量（判定"越界"是缺陷还是正常）

改页型/边距后，用 `pymupdf` 量正文实际落界，**但下结论前必须做版本对照**：

```python
right_edge = page_width - margin          # A4+2.5cm → 524.38pt
# 逐行 dump l["bbox"]，看非首行缩进行的 x1 是否 ≤ right_edge
```

- **正常**：非缩进行 x1 精确贴合 `right_edge`（本例 524.40 vs 524.38，误差 0.02pt）。
- **⭐ 行尾标点溢出 ~12pt 是「中文标点悬挂」，不是缺陷**：Word 默认允许标点溢出边界（GB/T 15834 推荐做法）。**判据 = 拿改动前的版本做同口径测量**——若两版溢出量几乎相同（本例 Letter 12.1pt vs A4 12.12pt），则**是既有行为、非本次引入**；且只要仍落在页边距区域内就不会被裁切。
- **⭐ 公式编号定位**：LibreOffice 会把 OMML 公式（含右侧制表位编号）渲染为**矢量图形**，普通 span 遍历取不到。要用 `page.search_for("(3.1)")`，**格式是 `(章节.序号)` 而非 `(1)`**，写错正则会得到"未找到"的假结论。编号位置由**绝对制表位**决定，与纸张无关 → 换页型后编号应几乎不动。

## 输出纪律
- 写回原路径前先 `cp` 到工作区留底；`shutil.copy` 报 PermissionError 说明文件被 Word 预览/编辑占用，退而另存 `_编号版.docx` 并告知。
- 写回后校验 zip 部件数、图片二进制 md5 是否与原文件一致（python-docx 重打包会把体积压小，属正常）。

### F. 结构性重排（搬移段落 / 图 / 表）与表格列删除

比"改文字"危险得多，四条必守：

1. **⭐ 搬移前先把下标引用换成元素引用**。`tbls = [c for c in body if c.tag == qn('w:tbl')]` 这种**按下标取表**（如 `tbls[10]` 是 7.1 清单、`tbls[12]` 是源码表）在**任何节点搬移后立即失效**——把附录里的一张表搬到正文，后面所有表的序号整体位移，取值会静默拿到另一张表（表现为"断言 15 != 76"这类莫名其妙的失败）。**做法**：在 `find_p(...)` / 任何 `addnext` 之前，先 `T10 = _tbls0[10]; T12 = _tbls0[12]` 锁死元素引用，再动手搬。
2. **⭐ 文本替换锚点不能跨 run**。Word 会把一句话切成多个 run（换字体/上下标/西文数字都会切）。真实案例：`见附录六图4-2` 实际是 `run0='见附录六图'` + `run1='4-2'`，用整串替换必然 `AssertionError`。**做法**：锚点取**不跨 run 的最短串**（`见附录六图`→`见图`，`见附录六表`→`见表`），替换后断言 `new` 存在于该段。先写一个 dump 脚本打印目标段的**逐 run 文本**再决定锚点。
3. **⭐ 删表格列 = 三件事，少一件就漏**：
   - `w:tblGrid` 里删对应的 `w:gridCol`；
   - 每个 `w:tr` 里删对应的 `w:tc`；
   - **删 `word/_rels/document.xml.rels` 里的超链接关系条目**（`Target` 含目标域名/仓库名的 `Rel`）。**只删单元格会留下一份 XML 里仍能 `grep` 到 URL 的 rels**，匿名化等于没做。
   删完还要**重算剩余列宽保总宽**（`gridCol` 的 `w:w` 与各 `tcW` 一起改；例：4 列等宽 `2351/2351/2352/2352`=9406 → 3 列 `3135/3135/3136`），否则表突然只占 75% 宽。
4. **⭐ 改单元格对齐要按 `w:pPr` schema 序插**。`w:jc` 不能 `append`，须按 `pStyle→…→ind→jc→…→rPr` 的顺序插到第一个"序在 jc 之后"的子元素之前。写一个 `PPR_ORDER` 列表 + `RANK` 字典即可（本技能 §A 已用同一手法处理 `w:tabs`/`w:ind`）。

**另外两件事**：
- **全空表格要专门检测**：`if cells and not any(txt(tc).strip() for tc in cells)` —— 文档里可能藏着**带边框的全空表格**（本例附录末尾一张 7×4，`tblStyle=af4`），在 PDF 里渲染成一个空网格，很扎眼。删之前先渲染末页 PNG 目视确认。
- **每张图有 2 个 `descr`**：`wp:docPr` 上的常是**中文图注**（无障碍替代文本，**应保留**），`pic:cNvPr` 上的常是**图片路径**（`../../01_OUTBOX/figs/x.png`，**必须清**）。所以"清 descr"后**非空数量不为 0 是正常的**，判据是「非空值里有没有 `/` 或 `.png`」。
- **⭐ 二进制文件的"内部串残留"多为假命中**：把 PNG 按 `decode('utf-8', errors='ignore')` 再正则扫，无效字节被丢弃会让两段合法 ASCII **粘连**出关键词（本例 `image8.png` 里扫出 `PR#`，原始字节里根本没有）。**判定必须回到原始字节**（`re.finditer(rb'PR\s*#', raw)`）。

### 交付前的一致性红线（支撑材料类项目）
- **附录文件清单 ↔ 压缩包文件集合必须严格相等**（官方常写「不相符可能取消评奖资格」）。做法：清单里的相对路径按约定映射成 zip 内路径（`common/x.py`→`code/common/x.py`、`../figs/y.png`→`figs/y.png`），然后 **双向求差集，两个方向都必须为空**，不能只数个数。
- 压缩包内**目录名本身也是匿名风险点**：`01_OUTBOX/`、`00_CONTRACT/` 这类流水线目录名要改（改成 `code/`、`schema/`），但要**保证论文清单里的相对路径写法不变**（把 `code/` 设为新的"当前目录"即可）。
- 包内脚本的**注释/docstring 清洗后要做 AST 归一自证**：把所有字符串常量值抹平后 `ast.dump` 对比，证明"只有文案变了、逻辑没变"；**预期会 FAIL 的那 1~2 个文件**正是你刻意做的功能性修正，逐行 `diff` 单独核验。
- 官方第五条 =「全部完整、可运行的源程序」是**硬性**，不能为了"干净"删程序；能删的只是官方未要求的**说明文档**（内部报告/README），且删了要同步删清单条目。

### G. 批量替换（Ctrl+H）的安全边界与连锁改动台账 ★ 2026-09-13 补

改 Word 的最后一公里往往是"给用户一份替换清单，让他自己 Ctrl+H"。**这一步最容易出错，且出错无声无息**。

#### G1. 给清单前必须先做「命中面盘点」

把目标串在**全文**（含表格、页眉页脚、文本框、附录）扫一遍，**逐个命中判定三分类**：

| 分类 | 含义 | 处置 |
|---|---|---|
| 🎯 目标 | 就是你要改的地方 | 计数 |
| ⚠️ 数据 | 是别处的**数值/代码/引用**，恰好包含该串 | 必须排除 |
| 🎲 巧合 | 更长的数字/单词里嵌着它 | 必须排除 |

**真实案例**：清单让用户把附录源码表里 `latent_check.py` 的 **`149` 行** 改成 `133`。
盘点发现 **`149` 在正文出现 2 处**：① 附录 7.2 表 → 🎯；② **表4-5（问题三水分浓度）第 8 行第 3 列 = `0.1494`** → ⚠️。
**直接 Ctrl+H 全替换会把数据值改成 `0.1334`**，且不报错、不显眼。
→ 结论：**必须勾「全字匹配」**，或干脆脚本定点改那一格。

同一批里 `7741→7725` 全文**仅 1 处**（引言里的总行数）→ 安全，可裸替换。
**"同一份清单里两个替换，一个安全一个致命"——所以必须逐条盘点，不能整份打包说"都可以直接替换"。**

#### G2. 短数字/短串必勾「全字匹配」

- **动作前自检**：串长 ≤4 且为纯数字 → **默认假定它会命中数据**，先扫再给。
- 长数字（如 `7741`）**也要扫**，不能凭"它够长"跳过——论文里可能同时出现 `77410`、`0.7741`。
- 短串同理：`--air` vs 破折号 `—air`（见铁律 4）；`xlsxwriter` 里含 `writer`。

#### G3. 格式类批量替换：先证明"命中面 = 目标面"

小数位、科学计数法这类**格式模式**替换（如全部 6 位小数 → 3 位、`0.0e+00` → `0.0000`），做法：

```python
import re
# 1) 先扫描该格式模式的全部命中
pat = re.compile(r'\d+\.\d{6}|\d\.\de\+00')
for i, p in enumerate(paras):
    for m in pat.finditer(p.text):
        print(i, m.group(), p.text[:60])
# 2) 判据：命中数量 == 你预期的格数，且全部落在目标表格区间内
```

**实测结果**：全文扫出 24 个格式格，**全部在表4-7 内，其余零命中** → 这才敢说"Ctrl+H 全文档替换安全"。
**若扫出目标表以外的命中（公式、数据、图注里的小数）→ 不能给全文替换指令**，改成"定位到第 X 表"。

#### G4. 连锁改动台账：改 A 必改 B

脚本侧改动会产生**跨文档连锁**，必须列成表逐条核（不能凭记忆）：

| 改动源 | 受影响位置 | 当前值 | 目标值 |
|---|---|---|---|
| `latent_check.py` 删 16 行 | 附录 7.2 表「行数」列 | 149 | 133 |
| 同上（总行数变化） | 引言里"全部源码共计 N 行" | 7741 | 7725 |
| 同上（脚本描述） | 附录 7.1 表说明列 | 两道自检 | 一道自检 |

**规律：同一个数在论文里往往出现两次**（明细表 + 汇总句）——只改一处必留矛盾。
**动作前自检**：每改一个数，先 grep 它在全文出现的次数，再决定改哪几处。

#### G5. ⭐ 备件生成 ≠ 已落盘（最贵的认知）

用户手工改 Word 时，AI 的正确边界是：**不代改，但有义务回读并报告"执行/未执行"**。

- **禁止**以"清单已提供 / 备件已生成"作为任务终结。
- **必做**：给出清单后，回读目标文件（比 MD5 / 比逐段文本 / 查清单里的串是否还在），明确报告「✅ 已执行」或「⬜ 尚未执行」。
- 真实案例：给完清单后未回读，用户以为改了、我以为交出去了；事后比 MD5 发现**交付版与桌面终稿完全相同** → 两处替换**都没执行**。
- **判据句式**：「桌面终稿 MD5 `XX` == 交付版 MD5 `XX` → 两处替换尚未执行，清单仍待办」。
- **顺带产出**：备件（已改好版 `终稿_fixed.docx`）要放在用户能找到的路径并**明确告知路径**，否则等于没做。

