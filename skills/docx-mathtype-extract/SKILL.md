---
name: docx-mathtype-extract
description: 从 .docx 提取 MathType/OLE 公式里的真实数字参数（二次项系数、均值方差、常量等）。当本机 GDI/Word 渲染失败、公式变成 OLE 对象或 WMF 图片时，用"docx 文本层 + WMF 字节扫描 + document.xml 交叉校验"三级回退术拿到可读数字。当用户说"提取公式参数""MathType公式取不出""docx公式提取""WMF/oleObject""Table 1参数""公式渲染失败""论文里的公式读不出来"时触发。
---

# docx 内 MathType 公式参数提取（GDI 渲染失败回退术）

## 何时使用
- 源文件（题目/规范/论文）.docx 里的关键参数以 MathType 公式形式给出，直接读 docx 文本读不到。
- 本机 GDI/Word 渲染环境缺失，公式变成 OLE 对象（oleObject*.bin）或 WMF 图片（media/*.wmf），`python-docx` 拿不到公式内容。
- 需要从"图片里的公式"反推出可录入代码的魔法数。

## 为什么普通方法会失败
MathType 在 docx 里有两种存储：
1. **OLE 对象**：`word/embeddings/oleObject*.bin`，是二进制复合文档，普通文本提取读不出。
2. **WMF 图片**：`word/media/*.wmf`，公式被矢量化成图片。
本机若缺 GDI32 渲染（sandbox 常见），连 WMF 都画不出来，只能当字节流处理。

## 三级回退提取术（按顺序）

### 第 1 级：docx 文本层（最快，常有意料外收获）
很多 MathType 公式在 docx 的段落 XML 里**留有纯文本等价物**（尤其简单公式、单位、数字）。
```python
from docx import Document
doc = Document('源文件.docx')
for p in doc.paragraphs:
    if any(k in p.text for k in ['系数', '参数', '表']):   # ← 换成你要找的参数名
        print(repr(p.text))
```
若文本层能直接读到形如 `a=…, b=…, c=…` 的系数值，直接用，跳过下级。

### 第 2 级：WMF 字节扫描
WMF 内部仍可能含 ASCII/UTF-16 的数字与变量名（公式的文字部分）。解压 docx 后扫所有 wmf：
```python
import zipfile, re
z = zipfile.ZipFile('源文件.docx')
for n in z.namelist():
    if n.startswith('word/media/') and n.endswith('.wmf'):
        data = z.read(n)
        nums = re.findall(rb'[-\d.]{3,}', data)   # ASCII 数字串
        print(n, [x.decode('latin1','ignore') for x in nums][:30])
```
把扫到的数字与第 1 级交叉比对，定位 U 形二次 `a*v*v - b*v + c` 的 a/b/c。

### 第 3 级：document.xml 交叉校验
`word/document.xml` 里公式区可能是 `m:oMath`（Office Math，非 OLE）或文本。
```bash
unzip -o 源文件.docx word/document.xml -d /tmp/docx
grep -o '<m:t>[^<]*</m:t>' /tmp/docx/word/document.xml | head -200
```
`m:t` 是 Math 文本节点，常含公式里的字母变量与数字。

## 配套脚本
`scripts/extract_mathtype_params.py`：自动跑完三级回退，输出每个 media wmf 的数字候选 + document.xml 的 m:t 节点，供人工定参。用法：
```bash
python extract_mathtype_params.py <源文件.docx> [--out report.md]
```

## 提取后必做：物理量级复核
提取到的数字**不要直接入主线**，先用物理常识复核数量级再决定采不采。典型失效：某个系数的常数项被多读/少读一位，
代回原式后同类两个指标相差一个数量级——这种"物理上不可能"的落差才是发现提取错误的信号，而不是模型的错。

判据：**同一模型里两个本应同量级的指标若相差 ≥10 倍，先怀疑提取，别先怀疑建模。**

## 禁止
- 不要假设"能渲染出图片就能直接 OCR"——本机 GDI 失败时图片也是空的，必须走字节扫描。
- 不要跳过物理复核：提取出的数字若导致同类指标数量级异常，优先怀疑提取错，不是模型错。
- 不要用提取到的数直接覆盖主线已验证值，除非三级回退结果一致且通过物理复核。
