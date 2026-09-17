#!/usr/bin/env python3
"""
三级回退提取 docx 内 MathType / OMath 公式里的参数候选数字。

为什么需要：本机 GDI/Word 渲染失败时，MathType 公式变成 oleObject*.bin 或
media/*.wmf，python-docx 读不到内容。本脚本走：
  1) word/document.xml 的 m:t（Math 文本节点）+ w:t（普通文本 run）
  2) word/media/*.wmf 的字节扫描（ASCII + UTF-16LE 数字串）
  3) 列出 word/embeddings/oleObject*.bin 的存在（仅提示，不解析二进制）

用法:
  python extract_mathtype_params.py <源文件.docx> [--out report.md]
"""
import sys
import re
import zipfile
import argparse


def extract(docx_path):
    report = []
    with zipfile.ZipFile(docx_path) as z:
        names = z.namelist()

        # 第 1 级：document.xml 文本节点
        if 'word/document.xml' in names:
            xml = z.read('word/document.xml').decode('utf-8', 'ignore')
            mt = re.findall(r'<m:t[^>]*>([^<]*)</m:t>', xml)
            runs = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml)
            report.append(('document.xml math-text (m:t)',
                           [x for x in mt if x.strip()][:300]))
            report.append(('document.xml run-text (w:t 含数字)',
                           [x for x in runs if re.search(r'[0-9]', x)][:300]))
        else:
            report.append(('document.xml', ['__MISSING__']))

        # 第 2 级：WMF 字节扫描
        wmf_entries = []
        for n in names:
            if n.startswith('word/media/') and n.endswith('.wmf'):
                data = z.read(n)
                asc = [x.decode('latin1', 'ignore')
                       for x in re.findall(rb'[-\d.]{3,}', data)]
                try:
                    u16 = data.decode('utf-16-le', 'ignore')
                    uni = re.findall(r'[−\-]?\d+\.?\d*', u16)
                except Exception:
                    uni = []
                wmf_entries.append((n, asc[:60], uni[:60]))
        report.append(('wmf_byte_scan', wmf_entries))

        # 第 3 级：OLE 对象仅列存在
        ole = [n for n in names if 'oleObject' in n]
        report.append(('ole_objects_present (二进制, 未解析)', ole))

    return report


def render(report):
    lines = []
    for kind, val in report:
        lines.append(f'\n## {kind}\n')
        if isinstance(val, list) and val and isinstance(val[0], tuple):
            for name, asc, uni in val:
                lines.append(f'### {name}')
                lines.append('  ASCII: ' + ' '.join(asc))
                lines.append('  UTF16: ' + ' '.join(uni))
        else:
            if not val:
                lines.append('  (空)')
            for v in val:
                lines.append('- ' + str(v))
    return '\n'.join(lines)


def main():
    ap = argparse.ArgumentParser(description='docx MathType 参数三级回退提取')
    ap.add_argument('docx', help='待提取的 .docx 路径')
    ap.add_argument('--out', help='报告输出路径 (markdown)')
    a = ap.parse_args()

    rep = extract(a.docx)
    txt = f'# MathType 参数提取回退报告: {a.docx}\n' + render(rep)

    if a.out:
        with open(a.out, 'w', encoding='utf-8') as f:
            f.write(txt)
        print('报告已写入:', a.out)
    else:
        print(txt)


if __name__ == '__main__':
    main()
