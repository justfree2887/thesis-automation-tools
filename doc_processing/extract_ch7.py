#!/usr/bin/env python3
"""提取第七章机理分析内容"""
import zipfile
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

doc_path = '/Users/shiberlin/Desktop/毕业论文/毕业论文_修改版.docx'
with zipfile.ZipFile(doc_path) as z:
    xml_content = z.read('word/document.xml')

root = etree.fromstring(xml_content)
body = root.find('.//' + W + 'body')

def get_run_text(r):
    texts = []
    for t in r.findall(W + 't'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

def get_para_text(p):
    return ''.join(get_run_text(r) for r in p.findall(W + 'r'))

# 找到第七章内容（从第七章标题开始到结论章之前）
ch7_paras = []
in_ch7 = False
for i, elem in enumerate(body):
    if elem.tag != W + 'p':
        continue
    text = get_para_text(elem).strip()
    if text.startswith('第七章') or text == '7':
        in_ch7 = True
    if in_ch7:
        ch7_paras.append((i, text))

# 输出第七章所有段落
out = []
for bi, text in ch7_paras:
    if text:
        out.append(f'body[{bi}] {text}')

with open('/Users/shiberlin/Desktop/毕业论文/ch7_content.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(out))
print(f'共提取 {len(ch7_paras)} 个段落')
