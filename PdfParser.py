import importlib
import sys

from pdfminer.cmapdb import main

importlib.reload(sys)
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBoxHorizontal
from pdfminer.pdfinterp import (
    PDFPageInterpreter,
    PDFResourceManager,
    PDFTextExtractionNotAllowed,
)
from pdfminer.pdfparser import PDFDocument, PDFParser

"""[解析pdf文本，保存在txt文件中]
"""
# 每次更改输入pdf的路径即可
pdfpath = r"./1.pdf"

# 这几个不要动
textpath = r"./tmpout.txt"
rmdictpath = r"./rmdict.csv"
outputpath = r"./outdict.csv"


def parse():
    fp = open(pdfpath, "rb")  # 以二进制读模式打开
    # 用文件对象来创建一个pdf文档分析器
    praser = PDFParser(fp)
    # 创建一个PDF文档
    doc = PDFDocument()
    # 连接分析器 与文档对象
    praser.set_document(doc)
    doc.set_parser(praser)

    # 提供初始化密码
    # 如果没有密码 就创建一个空的字符串
    doc.initialize()

    # 检测文档是否提供txt转换，不提供就忽略
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        # 创建PDf 资源管理器 来管理共享资源
        rsrcmgr = PDFResourceManager()
        # 创建一个PDF设备对象
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        # 创建一个PDF解释器对象
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        # 循环遍历列表，每次处理一个page的内容
        for page in doc.get_pages():  # doc.get_pages() 获取page列表
            interpreter.process_page(page)
            # 接受该页面的LTPage对象
            layout = device.get_result()
            # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象
            # 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal
            # 等等 想要获取文本就获得对象的text属性，
            for x in layout:
                if isinstance(x, LTTextBoxHorizontal):
                    with open(textpath, "a", encoding="utf-8") as f:
                        results = x.get_text()
                        f.write(results + "\n")


def solve():
    import jieba
    import pandas as pd

    df = pd.read_csv(rmdictpath, header=None)
    excludes = set(df[0].values.tolist())
    excludes.update(
        r"1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.-_,() -<>αβγµμ&@%×*′+−—–|°:;  /=•±Δ’"
    )
    txt = open(textpath, "r", encoding="utf-8").read()
    words = jieba.lcut(txt)
    counts = {}
    for word in words:
        if len(word) == 0:  # 排除单个字符的分词结果
            continue
        else:
            counts[word] = counts.get(word, 0) + 1
    for word in excludes:
        try:
            del counts[word]
        except KeyError:
            continue
    items = list(counts.items())
    items.sort(key=lambda x: x[1], reverse=True)
    for i in range(30):
        word, count = items[i]
        print("{: <20}{:->10}".format(word, count))
    pd.DataFrame(items, columns=["letter", "frequency"]).to_csv(outputpath, index=None)


if __name__ == "__main__":
    parse()
    solve()
