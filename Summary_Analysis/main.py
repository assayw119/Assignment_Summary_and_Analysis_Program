from textrankr import TextRank
import docx
import olefile
import os
import pandas as pd
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
from konlpy.tag import Okt
from konlpy import jvm
# from eunjeon import Mecab

def read_docx(path):
    """" read docx file in path"""
    document = docx.Document(path)
    docText = '\n\n'.join(
        paragraph.text for paragraph in document.paragraphs
    )
    return docText

def read_hwp(path):
    """ read hwp file in path"""
    f = olefile.OleFileIO(path)
    encoded_text = f.openstream('PrvText').read()
    decoded_text = encoded_text.decode('utf-16')    # unicode
    text = "\n".join(decoded_text.split("\n"))
    return text

def read_pdf(path):    
    """ read pdf file in path"""
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()    
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)
        text = retstr.getvalue()
    fp.close()
    device.close()
    retstr.close()
    text = " ".join(text.split())
    return text


class Tokenizer:
    def __init__(self):
        # self.tokenizer = Mecab()
        self.tokenizer = Okt()

    def __call__(self, text):
        # tokens = self.tokenizer.morphs(text)
        tokens = self.tokenizer.phrases(text)
        return tokens

if __name__ == "__main__":
    jvm.init_jvm()
    df = pd.DataFrame(columns=["학번", "이름", "요약"])

    tokenizer = Tokenizer()
    textrank = TextRank(tokenizer)

    # get file name from dirPATH
    dirPath = "./sample"
    filePaths = os.listdir(dirPath)

    idx = 0
    for filePath in filePaths:
        try:
            pp = filePath.split("/")[-1].split(".")[0]
            id, name = pp.split("_")
            path = os.path.join(dirPath, filePath)
            # docx
            if path.endswith(".docx"):
                text = read_docx(path)
            # hwp
            elif path.endswith(".hwp"):
                text = read_hwp(path)
            # pdf
            elif path.endswith(".pdf"):
                text = read_pdf(path)
            else:
                raise Exception("")
            # summarize text
            summarized = textrank.summarize(text, 3)
            df.loc[idx] = [id, name, summarized]
        except OSError as e:
            print(f"cannot read {filePath} cause {e}.")
            df.loc[idx] = [id, name, None]
        except Exception as e:
            print(f"cannot read {filePath} cause {e}.")
            df.loc[idx] = [id, name, None]
        else:
            idx += 1

    df.to_csv("./summary.csv", index=False, encoding='utf-8-sig')
    print(df.head())