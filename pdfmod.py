import os
import fitz
import base64
import magic
import requests
import json
from PIL import Image
from PyPDF2 import PdfMerger
from logsys import logger

def mergeImages(pdfp, files, wid=240, hei=140):
    valid_files = []
    for fi in files:
        try:
            with open(fi, 'rb') as f:
                contents = f.read()
            img_format = magic.from_buffer(contents, mime=True).split('/')[-1]
            if img_format == 'heic':
                ctr = base64.b64encode(contents).decode("utf-8")
                open("./cache/text.txt", "w").write(ctr)
                resp = requests.post("http://tun4.cquluna.top:20062/convert", headers={"Content-Type": "application/json"},
                                     data=json.dumps({"data": ctr}))
                open(fi, 'wb').write(base64.b64decode(resp.text))
            tmp = Image.open(fi)
            if (tmp.height > tmp.width) != (hei > wid):
                tmp = tmp.transpose(Image.ROTATE_90)
            tmp.save(fi)
            valid_files.append(fi)
        except Exception as ex:
            logger.warning(f"画像処理失敗、スキップ: {fi} - エラー: {ex}")
            continue

    pdf = fitz.open()
    for fi in valid_files:
        try:
            img = fitz.open(fi)
            # rect = img[0].rect
            pdfbytes = img.convert_to_pdf()
            imgPdf = fitz.open("pdf", pdfbytes)
            page = pdf.new_page(width=wid*5, height=hei * 5)
            page.show_pdf_page(fitz.Rect(0.0, 0.0, wid*5, hei*5), imgPdf, 0)
        except Exception as ex:
            logger.warning(f"PDF変換失敗、スキップ: {fi} - エラー: {ex}")
            continue
    pdf.save(pdfp)


def mergePdfs(pdfp, files):
    file_merger = PdfMerger()
    for fi in files:
        file_merger.append(fi)
    file_merger.write(os.path.abspath(pdfp))
    logger.debug(os.path.abspath(pdfp))


def pdfexport(pdfp, prefix):
    pdf = fitz.open(pdfp)
    res = list()
    for pg in range(0, pdf.page_count):
        page = pdf[pg]
        pm = page.get_pixmap(dpi=300)
        pm.save(f"{prefix}_{pg}.png")
        res.append(f"{prefix}_{pg}.png")
    pdf.close()
    return res
