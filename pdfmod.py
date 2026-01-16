import os
import fitz
import base64
import whatimage
import requests
import json
from PIL import Image
from PyPDF2 import PdfMerger
from logsys import logger

def mergeImages(pdfp, files, wid=240, hei=140):
    for fi in files:
        with open(fi, 'rb') as f:
            contents = f.read()
        img_format = whatimage.identify_image(contents)
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

    pdf = fitz.open()
    for fi in files:
        img = fitz.open(fi)
        # rect = img[0].rect
        pdfbytes = img.convert_to_pdf()
        imgPdf = fitz.open("pdf", pdfbytes)
        page = pdf.new_page(width=wid*5, height=hei * 5)
        page.show_pdf_page(fitz.Rect(0.0, 0.0, wid*5, hei*5), imgPdf, 0)
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
