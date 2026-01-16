import requests
import dingLib
import pdfmod
import json
import random
import string
from logsys import logger
characters = string.ascii_letters + string.digits


def img(images: list, pid, link):
    req = requests.get(link)
    randomdid = ''.join(random.choices(characters, k=8))
    # 元のファイル拡張子を保持
    import os
    original_ext = os.path.splitext(link)[1] or '.png'
    filename = f"./cache/{pid}_{randomdid}{original_ext}"
    logger.debug(f"本地文件名：{filename} download link: {link}")
    with open(filename, "wb") as f:
        f.write(req.content)
    images.append(filename)


def pdf(images: list, pid, link):
    req = requests.get(link)
    randomdid = ''.join(random.choices(characters, k=8))
    filename = f"./cache/{pid}_{link[-20:]}_{randomdid}.pdf"
    with open(filename, "wb") as f:
        f.write(req.content)
    for d in pdfmod.pdfexport(filename, f"{filename}_export"):
        images.append(d)


def singleat(images: list, pid, data):
    file_id = data.get('fileId')
    if not file_id:
        logger.warning(f"附件数据缺少 fileId: {data}")
        return
    
    atdata = dingLib.getAttachment(pid, file_id)
    if atdata == None:
        return
    duri = atdata.download_uri
    filetype = data.get('fileType', 'unknown')
    if (filetype == 'pdf'):
        pdf(images, pid, duri)
    elif (filetype == 'png'):
        img(images, pid, duri)
    elif filetype == 'odf':
        logger.debug(f"unsupported file:{duri}\n")
    elif filetype == 'mp4':
        logger.debug(f"unsupported file:{duri}\n")
    elif filetype == 'xls':
        logger.debug(f"unsupported file:{duri}\n")
    else:
        file_name = data.get('fileName', 'unknown')
        logger.debug(f"unsupported file:{duri},file_type:{filetype},file_name:{file_name}\n")
        img(images, pid, duri)


def subattachments(images: list, data, pid):
    for tmp in data:
        row_values = tmp.get('rowValue', [])
        for tmp_rowValue in row_values:
            key = tmp_rowValue.get('key', '')
            if tmp_rowValue.get('value') == None:
                continue
            if key.find('DDPhotoField') != -1:
                for vl in tmp_rowValue.get('value', []):
                    img(images, pid, vl)
            elif key.find('DDAttachment') != -1:
                for vr in tmp_rowValue.get('value', []):
                    singleat(images, pid, vr)
            elif key.find('InvoiceField') != -1:
                # 2024/10/22，新增发票列
                rdata = tmp_rowValue.get('extendValue')
                if not rdata:
                    logger.warning(f"InvoiceField 缺少 extendValue")
                    continue
                
                try:
                    if type(rdata) == str:
                        parsed_data = json.loads(s=rdata)
                        invoices = parsed_data.get('invoiceList', [])
                    else:
                        invoices = rdata.get('invoiceList', [])

                    for invoice in invoices:
                        pdfurl = invoice.get('invoicePdfUrl', '')
                        imgurl = invoice.get('invoiceImgUrl', '')
                        if pdfurl != '':
                            pdf(images, pid, pdfurl)
                        elif imgurl != '':
                            img(images, pid, imgurl)
                except Exception as e:
                    logger.error(f"处理发票数据失败: {e}, 数据: {rdata}")
                        
            elif key.find("TableField") != -1:
                table_value = tmp_rowValue.get('value')
                if table_value:
                    subattachments(images, table_value, pid)


def attachments(images: list, data, pid):
    for vd in data.form_component_values:
        id: str = vd.component_type
        if vd.value == None:
            continue
        if id.find('DDPhotoField') != -1:
            try:
                tmp = json.loads(s=vd.value)
                for vl in tmp:
                    img(images, pid, vl)
            except Exception as e:
                logger.error(f"处理图片字段失败: {e}")
        elif id.find('DDAttachment') != -1:
            try:
                tmp = json.loads(s=vd.value)
                for vr in tmp:
                    singleat(images, pid, vr)
            except Exception as e:
                logger.error(f"处理附件字段失败: {e}")
        elif id.find('InvoiceField') != -1:
            # 2024/10/22，新增发票列
            if vd.ext_value == None:
                continue
            try:
                if type(vd.ext_value) == str:
                    invoices = json.loads(s=vd.ext_value)
                else:
                    invoices = vd.ext_value
                for invoice in invoices:
                    invoice_pdf_url = invoice.get('invoicePdfUrl')
                    if invoice_pdf_url:
                        pdf(images, pid, invoice_pdf_url)
                    else:
                        logger.warning(f"发票缺少 invoicePdfUrl: {invoice}")
            except Exception as e:
                logger.error(f"处理发票字段失败: {e}")
        elif id.find("TableField") != -1:
            # 子组件中的附件
            try:
                tmp = json.loads(s=vd.value)
                subattachments(images, tmp, pid)
            except Exception as e:
                logger.error(f"处理表格字段失败: {e}")
