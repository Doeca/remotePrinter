#  本文件处理：合作档口营业款付款申请
import os
import json
import dingLib
import requests
import shutil
import openpyxl
import printmod
import pdfmod
import handle
import convert
from openpyxl.styles import Alignment, Font
from logsys import logger

alfont = Font(name="仿宋", size=12)
tifont = Font(name="仿宋", size=14, bold=True)
alstyle = Alignment(wrap_text=True, vertical='top')


def get_value(raw: list, key: str, id: str = ''):
    for v in raw:
        if v.name == key and id == '':
            return v.value
        elif v.id == id:
            return v.value
    return ""


def handle_xlsx(data, file):
    # 修改表格数据
    workbook = openpyxl.load_workbook(file)
    worksheet = workbook["Sheet1"]
    worksheet['B2'].value = str(data.business_id)  # 审批编号
    worksheet['B3'].value = data.title.replace("提交的合作档口营业款付款申请", "")  # 创建人
    worksheet['B4'].value = data.originator_dept_name  # 创建部门
    worksheet['B5'].value = get_value(data.form_component_values, "企业主体")
    worksheet['B6'].value = get_value(
        data.form_component_values, "请选择归属战区")  # 归属战区
    worksheet['B7'].value = get_value(data.form_component_values, "门店名称")
    worksheet['B8'].value = get_value(data.form_component_values, "营业款结算月份")

    # 2024.10.22新增，钉钉智能组件
    subFields = json.loads(s=get_value(data.form_component_values, "报销明细"))
    field_names = []
    field_moneys = []
    field_invoices = []
    for i in range(0, len(subFields)):
        row_values = subFields[i].get('rowValue', [])
        for subData in row_values:
            if subData.get('label', "") == '档口名称':
                value = subData.get('value')
                if value is not None:
                    field_names.append(str(value))
            if subData.get('label', "") == '报销金额（元）':
                value = subData.get('value')
                if value is not None:
                    field_moneys.append(str(value))
            if subData.get('label', "") == '发票':
                rdata = subData.get('extendValue')
                if not rdata:
                    continue
                try:
                    if type(rdata) == str:
                        parsed_data = json.loads(s=rdata)
                        invoices = parsed_data.get('invoiceList', [])
                    else:
                        invoices = rdata.get('invoiceList', [])
                    for invoice in invoices:
                        invoice_no = invoice.get('invoiceNo')
                        if invoice_no:
                            field_invoices.append(invoice_no)
                except Exception as e:
                    logger.error(f"处理发票数据失败: {e}")
    worksheet['B9'].value = "、".join(field_names)
    worksheet['B10'].value = "、".join(field_moneys)
    worksheet['B11'].value = "、".join(field_invoices)
    # 2024.10.22新增，钉钉智能组件

    worksheet['B12'].value = get_value(
        data.form_component_values, "报销总额（元）")

    if worksheet['B12'].value != None:
        worksheet['B12'].value += "元"

    val_fksqd = get_value(data.form_component_values, "付款申请单")
    if val_fksqd == None:
        val_fksqd = ""
    worksheet['B13'].value = val_fksqd + \
        get_value(data.form_component_values, "图片")
    worksheet['B14'].value = get_value(data.form_component_values, "附件")
    # 从18开始自行添加
    st_id = 18

    # 审批流程
    for vd in data.operation_records:
        if vd.result != 'AGREE':
            continue
        worksheet.merge_cells(f'A{st_id}:B{st_id}')
        line = ''
        if (vd.remark != ''):
            worksheet[f'A{st_id}'].font = alfont
            worksheet[f'A{st_id}'].alignment = alstyle
            line = str(vd.remark) + "\n"
        line += dingLib.getUserName(vd.user_id) + "  已同意    "+vd.date
        worksheet[f'A{st_id}'].value = line
        st_id += 1

    # 保存文件
    workbook.save(filename=file)


def handle_all(pid, status: str = 'COMPLETED', title=''):
    images = []
    data = dingLib.getDetail(pid)
    if status == 'COMPLETED':
        # 将xlsx转成pdf，然后导出为图片，添加到images里
        shutil.copyfile("./template/fksq_new.xlsx",
                        f"./cache/{pid}_fksq_new.xlsx")
        handle_xlsx(data, f"./cache/{pid}_fksq_new.xlsx")
        convert.xlsx2pdf(
            f"./cache/{pid}_fksq_new.xlsx", f"./cache/{pid}_temp.pdf")
        filename = f"./cache/{pid}_temp.pdf"
        for d in pdfmod.pdfexport(filename, f"{filename}_export"):
            images.append(d)
        # 把images merge成一个pdf
        pdfmod.mergeImages(f"./files/{pid}.pdf", images, 210, 297)
        # 递交打印
        printmod.a4_print(os.path.abspath(f"./files/{pid}.pdf"))
        logger.critical(f"【表单】{title} - {pid}.pdf 提交打印")
    else:
        continue_handle = False
        for vd in data.operation_records:
            if vd.show_name == '店经理' and vd.result == 'AGREE':
                continue_handle = True
                break
        if not continue_handle:
            logger.info("[Task3] 店经理暂未同意，不继续处理")
            return False
        # handle_attachments，收集整个表单中的附件，添加到images
        handle.attachments(images, data, pid)
        # 把images merge成一个pdf
        pdfmod.mergeImages(f"./files/{pid}_attachments.pdf", images, 210, 297)
        # 递交打印
        printmod.a4_print_singleside(
            os.path.abspath(f"./files/{pid}_attachments.pdf"))
        logger.critical(f"【附件】{title} - {pid}_attachments.pdf 提交打印")

    return True
