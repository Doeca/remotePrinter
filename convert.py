import os
import platform
import subprocess
import shutil
from logsys import logger

# 条件导入 Windows COM
IS_WINDOWS = platform.system() == 'Windows'
if IS_WINDOWS:
    try:
        from win32com.client import DispatchEx
        HAS_WIN32COM = True
    except ImportError:
        HAS_WIN32COM = False
        logger.warning("win32com not available, will use LibreOffice fallback")
else:
    HAS_WIN32COM = False


def _xlsx2pdf_excel(exp, pdfp):
    """使用 Excel COM 接口转换（仅 Windows）"""
    excel_path = os.path.abspath(exp)
    pdf_path = os.path.abspath(pdfp)
    xlApp = DispatchEx("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = 0
    books = xlApp.Workbooks.Open(excel_path, False)
    books.ExportAsFixedFormat(0, pdf_path)
    books.Close(False)
    xlApp.Quit()


def _xlsx2pdf_libreoffice(exp, pdfp):
    """使用 LibreOffice 转换（跨平台）"""
    excel_path = os.path.abspath(exp)
    pdf_path = os.path.abspath(pdfp)
    output_dir = os.path.dirname(pdf_path)
    
    # 查找 LibreOffice 可执行文件
    libreoffice_paths = []
    if IS_WINDOWS:
        libreoffice_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    else:  # macOS 或 Linux
        libreoffice_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/usr/bin/libreoffice",
            "/usr/local/bin/libreoffice",
        ]
    
    soffice_cmd = None
    for path in libreoffice_paths:
        if os.path.exists(path):
            soffice_cmd = path
            break
    
    if not soffice_cmd:
        # 尝试从 PATH 中查找
        soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
    
    if not soffice_cmd:
        raise RuntimeError(
            "LibreOffice not found. Please install LibreOffice or use Windows with Excel."
        )
    
    # 执行转换
    cmd = [
        soffice_cmd,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        excel_path
    ]
    
    logger.info(f"使用 LibreOffice 转换: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
    
    # LibreOffice 会生成与源文件同名的 PDF，可能需要重命名
    temp_pdf = os.path.join(output_dir, os.path.splitext(os.path.basename(excel_path))[0] + ".pdf")
    if temp_pdf != pdf_path and os.path.exists(temp_pdf):
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        os.rename(temp_pdf, pdf_path)


def xlsx2pdf(exp, pdfp):
    """
    将 Excel 文件转换为 PDF
    
    优先使用 Excel COM（Windows），否则使用 LibreOffice（跨平台）
    """
    excel_path = os.path.abspath(exp)
    pdf_path = os.path.abspath(pdfp)
    
    # Windows 下优先使用 Excel
    if IS_WINDOWS and HAS_WIN32COM:
        try:
            logger.info(f"使用 Excel 转换 {excel_path} -> {pdf_path}")
            _xlsx2pdf_excel(exp, pdfp)
            return
        except Exception as e:
            logger.warning(f"Excel 转换失败: {e}, 尝试使用 LibreOffice")
    
    # 使用 LibreOffice 作为后备方案
    logger.info(f"使用 LibreOffice 转换 {excel_path} -> {pdf_path}")
    _xlsx2pdf_libreoffice(exp, pdfp)
