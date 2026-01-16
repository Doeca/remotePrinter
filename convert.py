import os
import platform
import subprocess
import shutil
import time
from logsys import logger

# 条件导入 Windows COM
IS_WINDOWS = platform.system() == 'Windows'
if IS_WINDOWS:
    try:
        import win32com.client
        import pythoncom
        import pywintypes
        HAS_WIN32COM = True
    except ImportError:
        HAS_WIN32COM = False
        logger.warning("win32com not available, will use LibreOffice fallback")
else:
    HAS_WIN32COM = False


def _kill_excel_processes():
    """强制结束所有 Excel 进程（清理僵尸进程）"""
    if not IS_WINDOWS:
        return
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'], 
                      capture_output=True, timeout=5)
    except Exception:
        pass


def _xlsx2pdf_excel(exp, pdfp, retry_count=0):
    """使用 Excel COM 接口转换（仅 Windows）- 增强版"""
    import win32com.client
    import pythoncom
    import pywintypes
    
    excel_path = os.path.abspath(exp)
    pdf_path = os.path.abspath(pdfp)
    
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")
    
    xlApp = None
    books = None
    success = False
    
    # 如果之前失败过，先清理 Excel 进程
    if retry_count > 0:
        logger.warning(f"第 {retry_count + 1} 次尝试，先清理 Excel 进程")
        _kill_excel_processes()
        time.sleep(2)
    
    try:
        # 初始化 COM（在单线程环境中）
        pythoncom.CoInitialize()
        
        # 创建 Excel 应用实例（使用 DispatchEx 创建新实例）
        xlApp = win32com.client.DispatchEx("Excel.Application")
        xlApp.Visible = False
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = False
        xlApp.EnableEvents = False
        xlApp.AskToUpdateLinks = False
        
        # 以只读模式打开工作簿
        books = xlApp.Workbooks.Open(
            excel_path,
            UpdateLinks=0,      # 不更新链接
            ReadOnly=True,      # 只读模式
            IgnoreReadOnlyRecommended=True,
            Notify=False
        )
        
        # 设置页面布局 - 确保内容适应一页
        for sheet in books.Worksheets:
            try:
                # 设置页面方向为横向（如果内容较宽）
                sheet.PageSetup.Orientation = 1  # xlLandscape (横向)
                
                # 缩放以适应页面
                sheet.PageSetup.Zoom = False  # 禁用固定缩放
                sheet.PageSetup.FitToPagesWide = 1  # 适应1页宽
                sheet.PageSetup.FitToPagesTall = False  # 高度不限制（或设为1）
                
                # 设置边距（缩小边距以容纳更多内容）
                sheet.PageSetup.LeftMargin = xlApp.CentimetersToPoints(0.5)
                sheet.PageSetup.RightMargin = xlApp.CentimetersToPoints(0.5)
                sheet.PageSetup.TopMargin = xlApp.CentimetersToPoints(0.5)
                sheet.PageSetup.BottomMargin = xlApp.CentimetersToPoints(0.5)
                
                logger.debug(f"已设置工作表 '{sheet.Name}' 的页面布局")
            except Exception as e:
                logger.warning(f"设置工作表 '{sheet.Name}' 页面布局失败: {e}")
        
        # 确保输出目录存在
        output_dir = os.path.dirname(pdf_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 删除已存在的 PDF
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except Exception as e:
                logger.warning(f"无法删除旧 PDF: {e}")
        
        # 导出为 PDF（使用完整参数）
        books.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=pdf_path,
            Quality=0,  # xlQualityStandard
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        
        success = True
        logger.info(f"Excel 转换成功: {pdf_path}")
        
    except pywintypes.com_error as e:
        logger.error(f"COM 错误: {e}")
        # 如果是首次失败且重试次数少于2次，尝试重试
        if retry_count < 2:
            logger.warning(f"COM 错误，将进行重试...")
            return _xlsx2pdf_excel(exp, pdfp, retry_count + 1)
        raise RuntimeError(f"Excel COM 操作失败: {e}")
        
    except Exception as e:
        logger.error(f"Excel 转换异常: {type(e).__name__}: {e}")
        raise
        
    finally:
        # 清理资源 - 更安全的顺序
        try:
            if books is not None:
                try:
                    books.Saved = True  # 标记为已保存，避免保存提示
                    books.Close(SaveChanges=False)
                except Exception as e:
                    logger.warning(f"关闭工作簿失败: {e}")
                try:
                    del books
                except Exception:
                    pass
                books = None
        except Exception as e:
            logger.warning(f"清理工作簿引用失败: {e}")
        
        try:
            if xlApp is not None:
                try:
                    xlApp.DisplayAlerts = False
                    xlApp.Quit()
                except Exception as e:
                    logger.warning(f"退出 Excel 失败: {e}")
                try:
                    del xlApp
                except Exception:
                    pass
                xlApp = None
        except Exception as e:
            logger.warning(f"清理 Excel 引用失败: {e}")
        
        # 等待 Excel 进程完全退出
        time.sleep(1.5)
        
        # COM 反初始化（放在最后，且增加保护）
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            logger.warning(f"COM 反初始化失败: {e}")
            # 忽略反初始化错误，不影响结果
    
    # 验证输出
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"PDF 未生成: {pdf_path}")
    
    return success


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


def xlsx2pdf(exp, pdfp, force_excel=True):
    """
    将 Excel 文件转换为 PDF
    
    Args:
        exp: Excel 文件路径
        pdfp: 输出 PDF 路径
        force_excel: 是否强制使用 Excel（Windows 下）。
                     True = 只用 Excel，失败则抛异常
                     False = Excel 失败后降级到 LibreOffice
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
            if force_excel:
                logger.error(f"Excel 转换失败且 force_excel=True: {e}")
                raise
            else:
                logger.warning(f"Excel 转换失败: {e}, 尝试使用 LibreOffice")
    
    # 使用 LibreOffice 作为后备方案
    logger.info(f"使用 LibreOffice 转换 {excel_path} -> {pdf_path}")
    _xlsx2pdf_libreoffice(exp, pdfp)
