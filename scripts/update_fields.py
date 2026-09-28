# -*- coding: utf-8 -*-
"""
word-deliverable / update_fields.py
更新 .docx 的 TOC 与所有字段（PAGE/NUMPAGES 等），不导出 PDF。

用法:
    python update_fields.py <file.docx>
    python update_fields.py <file.docx> --libreoffice

优先级:
  1. Microsoft Word COM（WD_NO_WORD=1 或 --libreoffice 时跳过）
  2. LibreOffice UNO（需 LibreOffice 自带 Python）
  3. LibreOffice headless convert-to docx（基本字段更新，TOC 可能不更新）

注意:
  Word COM 更新字段时会"规范化" numbering.xml，可能清空多级编号的非标准级别
  （如 isLgl=True 的 ilvl 1-8），导致 H2/H3/H4 自动编号消失。
  含多级编号的文档应使用 --libreoffice 或 WD_NO_WORD=1。
  UNO 路径完成后会自动修复非首 section 的页眉页脚链接（移除独立的
  headerReference/footerReference，使其继承前节），通过 validate 的链接检查。
"""
import os
import sys
import shutil
import subprocess
import time

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
LO_PYTHON = r"C:\Program Files\LibreOffice\program\python.exe"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _com_retry(fn, retries=5, delay=1.5):
    last = None
    for i in range(retries):
        try:
            return fn()
        except Exception as e:
            last = e
            if i < retries - 1:
                time.sleep(delay * (i + 1))
            continue
    raise last if last else RuntimeError("COM retry exhausted")


def update_with_word(abs_path):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        # 使用独立 Word 实例，避免连接到用户正在查看、受保护模式或弹窗阻塞的会话。
        # Dispatch/EnsureDispatch 可能复用现有实例并触发 RPC_E_CALL_REJECTED。
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception:
            try:
                word = win32com.client.gencache.EnsureDispatch("Word.Application")
            except Exception:
                word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        doc = _com_retry(lambda: word.Documents.Open(abs_path))
        try:
            _com_retry(lambda: [t.Update() for t in doc.TablesOfContents])
        except Exception as e:
            print(f"  TOC.Update 警告: {e}")
        try:
            _com_retry(lambda: doc.Fields.Update())
        except Exception as e:
            print(f"  Fields.Update 警告: {e}")
        try:
            for sec in doc.Sections:
                for headers in (sec.Headers, sec.Footers):
                    for h in headers:
                        try:
                            h.Range.Fields.Update()
                        except Exception:
                            pass
        except Exception:
            pass
        try:
            _com_retry(lambda: doc.Repaginate())
        except Exception:
            pass
        _com_retry(lambda: doc.Save())
        return True
    finally:
        if doc is not None:
            try:
                _com_retry(lambda: doc.Close(SaveChanges=False), retries=3, delay=1.0)
            except Exception:
                pass
        if word is not None:
            try:
                _com_retry(lambda: word.Quit(), retries=3, delay=1.0)
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def update_with_uno(abs_path):
    """用 LibreOffice UNO 更新 TOC/字段并保存（不导出 PDF）。"""
    if not os.path.exists(LO_PYTHON):
        return False
    # 内联 UNO 脚本（复用 render_pdf_uno.py 的逻辑但只保存不导出）
    inline = """
import os, sys, time, subprocess, socket
SOFFICE = r"C:\\Program Files\\LibreOffice\\program\\soffice.exe"
def _free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p
docx_path = os.path.abspath(sys.argv[1])
if not os.path.exists(docx_path):
    print(f"文件不存在: {docx_path}"); sys.exit(1)
import uno
from com.sun.star.beans import PropertyValue
port = _free_port()
proc = subprocess.Popen([SOFFICE, "--headless", "--invisible", "--nocrashreport",
    "--nodefault", "--nologo", "--nofirststartwizard", "--norestore",
    f"--accept=socket,host=127.0.0.1,port={port};urp;"],
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)
    ctx = None
    for i in range(30):
        try:
            ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(0.5)
    if ctx is None:
        print("错误: 无法连接 UNO"); sys.exit(2)
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    url = uno.systemPathToFileUrl(docx_path)
    hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
    doc = desktop.loadComponentFromURL(url, "_blank", 0, (hidden,))
    if doc is None:
        print("错误: 无法打开文档"); sys.exit(3)
    try: doc.refresh()
    except Exception as e: print(f"  refresh 警告: {e}")
    try:
        indexes = doc.getDocumentIndexes()
        for i in range(indexes.getCount()):
            indexes.getByIndex(i).update()
    except Exception as e: print(f"  indexes 警告: {e}")
    try: doc.refresh()
    except Exception: pass
    # 保存为 docx
    save_props = PropertyValue(); save_props.Name = "FilterName"; save_props.Value = "MS Word 2007 XML"
    doc.storeToURL(url, (save_props,))
    doc.close(True)
    print("OK UNO 更新完成")
finally:
    try: desktop.terminate()
    except Exception: pass
    try: proc.wait(timeout=10)
    except Exception: proc.kill()
"""
    tmp_script = os.path.join(SCRIPT_DIR, "_uno_update_tmp.py")
    try:
        with open(tmp_script, "w", encoding="utf-8") as f:
            f.write(inline)
        result = subprocess.run(
            [LO_PYTHON, tmp_script, abs_path],
            check=True, timeout=120, capture_output=True, text=True,
        )
        print(result.stdout.strip())
        return True
    except subprocess.CalledProcessError as e:
        print(f"  UNO 失败: {e}")
        if e.stderr:
            print(f"  stderr: {e.stderr[:500]}")
        return False
    except Exception as e:
        print(f"  UNO 失败: {e}")
        return False
    finally:
        try:
            os.remove(tmp_script)
        except Exception:
            pass


def update_with_libreoffice(abs_path):
    if not os.path.exists(SOFFICE):
        return False
    outdir = os.path.dirname(abs_path)
    tmpdir = os.path.join(outdir, "_lo_tmp")
    os.makedirs(tmpdir, exist_ok=True)
    try:
        subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "docx:MS Word 2007 XML",
             "--outdir", tmpdir, abs_path],
            check=True, timeout=180, capture_output=True,
        )
        result = os.path.join(tmpdir, os.path.basename(abs_path))
        if os.path.exists(result):
            shutil.copy2(result, abs_path)
            return True
    except Exception as e:
        print(f"  LibreOffice 失败: {e}")
    finally:
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass
    return False


def _fix_section_headers(abs_path):
    """UNO 保存后非首 section 可能引入独立的 headerReference/footerReference，
    导致 validate 报"页眉页脚未链接前节"。移除非首 section 的这些引用，使其继承前节。"""
    try:
        from docx import Document
        from docx.oxml.ns import qn
        doc = Document(abs_path)
        sections = doc.sections
        fixed = []
        for i, section in enumerate(sections):
            if i == 0:
                continue
            sectPr = section._sectPr
            removed = 0
            for ref_tag in ("w:headerReference", "w:footerReference"):
                for ref in sectPr.findall(qn(ref_tag)):
                    sectPr.remove(ref)
                    removed += 1
            if removed:
                fixed.append(f"Section#{i+1}:{removed}")
        if fixed:
            doc.save(abs_path)
            print(f"  section 页眉链接修复: {', '.join(fixed)}")
    except Exception as e:
        print(f"  section 页眉链接修复跳过: {e}")


def main():
    if len(sys.argv) < 2:
        print("用法: python update_fields.py <file.docx> [--libreoffice]")
        sys.exit(1)
    path = os.path.abspath(sys.argv[1])
    if not os.path.exists(path):
        print(f"文件不存在: {path}")
        sys.exit(1)
    print(f"更新字段: {path}")
    force_lo = "--libreoffice" in sys.argv or os.environ.get("WD_NO_WORD", "") == "1"
    ok = False

    # 1. Word COM
    if not force_lo:
        try:
            ok = update_with_word(path)
            if ok:
                print("OK Word COM 更新完成")
        except Exception as e:
            print(f"Word COM 失败: {e}")

    # 2. LibreOffice UNO
    if not ok:
        if not force_lo:
            print("尝试 UNO ...")
        ok = update_with_uno(path)
        if ok:
            # UNO 保存后修复 section 页眉链接（Word COM 路径不需要）
            _fix_section_headers(path)
            return

    # 3. LibreOffice 简单转换
    if not ok:
        print("回退 LibreOffice ...")
        if update_with_libreoffice(path):
            print("OK LibreOffice 更新完成")
        else:
            print("失败: 所有方式均不可用或未成功")
            sys.exit(2)


if __name__ == "__main__":
    main()
