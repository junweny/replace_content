from docx import Document
import re
from utils.unicode_utils import to_fullwidth, to_halfwidth
import win32com.client
import pythoncom
import os
import time
import shutil
import tempfile
import re
import traceback

def replace_in_docx_keep_format(file_path, replacements):
    from docx.text.paragraph import Paragraph
    from docx.table import _Cell, Table

    doc = Document(file_path)
    try:
        # 正文
        for para in getattr(doc, 'paragraphs', []):
            for run in getattr(para, 'runs', []):
                for old, new in replacements:
                    if old in run.text:
                        run.text = run.text.replace(old, new)
        # 页眉和页脚
        for section in getattr(doc, 'sections', []):
            # 页眉
            header = getattr(section, 'header', None)
            if header:
                for para in getattr(header, 'paragraphs', []):
                    for run in getattr(para, 'runs', []):
                        for old, new in replacements:
                            if old in run.text:
                                run.text = run.text.replace(old, new)
            # 页脚
            footer = getattr(section, 'footer', None)
            if footer:
                for para in getattr(footer, 'paragraphs', []):
                    for run in getattr(para, 'runs', []):
                        for old, new in replacements:
                            if old in run.text:
                                run.text = run.text.replace(old, new)
        # 表格内容
        for table in getattr(doc, 'tables', []):
            for row in getattr(table, 'rows', []):
                for cell in getattr(row, 'cells', []):
                    for para in getattr(cell, 'paragraphs', []):
                        for run in getattr(para, 'runs', []):
                            for old, new in replacements:
                                if old in run.text:
                                    run.text = run.text.replace(old, new)
        doc.save(file_path)
    except Exception as e:
        print(f'处理docx内容时出错：{file_path}，原因：{e}')
        raise

def doc_to_docx(doc_path):
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    docx_path = doc_path + 'x' if not doc_path.lower().endswith('.docx') else doc_path
    try:
        doc = word.Documents.Open(os.path.normpath(os.path.abspath(doc_path)))
        if doc is None:
            print(f"处理失败：{doc_path}，原因：Word无法打开文档，可能路径、权限或文件损坏")
            if os.path.exists(docx_path):
                try:
                    os.remove(docx_path)
                except Exception:
                    pass
            return docx_path
        try:
            doc.SaveAs(docx_path, FileFormat=16)  # 16=wdFormatDocumentDefault (.docx)
        except Exception as e:
            print(f"处理失败：{doc_path}，原因：SaveAs失败: {e}")
            doc.Close()
            if os.path.exists(docx_path):
                try:
                    os.remove(docx_path)
                except Exception:
                    pass
            return docx_path
        doc.Close()
    except Exception as e:
        # 只在处理失败时输出
        print(f"处理失败：{doc_path}，原因：{e}")
        if os.path.exists(docx_path):
            try:
                os.remove(docx_path)
            except Exception:
                pass
        raise
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return docx_path

def docx_to_doc(docx_path, doc_path):
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(os.path.normpath(os.path.abspath(docx_path)))
        if doc is None:
            print(f"处理失败：{docx_path}，原因：Word无法打开文档，可能路径、权限或文件损坏")
            if os.path.exists(docx_path):
                try:
                    os.remove(docx_path)
                except Exception:
                    pass
            return
        try:
            doc.SaveAs(doc_path, FileFormat=0)  # 0=wdFormatDocument (.doc)
        except Exception as e:
            print(f"处理失败：{docx_path}，原因：SaveAs失败: {e}")
            doc.Close()
            if os.path.exists(docx_path):
                try:
                    os.remove(docx_path)
                except Exception:
                    pass
            return
        doc.Close()
    except Exception as e:
        # 只在处理失败时输出
        print(f"处理失败：{docx_path}，原因：{e}")
        if os.path.exists(docx_path):
            try:
                os.remove(docx_path)
            except Exception:
                pass
        raise
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def replace_in_word_doc(file_path, replacements, wildcard=False, keep_format=True):
    try:
        if keep_format:
            docx_path = doc_to_docx(file_path)
            replace_in_docx_keep_format(docx_path, replacements)
            docx_to_doc(docx_path, file_path)
            if os.path.exists(docx_path):
                os.remove(docx_path)
        else:
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = 0
            orig_path = os.path.normpath(os.path.abspath(file_path))
            temp_path = orig_path
            temp_path = os.path.normpath(temp_path).replace('/', '\\')
            try:
                for i in range(3):
                    try:
                        doc = word.Documents.Open(temp_path)
                        break
                    except Exception as e:
                        if i == 2:
                            raise
                        time.sleep(1)
                for old, new in replacements:
                    find = doc.Content.Find
                    find.Text = old
                    find.Replacement.Text = new
                    find.Forward = True
                    find.Wrap = 1  # wdFindContinue
                    find.MatchWildcards = wildcard
                    find.Execute(Replace=2)  # wdReplaceAll
                doc.Save()
                doc.Close()
            finally:
                try:
                    word.Quit()
                except Exception:
                    pass
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
    except Exception as e:
        print(f"处理失败：{file_path}，原因：{e}")
        try:
            docx_path = doc_to_docx(file_path)
            replace_in_docx_keep_format(docx_path, replacements)
            docx_to_doc(docx_path, file_path)
            if os.path.exists(docx_path):
                os.remove(docx_path)
        except Exception as e2:
            print(f"处理失败：{file_path}，原因：{e2}")
            print(traceback.format_exc())
            raise

def replace_in_word(file_path, replacements, wildcard=False, fullwidth=False, halfwidth=False, keep_format=True):
    # 只对docx用python-docx保留格式
    if file_path.lower().endswith('.docx') and keep_format:
        replace_in_docx_keep_format(file_path, replacements)
    else:
        replace_in_word_doc(file_path, replacements, wildcard=wildcard, keep_format=keep_format) 