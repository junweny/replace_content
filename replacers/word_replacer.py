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

    def safe_replace_in_paragraphs(paragraphs, replacements, context=""):
        for para in paragraphs:
            try:
                for run in getattr(para, 'runs', []):
                    for old, new in replacements:
                        if old in run.text:
                            run.text = run.text.replace(old, new)
            except Exception as e:
                print(f"处理段落时出错（{context}）：{e}")

    def safe_replace_in_table(table, replacements, context=""):
        for row_idx, row in enumerate(getattr(table, 'rows', [])):
            for cell_idx, cell in enumerate(getattr(row, 'cells', [])):
                try:
                    safe_replace_in_paragraphs(getattr(cell, 'paragraphs', []), replacements, f"表格{context} 行{row_idx} 列{cell_idx}")
                    # 递归处理嵌套表格
                    for nested_table_idx, nested_table in enumerate(getattr(cell, 'tables', [])):
                        safe_replace_in_table(nested_table, replacements, f"{context}-嵌套{nested_table_idx}")
                except Exception as e:
                    print(f"处理表格单元格时出错（{context} 行{row_idx} 列{cell_idx}）：{e}")

    doc = Document(file_path)
    try:
        # 正文
        safe_replace_in_paragraphs(getattr(doc, 'paragraphs', []), replacements, "正文")
        # 页眉和页脚
        for section_idx, section in enumerate(getattr(doc, 'sections', [])):
            header = getattr(section, 'header', None)
            if header:
                safe_replace_in_paragraphs(getattr(header, 'paragraphs', []), replacements, f"页眉{section_idx}")
            footer = getattr(section, 'footer', None)
            if footer:
                safe_replace_in_paragraphs(getattr(footer, 'paragraphs', []), replacements, f"页脚{section_idx}")
        # 表格内容
        for table_idx, table in enumerate(getattr(doc, 'tables', [])):
            safe_replace_in_table(table, replacements, f"主表格{table_idx}")
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
            try:
                replace_in_docx_keep_format(docx_path, replacements)
                docx_to_doc(docx_path, file_path)
            except Exception as e:
                print(f"处理失败：{file_path}，原因：{e}")
                print(traceback.format_exc())
                # 只要出错就尝试删除 docx_path
                if os.path.exists(docx_path):
                    try:
                        os.remove(docx_path)
                    except Exception:
                        pass
                raise
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
            try:
                replace_in_docx_keep_format(docx_path, replacements)
                docx_to_doc(docx_path, file_path)
            except Exception as e2:
                print(f"处理失败：{file_path}，原因：{e2}")
                print(traceback.format_exc())
                if os.path.exists(docx_path):
                    try:
                        os.remove(docx_path)
                    except Exception:
                        pass
                raise
            if os.path.exists(docx_path):
                os.remove(docx_path)
        except Exception as e2:
            print(f"处理失败：{file_path}，原因：{e2}")
            print(traceback.format_exc())
            if 'docx_path' in locals() and os.path.exists(docx_path):
                try:
                    os.remove(docx_path)
                except Exception:
                    pass
            raise

def replace_in_word(file_path, replacements, wildcard=False, fullwidth=False, halfwidth=False, keep_format=True):
    # 只对docx用python-docx保留格式
    if file_path.lower().endswith('.docx') and keep_format:
        replace_in_docx_keep_format(file_path, replacements)
    else:
        replace_in_word_doc(file_path, replacements, wildcard=wildcard, keep_format=keep_format) 