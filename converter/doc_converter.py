# converter/doc_converter.py
import os


def is_doc(path: str) -> bool:
    return path.lower().endswith(".doc") and not path.lower().endswith(".docx")


def to_docx(doc_path: str, output_dir: str) -> str:
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    abs_path = os.path.abspath(doc_path)
    out_path = os.path.join(
        os.path.abspath(output_dir),
        os.path.basename(abs_path).replace(".doc", ".docx")
    )
    doc = word.Documents.Open(abs_path)
    doc.SaveAs2(out_path, FileFormat=16)
    doc.Close()
    word.Quit()
    return out_path


def to_doc(docx_path: str, output_dir: str) -> str:
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    abs_path = os.path.abspath(docx_path)
    out_path = os.path.join(
        os.path.abspath(output_dir),
        os.path.basename(abs_path).replace(".docx", ".doc")
    )
    doc = word.Documents.Open(abs_path)
    doc.SaveAs2(out_path, FileFormat=0)
    doc.Close()
    word.Quit()
    return out_path
