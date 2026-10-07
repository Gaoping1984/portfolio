"""Word 批量转 PDF / PDF 合并 / PDF 拆分。
用法:
  python pdf_tools.py convert <docx目录> <输出目录>      # 需要 LibreOffice(Linux/Mac) 或 Word(Windows, 用 docx2pdf)
  python pdf_tools.py merge <输出.pdf> <a.pdf> <b.pdf> ...
  python pdf_tools.py split <输入.pdf> <输出目录>          # 每页一个文件
  python pdf_tools.py demo                               # 用示例数据跑完整流程
"""
import shutil, subprocess, sys
from pathlib import Path
from pypdf import PdfReader, PdfWriter

def convert(src, dst):
    src, dst = Path(src), Path(dst); dst.mkdir(parents=True, exist_ok=True)
    files = sorted(src.glob("*.docx"))
    office = shutil.which("soffice") or shutil.which("libreoffice")
    if office:
        subprocess.run([office, "--headless", "--convert-to", "pdf", "--outdir", str(dst), *map(str, files)],
                       check=True, capture_output=True)
    else:
        from docx2pdf import convert as d2p  # Windows/Mac 装了 Word 时
        for f in files: d2p(str(f), str(dst / (f.stem + ".pdf")))
    pdfs = [dst / (f.stem + ".pdf") for f in files]
    print(f"已转换 {sum(p.exists() for p in pdfs)}/{len(files)} 个文件")
    return pdfs

def merge(out, *pdfs):
    w = PdfWriter()
    for p in pdfs:
        for page in PdfReader(p).pages: w.add_page(page)
    with open(out, "wb") as f: w.write(f)
    print(f"已合并 {len(pdfs)} 个文件 -> {out}（{len(w.pages)} 页）")

def split(pdf, dst):
    dst = Path(dst); dst.mkdir(parents=True, exist_ok=True)
    r = PdfReader(pdf)
    for i, page in enumerate(r.pages, 1):
        w = PdfWriter(); w.add_page(page)
        with open(dst / f"{Path(pdf).stem}_第{i}页.pdf", "wb") as f: w.write(f)
    print(f"已拆分为 {len(r.pages)} 个文件 -> {dst}")

def demo():
    base = Path(__file__).parent
    pdfs = convert(base / "sample_input", base / "output/pdf")
    merge(base / "output/合并结果.pdf", *pdfs)
    split(base / "output/合并结果.pdf", base / "output/拆分结果")

if __name__ == "__main__":
    cmd, *args = sys.argv[1:] or ["demo"]
    {"convert": convert, "merge": merge, "split": split, "demo": demo}[cmd](*args)
