"""文件批量重命名 + 按类型归类。命名规则: 类型/日期_序号_原名清理.后缀
默认先预览(--dry-run)，加 --apply 才执行；执行时复制到输出目录（不动原文件），并生成对照表 CSV。
用法: python organize.py [输入目录] [输出目录] [--apply]
"""
import csv, re, shutil, sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

TYPES = {"图片": {".jpg", ".jpeg", ".png", ".gif"}, "表格": {".xlsx", ".xls", ".csv"},
         "文档": {".docx", ".doc", ".txt", ".pdf"}, "演示": {".pptx", ".ppt"}, "音频": {".mp3", ".wav", ".m4a"}}

def kind(ext):
    return next((k for k, v in TYPES.items() if ext in v), "其他")

def clean(stem):
    s = re.sub(r"[\s()（）\[\]]+", "_", stem).strip("_")
    return re.sub(r"_+", "_", s)

def main(*argv):
    args = [a for a in argv if not a.startswith("--")]
    apply = "--apply" in argv
    base = Path(__file__).parent
    src = Path(args[0]) if args else base / "sample_input"
    dst = Path(args[1]) if len(args) > 1 else base / "output"
    counter, plan = defaultdict(int), []
    for f in sorted(p for p in src.iterdir() if p.is_file()):
        ext = f.suffix.lower(); k = kind(ext)
        d = datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y%m%d")
        counter[k] += 1
        plan.append((f, dst / k / f"{d}_{counter[k]:03d}_{clean(f.stem)}{ext}"))
    for a, b in plan: print(f"{a.name:28s} -> {b.relative_to(dst)}")
    if not apply:
        print(f"\n[预览] 共 {len(plan)} 个文件，加 --apply 执行"); return
    for a, b in plan:
        b.parent.mkdir(parents=True, exist_ok=True); shutil.copy2(a, b)
    with open(dst / "重命名对照表.csv", "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.writer(fp); w.writerow(["原文件名", "新路径"])
        w.writerows((a.name, str(b.relative_to(dst))) for a, b in plan)
    print(f"\n已处理 {len(plan)} 个文件 -> {dst}")

if __name__ == "__main__":
    main(*sys.argv[1:])
