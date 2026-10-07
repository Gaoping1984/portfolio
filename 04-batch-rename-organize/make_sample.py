"""生成虚构的杂乱文件夹。"""
from pathlib import Path
import os, time
out = Path(__file__).parent / "sample_input"; out.mkdir(exist_ok=True)
names = ["IMG_0012.jpg", "IMG_0013.JPG", "微信图片_20260301.png", "报价单 最终版(2).xlsx", "报价单最终版.xlsx",
         "合同 扫描件.pdf", "发票-三月.pdf", "会议记录.docx", "新建文本文档.txt", "演示稿 v3.pptx", "录音001.mp3", "未知文件.xyz"]
for i, n in enumerate(names):
    p = out / n; p.write_text(f"示例文件 {n}", encoding="utf-8")
    t = time.mktime((2026, 1 + i % 3, 5 + i, 10, 0, 0, 0, 0, -1)); os.utime(p, (t, t))
print("示例文件已生成")
