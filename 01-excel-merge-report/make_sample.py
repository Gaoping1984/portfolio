"""生成虚构示例数据：3 个门店的月度销售表（含常见脏数据）。"""
import random, pandas as pd
from pathlib import Path
random.seed(1)
out = Path(__file__).parent / "sample_input"; out.mkdir(exist_ok=True)
products = ["笔记本", "签字笔", "文件夹", "打印纸", "订书机"]
for store in ["东区店", "西区店", "南区店"]:
    rows = []
    for i in range(30):
        rows.append({"日期": f"2026-0{random.randint(1,3)}-{random.randint(1,28):02d}",
                     "商品": random.choice(products) + random.choice(["", " "]),
                     "数量": random.choice([random.randint(1, 50), random.randint(1, 50), None]),
                     "单价": random.choice([3.5, 12, 8.8, 25, 18])})
    rows += rows[:3]  # 重复行
    pd.DataFrame(rows).to_excel(out / f"{store}_销售.xlsx", index=False)
print("示例数据已生成")
