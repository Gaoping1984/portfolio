# Python 办公自动化 · 作品集

这里是我做的几个办公自动化示例项目，**所有示例数据都是虚构的**，代码可以直接运行、复现结果。

## 提供的服务

| 服务 | 价格 | 说明 |
|---|---|---|
| Excel 自动化 / 数据处理 | 49 元起 | 多表合并、数据清洗、汇总报表、图表 |
| Python 办公自动化脚本 | 149 元起 | 批量处理 Word/PDF/文件、按模板生成文档等 |

具体价格按需求复杂度报价，**咨询请在闲鱼私信**。只处理客户自有数据，不接爬取他人网站、刷量等违规需求。

## 示例项目

| 目录 | 解决什么问题 |
|---|---|
| [01-excel-merge-report](01-excel-merge-report/) | 多个 Excel 表合并、去重、清洗，自动生成带图表的汇总报表 |
| [02-word-pdf-tools](02-word-pdf-tools/) | Word 批量转 PDF，PDF 合并、按页拆分 |
| [03-batch-docx-from-template](03-batch-docx-from-template/) | 一个 Word 模板 + 一张 Excel 名单 → 批量生成证书/通知/合同 |
| [04-batch-rename-organize](04-batch-rename-organize/) | 杂乱文件按类型归类并统一命名，生成对照表 |

## 快速运行

```bash
cd 01-excel-merge-report
pip install -r requirements.txt
python make_sample.py      # 生成虚构示例数据
python merge_report.py     # 运行脚本，结果在 output/
```
