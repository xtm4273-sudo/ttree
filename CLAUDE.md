# 万邦船舶询价匹配系统

## 项目概述
万邦船舶的AI询价自动匹配报价系统MVP。上传客户维修询价单(PDF/Word/Excel)，AI解析+工艺库语义匹配，生成报价单。

## 技术栈
- 前端: Streamlit (localhost:8501)
- LLM: 阿里云通义千问 qwen-plus (OpenAI兼容接口)
- Embedding: text-embedding-v3 (同一接口)
- 向量检索: FAISS
- 语言: Python 3.12

## 核心流程
1. 文档解析 (document_parser.py) - pdfplumber提取文本 → LLM智能分段识别维修项目
2. 向量检索 (craft_library.py) - FAISS批量检索工艺库top-K候选
3. LLM精排 (match_engine.py) - 逐条LLM打分 + 置信度加权计算
4. 报价单生成 (excel_generator.py) - 多Sheet Excel输出

## 文件结构
- main.py - Streamlit主界面
- config.py - API密钥和模型参数
- app/document_parser.py - 文档解析
- app/craft_library.py - 工艺库加载/向量化/检索
- app/match_engine.py - AI匹配引擎
- app/excel_generator.py - Excel报价单生成
- app/craft_excel_import.py - 从已填报价单批量学习入库
- data/ - FAISS索引和工艺库JSON

## 工艺库自主学习（人机闭环）
- **目标**：无匹配/待确认条目经人工按主数据填写后写入本地工艺库，更新向量索引，再次匹配即可命中。
- **Streamlit**：匹配完成后页面底部「工艺库自主学习」展开，逐条填写标题/SFI 等后提交；支持单条重匹配并需重新下载报价单。
- **Excel 批量**：报价单主表新增列「学习入库」；填写 **ADD** 并填「人工确认标题」等列后，在侧栏「从报价单 Excel 批量学习」上传导入（见 [app/craft_excel_import.py](app/craft_excel_import.py)）。
- **重建索引不丢学习条目**：侧栏「构建索引」会先读模板 Excel，再与 `data/craft_library.json` 中 `source=user_added` 的条目合并去重后全量向量化（见 `merge_template_with_saved_user_entries`）。
- **权威字段**：SFI/标准工艺名以公司主数据为准；系统不自动把模型猜测写入工艺库。

## 匹配审计（可选）
- 环境变量 `MATCH_LLM_AUDIT_JSONL` 设为 jsonl 路径时，每条匹配追加一行 JSON（门控路径、是否调用 LLM 等），便于统计。

## 启动方式
```bash
cd C:\Users\24895\Desktop\ttree
python -m streamlit run main.py --server.port 8501 --server.headless true
```

## 工艺库来源
C:\Users\24895\Desktop\万邦船舶询价\万邦提供的报价单模板.xlsx

## 已知问题
- 匹配阶段LLM精排是逐条调用，50+条目耗时较长
- 向量检索已优化为批量embedding
