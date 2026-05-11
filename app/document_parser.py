"""
文档解析模块 - 支持 PDF/Word/Excel，使用AI智能分段
不依赖固定格式、固定页码、固定编码体系
"""
import os
import sys
import json
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from openai import OpenAI

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config


def parse_document(file_path: str, log_callback=None) -> list[dict]:
    """
    解析任意格式的询价单文档。

    支持: PDF, Word(.docx), Excel(.xlsx/.xls)

    返回:
    [
        {
            "title": "Crane Service | Shore Crane Hourly Usage",
            "description": "Please provide quotation for crane service...",
            "sfi_code": "1.EH.1.1" 或 None,
            "quantity": 3 或 None,
            "unit": "HR" 或 "",
        },
        ...
    ]
    """
    ext = os.path.splitext(file_path)[1].lower()

    _log = log_callback or (lambda msg: None)

    _log(f"检测到文件格式: {ext}")
    if ext == ".pdf":
        raw_text = _extract_pdf(file_path)
    elif ext in (".docx", ".doc"):
        raw_text = _extract_word(file_path)
    elif ext in (".xlsx", ".xls"):
        raw_text = _extract_excel(file_path)
    else:
        raise ValueError(f"不支持的文件格式: {ext}")

    _log(f"文本提取完成，共 {len(raw_text)} 字符")

    # 用AI做智能分段
    items = _ai_segment(raw_text, log_callback=_log)
    return items


def _extract_pdf(path: str) -> str:
    """从PDF提取全部文本"""
    import pdfplumber

    texts = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                texts.append(text)
    return "\n\n".join(texts)


def _extract_word(path: str) -> str:
    """从Word文档提取全部文本"""
    from docx import Document

    doc = Document(path)
    texts = []
    for para in doc.paragraphs:
        if para.text.strip():
            texts.append(para.text)

    # 也提取表格内容
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if row_text:
                texts.append(row_text)

    return "\n".join(texts)


def _extract_excel(path: str) -> str:
    """从Excel提取全部文本"""
    import pandas as pd

    dfs = pd.read_excel(path, sheet_name=None, dtype=str)
    texts = []
    for sheet_name, df in dfs.items():
        texts.append(f"=== Sheet: {sheet_name} ===")
        # 转为文本表格
        for _, row in df.iterrows():
            row_text = " | ".join(str(v) for v in row.values if pd.notna(v) and str(v).strip())
            if row_text:
                texts.append(row_text)

    return "\n".join(texts)


def _ai_segment(raw_text: str, log_callback=None) -> list[dict]:
    """
    用LLM对原始文本做智能分段，识别出每个独立的维修项目条目。
    不依赖固定格式。
    """
    _log = log_callback or (lambda msg: None)

    # 文本太长时分块处理
    chunks = _split_text_chunks(raw_text, max_chars=8000)
    all_items = []

    _log(f"文本已分为 {len(chunks)} 个块，开始AI并发解析...")

    client = OpenAI(
        api_key=config.LLM_API_KEY,
        base_url=config.LLM_BASE_URL,
    )
    total_chunks = len(chunks)

    # 并发调用LLM解析每个块
    chunk_results = [None] * total_chunks

    def _parse_one(i, chunk):
        items = _segment_chunk(client, chunk, i + 1, total_chunks)
        return i, items

    with ThreadPoolExecutor(max_workers=min(6, total_chunks)) as pool:
        futures = {pool.submit(_parse_one, i, c): i for i, c in enumerate(chunks)}
        done_count = 0
        for future in as_completed(futures):
            i, items = future.result()
            chunk_results[i] = items
            done_count += 1
            _log(f"  块 {i+1} 完成: {len(items)} 条 ({done_count}/{total_chunks})")

    all_items = []
    for items in chunk_results:
        if items:
            all_items.extend(items)

    # 去重（同一个条目可能在分块边界重复出现）
    before_dedup = len(all_items)
    all_items = _deduplicate(all_items)
    _log(f"去重: {before_dedup} → {len(all_items)} 条")

    return all_items


def _segment_chunk(client: OpenAI, text_chunk: str, chunk_num: int, total_chunks: int) -> list[dict]:
    """对单个文本块调用LLM做分段"""

    prompt = f"""你是船舶维修行业专家。请从以下文本中识别出所有独立的维修/服务项目条目。

## 要求
1. 每个维修项目提取为一条记录
2. 如果有SFI编码就提取，没有就填null
3. 提取标题、描述、数量、单位
4. 忽略目录、页眉页脚、封面等非项目内容
5. 描述要保留关键技术细节，但不超过300字符

## 文本内容（第{chunk_num}/{total_chunks}块）
{text_chunk}

## 输出格式
返回JSON数组，每项格式如下：
```json
[
  {{
    "title": "项目标题（简洁，英文）",
    "description": "详细描述（保留技术要求）",
    "sfi_code": "1.EH.1.1" 或 null,
    "quantity": 数字 或 null,
    "unit": "PC/SET/DAY/HR/LOT等" 或 ""
  }}
]
```
只返回JSON数组，不要其他文字。如果这段文本中没有维修项目，返回空数组 []。"""

    try:
        response = client.chat.completions.create(
            model=config.LLM_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=8000,
        )

        content = response.choices[0].message.content
        json_str = _extract_json_array(content)
        if json_str:
            items = _safe_parse_json_array(json_str)
            # 验证格式
            valid_items = []
            for item in items:
                if isinstance(item, dict) and item.get("title"):
                    valid_items.append({
                        "title": str(item.get("title", "")),
                        "description": str(item.get("description", "")),
                        "sfi_code": item.get("sfi_code"),
                        "quantity": item.get("quantity"),
                        "unit": str(item.get("unit", "")),
                    })
            return valid_items

    except Exception as e:
        print(f"[文档解析] 第{chunk_num}块AI分段失败: {e}")

    return []


def _split_text_chunks(text: str, max_chars: int = 8000) -> list[str]:
    """
    将长文本切分为合适大小的块。
    尽量在段落边界切分，避免切断一个条目。
    """
    if len(text) <= max_chars:
        return [text]

    chunks = []
    paragraphs = text.split("\n\n")
    current_chunk = ""

    for para in paragraphs:
        if len(current_chunk) + len(para) + 2 > max_chars:
            if current_chunk:
                chunks.append(current_chunk)
            current_chunk = para
        else:
            current_chunk += "\n\n" + para if current_chunk else para

    if current_chunk:
        chunks.append(current_chunk)

    # 如果单个段落就超长，强制切分
    final_chunks = []
    for chunk in chunks:
        if len(chunk) > max_chars:
            for i in range(0, len(chunk), max_chars):
                final_chunks.append(chunk[i:i + max_chars])
        else:
            final_chunks.append(chunk)

    return final_chunks


def _extract_json_array(text: str) -> str | None:
    """从LLM返回文本中提取JSON数组"""
    import re
    # 尝试 ```json ... ```
    match = re.search(r'```json\s*([\s\S]*?)\s*```', text)
    if match:
        return match.group(1)
    # 尝试找 [ ... ]
    match = re.search(r'\[[\s\S]*\]', text)
    if match:
        return match.group(0)
    # 可能被截断，只有 [ 开头没有 ] 结尾
    match = re.search(r'\[[\s\S]*', text)
    if match:
        return match.group(0)
    return None


def _safe_parse_json_array(json_str: str) -> list:
    """解析JSON数组，支持截断修复"""
    # 先尝试直接解析
    try:
        result = json.loads(json_str)
        if isinstance(result, list):
            return result
    except json.JSONDecodeError:
        pass

    # 截断修复：逐步删减尾部，尝试闭合
    # 找到最后一个完整的 }, 然后补 ]
    for i in range(len(json_str) - 1, 0, -1):
        if json_str[i] == '}':
            attempt = json_str[:i+1] + ']'
            try:
                result = json.loads(attempt)
                if isinstance(result, list):
                    return result
            except json.JSONDecodeError:
                continue

    return []


def _deduplicate(items: list[dict]) -> list[dict]:
    """去重：基于标题相似度去重"""
    seen_titles = set()
    unique_items = []

    for item in items:
        # 简单去重：标题前30字符相同视为重复
        key = item["title"][:30].upper().strip()
        if key not in seen_titles:
            seen_titles.add(key)
            unique_items.append(item)

    return unique_items
