"""
工艺库加载、向量化、检索、增量更新模块
"""
import json
import os
import re
import numpy as np
from openpyxl import load_workbook
from openai import OpenAI

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config


def load_craft_library(excel_path: str = None) -> list[dict]:
    """
    从万邦报价单模板Excel中提取工艺库条目。

    返回:
    [
        {
            "id": 0,
            "sfi_code": "1.EH.1.1",
            "title": "CRANE SERVICE / SHORE CRANE HOURLY USAGE",
            "detail": "SHORE CRANE SERVICE",
            "unit": "LIFT",
            "qty_template": None,
            "category": "1",
            "full_text": "1.EH.1.1 CRANE SERVICE / SHORE CRANE HOURLY USAGE - SHORE CRANE SERVICE",
        },
        ...
    ]
    """
    path = excel_path or config.CRAFT_LIBRARY_PATH
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb["Quotation"]

    items = []
    current_title_row = None

    for row in ws.iter_rows(min_row=2, values_only=False):
        a_val = str(row[0].value).strip() if row[0].value else ""
        c_val = str(row[2].value).strip() if row[2].value else ""
        d_val = str(row[3].value).strip() if row[3].value else ""
        f_val = row[5].value if row[5].value else None

        if not c_val:
            continue

        # 判断行类型
        is_sfi_row = bool(re.match(r'^\d{1,2}\.[A-Z]{1,4}\.\d', a_val))
        is_category = bool(re.match(r'^\d{1,2}$', a_val))

        if is_category:
            continue

        if is_sfi_row:
            current_title_row = {
                "sfi_code": a_val,
                "title": c_val,
                "detail": "",
                "unit": d_val,
                "qty_template": f_val,
                "category": a_val.split(".")[0],
            }
            items.append(current_title_row)

        elif current_title_row and not is_category:
            # 明细行：补充描述
            if current_title_row["detail"]:
                current_title_row["detail"] += " | " + c_val
            else:
                current_title_row["detail"] = c_val

            if d_val and not current_title_row["unit"]:
                current_title_row["unit"] = d_val
            if f_val and not current_title_row["qty_template"]:
                current_title_row["qty_template"] = f_val

    wb.close()

    # 生成完整文本（用于embedding）和ID
    for i, item in enumerate(items):
        item["id"] = i
        parts = [item["sfi_code"], item["title"]]
        if item["detail"]:
            parts.append(item["detail"])
        item["full_text"] = " - ".join(parts)

    return items


def get_embeddings(texts: list[str], batch_size: int = 6) -> np.ndarray:
    """
    调用Embedding API获取文本向量。
    支持OpenAI兼容接口。
    """
    client = OpenAI(
        api_key=config.EMBED_API_KEY,
        base_url=config.EMBED_BASE_URL,
    )

    all_embeddings = []
    for i in range(0, len(texts), batch_size):
        batch = texts[i:i + batch_size]
        response = client.embeddings.create(
            model=config.EMBED_MODEL,
            input=batch,
        )
        batch_embeddings = [item.embedding for item in response.data]
        all_embeddings.extend(batch_embeddings)

    return np.array(all_embeddings, dtype=np.float32)


def build_vector_index(craft_items: list[dict], save_dir: str = "data"):
    """
    为工艺库构建FAISS向量索引。
    """
    import faiss

    os.makedirs(save_dir, exist_ok=True)

    texts = [item["full_text"] for item in craft_items]

    print(f"[向量化] 正在为 {len(texts)} 条工艺生成embedding...")
    embeddings = get_embeddings(texts)
    print(f"[向量化] 完成，维度: {embeddings.shape}")

    # 归一化（内积 = 余弦相似度）
    faiss.normalize_L2(embeddings)

    # 构建索引
    dim = embeddings.shape[1]
    index = faiss.IndexFlatIP(dim)
    index.add(embeddings)

    # 保存
    index_path = os.path.join(save_dir, "craft_library.index")
    data_path = os.path.join(save_dir, "craft_library.json")

    faiss.write_index(index, index_path)
    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(craft_items, f, ensure_ascii=False, indent=2)

    print(f"[向量化] 索引已保存: {index_path}")
    print(f"[向量化] 数据已保存: {data_path}")

    return index


def load_vector_index(data_dir: str = "data"):
    """加载已保存的向量索引和工艺库数据"""
    import faiss

    index_path = os.path.join(data_dir, "craft_library.index")
    data_path = os.path.join(data_dir, "craft_library.json")

    if not os.path.exists(index_path) or not os.path.exists(data_path):
        return None, None

    index = faiss.read_index(index_path)
    with open(data_path, "r", encoding="utf-8") as f:
        craft_items = json.load(f)

    return index, craft_items


def search_similar(query_text: str, index, craft_items: list[dict], top_k: int = None) -> list[dict]:
    """
    向量检索：找到与query最相似的top-K条工艺。
    """
    import faiss

    k = top_k or config.TOP_K

    query_emb = get_embeddings([query_text])
    faiss.normalize_L2(query_emb)

    scores, indices = index.search(query_emb, k)

    results = []
    for score, idx in zip(scores[0], indices[0]):
        if idx < 0:
            continue
        results.append({
            "item": craft_items[idx],
            "score": float(score),
        })

    return results


def batch_search_similar(query_texts: list[str], index, craft_items: list[dict], top_k: int = None) -> list[list[dict]]:
    """
    批量向量检索：一次API调用获取所有embedding，再批量FAISS检索。
    比逐条调用 search_similar 快 N 倍。
    """
    import faiss

    k = top_k or config.TOP_K
    query_embs = get_embeddings(query_texts)
    faiss.normalize_L2(query_embs)

    scores_batch, indices_batch = index.search(query_embs, k)

    all_results = []
    for scores, indices in zip(scores_batch, indices_batch):
        results = []
        for score, idx in zip(scores, indices):
            if idx < 0:
                continue
            results.append({
                "item": craft_items[idx],
                "score": float(score),
            })
        all_results.append(results)

    return all_results


def add_to_library(new_entry: dict, data_dir: str = "data") -> bool:
    """
    增量更新：将人工确认的新条目加入工艺库。
    （学习进化机制的核心）

    Args:
        new_entry: {
            "title": "...",
            "description": "...",
            "unit": "...",
            "sfi_code": "..." 或 None,
        }
        data_dir: 数据目录

    Returns:
        是否成功
    """
    import faiss

    index_path = os.path.join(data_dir, "craft_library.index")
    data_path = os.path.join(data_dir, "craft_library.json")

    if not os.path.exists(index_path) or not os.path.exists(data_path):
        return False

    # 加载现有数据
    index = faiss.read_index(index_path)
    with open(data_path, "r", encoding="utf-8") as f:
        craft_items = json.load(f)

    # 构建新条目
    new_id = len(craft_items)
    full_text = f"{new_entry.get('sfi_code', '')} {new_entry['title']} - {new_entry.get('description', '')}".strip()

    new_item = {
        "id": new_id,
        "sfi_code": new_entry.get("sfi_code", ""),
        "title": new_entry["title"],
        "detail": new_entry.get("description", ""),
        "unit": new_entry.get("unit", "LOT"),
        "qty_template": None,
        "category": "",
        "full_text": full_text,
        "source": "user_added",  # 标记来源
    }

    # 向量化新条目
    new_emb = get_embeddings([full_text])
    faiss.normalize_L2(new_emb)

    # 加入索引
    index.add(new_emb)
    craft_items.append(new_item)

    # 保存
    faiss.write_index(index, index_path)
    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(craft_items, f, ensure_ascii=False, indent=2)

    print(f"[工艺库更新] 新增条目 #{new_id}: {new_entry['title']}")
    return True
