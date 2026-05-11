"""
AI匹配引擎
- Stage 1: 向量粗检索 top-K 候选
- Stage 2: LLM精排打分
- Stage 3: 置信度计算 + SFI校验 + 无匹配项AI生成描述
"""
import os
import sys
import json
import re
from openai import OpenAI

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config
from app.craft_library import search_similar, batch_search_similar


def match_single_item(
    enquiry_item: dict,
    index,
    craft_items: list[dict],
) -> dict:
    """
    对单个询价条目执行完整匹配流程。

    Returns:
        {
            "enquiry_sfi": "1.EH.1.1" 或 None,
            "enquiry_title": "Crane Service...",
            "enquiry_description": "...",
            "quantity": 3 或 None,
            "unit": "HR",
            "matches": [...],       # top-5候选
            "best_match": {...},    # 最佳匹配
            "confidence": 88,
            "is_new_item": False,   # 是否为工艺库不存在的新项
            "suggested_entry": None, # 新项时AI生成的建议条目
        }
    """
    # 构建查询文本
    query = _build_query_text(enquiry_item)

    # === Stage 1: 向量粗检索 ===
    candidates = search_similar(query, index, craft_items, top_k=config.TOP_K)

    if not candidates:
        return _handle_no_match(enquiry_item)

    # === Stage 2: LLM精排 ===
    llm_results = _llm_rerank(enquiry_item, candidates)

    # === Stage 3: 置信度计算 ===
    matches = []
    for candidate, llm_result in zip(candidates, llm_results):
        craft = candidate["item"]
        vector_score = candidate["score"] * 100  # 归一化到0-100

        llm_score = llm_result.get("score", 0)
        sfi_score = _calc_sfi_match_score(
            enquiry_item.get("sfi_code"),
            craft.get("sfi_code", "")
        )

        # 加权计算置信度
        # 如果询价单没有SFI码，W3权重归零，W1和W2等比放大
        has_sfi = bool(enquiry_item.get("sfi_code"))
        if has_sfi:
            w1 = config.CONFIDENCE_WEIGHTS["vector_similarity"]
            w2 = config.CONFIDENCE_WEIGHTS["llm_score"]
            w3 = config.CONFIDENCE_WEIGHTS["sfi_match"]
        else:
            total_w = config.CONFIDENCE_WEIGHTS["vector_similarity"] + config.CONFIDENCE_WEIGHTS["llm_score"]
            w1 = config.CONFIDENCE_WEIGHTS["vector_similarity"] / total_w
            w2 = config.CONFIDENCE_WEIGHTS["llm_score"] / total_w
            w3 = 0

        confidence = w1 * vector_score + w2 * llm_score + w3 * sfi_score
        confidence = min(100, max(0, round(confidence)))

        matches.append({
            "craft_sfi": craft.get("sfi_code", ""),
            "craft_title": craft.get("title", ""),
            "craft_detail": craft.get("detail", ""),
            "craft_id": craft.get("id"),
            "unit": craft.get("unit", ""),
            "vector_score": round(vector_score, 1),
            "llm_score": llm_score,
            "sfi_score": sfi_score,
            "confidence": confidence,
            "llm_reason": llm_result.get("reason", ""),
        })

    # 按置信度排序
    matches.sort(key=lambda x: x["confidence"], reverse=True)
    best = matches[0] if matches else None

    # 判断是否为"无匹配"（最高置信度 < 40）
    top_confidence = best["confidence"] if best else 0
    is_new_item = top_confidence < config.CONFIDENCE_LEVELS["low"]

    result = {
        "enquiry_sfi": enquiry_item.get("sfi_code"),
        "enquiry_title": enquiry_item.get("title", ""),
        "enquiry_description": enquiry_item.get("description", "")[:300],
        "quantity": enquiry_item.get("quantity"),
        "unit": enquiry_item.get("unit", ""),
        "matches": matches[:5],
        "best_match": best if not is_new_item else None,
        "confidence": top_confidence if not is_new_item else 0,
        "is_new_item": is_new_item,
        "suggested_entry": None,
    }

    # 无匹配时，AI生成建议条目
    if is_new_item:
        result["suggested_entry"] = _generate_suggested_entry(enquiry_item)
        result["confidence"] = 0

    return result


def match_all_items(
    enquiry_items: list[dict],
    index,
    craft_items: list[dict],
    progress_callback=None,
    log_callback=None,
) -> list[dict]:
    """
    对所有询价条目执行匹配（批量优化版）。

    优化点：先批量获取所有embedding做向量检索，再逐条LLM精排。
    原来N条要N次embedding API调用，现在只需1次（或几次批量调用）。
    """
    import time

    _log = log_callback or (lambda msg: None)
    total = len(enquiry_items)

    # === Phase 1: 批量向量检索 ===
    _log(f"Phase 1/2: 批量向量检索 {total} 条...")
    t0 = time.time()

    query_texts = [_build_query_text(item) for item in enquiry_items]
    try:
        all_candidates = batch_search_similar(query_texts, index, craft_items, top_k=config.TOP_K)
        _log(f"  向量检索完成，耗时 {time.time()-t0:.1f}s")
    except Exception as e:
        _log(f"  ❌ 批量向量检索失败: {e}，退回逐条模式")
        all_candidates = None

    # === Phase 2: 并发LLM精排 ===
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import threading

    max_workers = 8
    _log(f"Phase 2/2: LLM并发精排 (并发={max_workers})...")

    results = [None] * total
    done_count = 0
    lock = threading.Lock()
    t_phase2 = time.time()

    def _do_match(i, item, candidates):
        if candidates is not None:
            return i, _match_with_candidates(item, candidates)
        else:
            return i, match_single_item(item, index, craft_items)

    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {}
        for i, item in enumerate(enquiry_items):
            cands = all_candidates[i] if all_candidates is not None else None
            fut = pool.submit(_do_match, i, item, cands)
            futures[fut] = i

        for fut in as_completed(futures):
            i, result = fut.result()
            results[i] = result
            with lock:
                done_count += 1
                cnt = done_count
            conf = result.get("confidence", 0)
            status = "新增" if result.get("is_new_item") else f"置信度={conf}"
            title_short = enquiry_items[i].get("title", "")[:40]
            _log(f"  [{cnt}/{total}] {title_short} → {status}")
            if progress_callback:
                progress_callback(cnt, total)

    _log(f"  并发精排完成，耗时 {time.time()-t_phase2:.1f}s")
    return results


def _match_with_candidates(enquiry_item: dict, candidates: list[dict]) -> dict:
    """使用预获取的候选列表执行匹配（跳过向量检索步骤）"""

    if not candidates:
        return _handle_no_match(enquiry_item)

    # LLM精排
    llm_results = _llm_rerank(enquiry_item, candidates)

    # 置信度计算
    matches = []
    for candidate, llm_result in zip(candidates, llm_results):
        craft = candidate["item"]
        vector_score = candidate["score"] * 100

        llm_score = llm_result.get("score", 0)
        sfi_score = _calc_sfi_match_score(
            enquiry_item.get("sfi_code"),
            craft.get("sfi_code", "")
        )

        has_sfi = bool(enquiry_item.get("sfi_code"))
        if has_sfi:
            w1 = config.CONFIDENCE_WEIGHTS["vector_similarity"]
            w2 = config.CONFIDENCE_WEIGHTS["llm_score"]
            w3 = config.CONFIDENCE_WEIGHTS["sfi_match"]
        else:
            total_w = config.CONFIDENCE_WEIGHTS["vector_similarity"] + config.CONFIDENCE_WEIGHTS["llm_score"]
            w1 = config.CONFIDENCE_WEIGHTS["vector_similarity"] / total_w
            w2 = config.CONFIDENCE_WEIGHTS["llm_score"] / total_w
            w3 = 0

        confidence = w1 * vector_score + w2 * llm_score + w3 * sfi_score
        confidence = min(100, max(0, round(confidence)))

        matches.append({
            "craft_sfi": craft.get("sfi_code", ""),
            "craft_title": craft.get("title", ""),
            "craft_detail": craft.get("detail", ""),
            "craft_id": craft.get("id"),
            "unit": craft.get("unit", ""),
            "vector_score": round(vector_score, 1),
            "llm_score": llm_score,
            "sfi_score": sfi_score,
            "confidence": confidence,
            "llm_reason": llm_result.get("reason", ""),
        })

    matches.sort(key=lambda x: x["confidence"], reverse=True)
    best = matches[0] if matches else None

    top_confidence = best["confidence"] if best else 0
    is_new_item = top_confidence < config.CONFIDENCE_LEVELS["low"]

    result = {
        "enquiry_sfi": enquiry_item.get("sfi_code"),
        "enquiry_title": enquiry_item.get("title", ""),
        "enquiry_description": enquiry_item.get("description", "")[:300],
        "quantity": enquiry_item.get("quantity"),
        "unit": enquiry_item.get("unit", ""),
        "matches": matches[:5],
        "best_match": best if not is_new_item else None,
        "confidence": top_confidence if not is_new_item else 0,
        "is_new_item": is_new_item,
        "suggested_entry": None,
    }

    if is_new_item:
        result["suggested_entry"] = _generate_suggested_entry(enquiry_item)
        result["confidence"] = 0

    return result


def _build_query_text(enquiry_item: dict) -> str:
    """构建用于向量检索的查询文本"""
    parts = []
    if enquiry_item.get("sfi_code"):
        parts.append(enquiry_item["sfi_code"])
    if enquiry_item.get("title"):
        parts.append(enquiry_item["title"])
    if enquiry_item.get("description"):
        # 取描述的前300字符
        parts.append(enquiry_item["description"][:300])
    return " | ".join(parts)


def _llm_rerank(enquiry_item: dict, candidates: list[dict]) -> list[dict]:
    """
    用LLM对候选进行精排打分。
    """
    client = OpenAI(
        api_key=config.LLM_API_KEY,
        base_url=config.LLM_BASE_URL,
    )

    # 构建候选描述
    candidates_text = ""
    for i, c in enumerate(candidates):
        craft = c["item"]
        candidates_text += (
            f"  [{i+1}] SFI: {craft.get('sfi_code', '无')} | "
            f"标题: {craft['title']} | "
            f"详情: {craft.get('detail', '无')}\n"
        )

    enquiry_desc = enquiry_item.get("description", "无")[:500]

    prompt = f"""你是船舶维修工艺匹配专家。判断客户询价项目与哪些工艺库条目匹配。

## 客户询价项目
- SFI编码: {enquiry_item.get('sfi_code') or '无'}
- 标题: {enquiry_item.get('title', '无')}
- 描述: {enquiry_desc}

## 工艺库候选条目
{candidates_text}

## 评分标准
- 90-100: 完全匹配，同一个维修项目
- 70-89: 高度相关，同一工作范畴
- 40-69: 部分相关
- 0-39: 不相关

如果所有候选都不匹配（都低于40分），请如实打低分。

返回JSON数组：
```json
[{{"index": 1, "score": 85, "reason": "简短原因"}}]
```
只返回JSON。"""

    try:
        response = client.chat.completions.create(
            model=config.LLM_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=1500,
        )

        content = response.choices[0].message.content
        json_str = _extract_json(content)
        if json_str:
            scores = json.loads(json_str)
            result = [{"score": 0, "reason": ""} for _ in candidates]
            for s in scores:
                idx = s.get("index", 0) - 1
                if 0 <= idx < len(result):
                    result[idx] = {
                        "score": min(100, max(0, s.get("score", 0))),
                        "reason": s.get("reason", ""),
                    }
            return result

    except Exception as e:
        print(f"[LLM精排] 调用失败: {e}")

    return [{"score": 0, "reason": "LLM调用失败"} for _ in candidates]


def _generate_suggested_entry(enquiry_item: dict) -> dict:
    """
    当工艺库无匹配时，用LLM基于行业知识生成建议条目。
    人工确认后可回写到工艺库。
    """
    client = OpenAI(
        api_key=config.LLM_API_KEY,
        base_url=config.LLM_BASE_URL,
    )

    prompt = f"""你是船舶维修报价专家。以下是一个客户询价项目，但我们的工艺库中没有对应条目。
请根据行业通用标准，生成一个标准化的工艺条目建议。

## 客户询价
- 标题: {enquiry_item.get('title', '')}
- 描述: {enquiry_item.get('description', '')[:500]}
- SFI编码: {enquiry_item.get('sfi_code') or '无'}

## 要求
生成一个标准化工艺条目，用于万邦船舶报价单：
1. title: 标准化英文标题（简洁，行业通用表达，全大写）
2. description: 简要英文描述（一句话说明工作内容）
3. unit: 建议计量单位（PC/SET/DAY/HR/LOT/M2等）
4. category: 所属分类

返回JSON：
```json
{{"title": "...", "description": "...", "unit": "...", "category": "..."}}
```
只返回JSON。"""

    try:
        response = client.chat.completions.create(
            model=config.LLM_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            max_tokens=500,
        )

        content = response.choices[0].message.content
        json_str = _extract_json(content)
        if json_str:
            entry = json.loads(json_str)
            return {
                "title": entry.get("title", enquiry_item.get("title", "")),
                "description": entry.get("description", ""),
                "unit": entry.get("unit", "LOT"),
                "category": entry.get("category", ""),
                "source": "ai_generated",
                "original_enquiry": enquiry_item.get("title", ""),
            }

    except Exception as e:
        print(f"[AI生成条目] 调用失败: {e}")

    # 失败时用客户原始描述
    return {
        "title": enquiry_item.get("title", "UNKNOWN ITEM"),
        "description": enquiry_item.get("description", "")[:200],
        "unit": enquiry_item.get("unit", "LOT"),
        "category": "",
        "source": "original_enquiry",
        "original_enquiry": enquiry_item.get("title", ""),
    }


def _calc_sfi_match_score(enquiry_sfi: str | None, craft_sfi: str) -> float:
    """
    计算SFI编码匹配度（0-100）
    """
    if not enquiry_sfi or not craft_sfi:
        return 0

    if enquiry_sfi == craft_sfi:
        return 100

    e_parts = enquiry_sfi.split(".")
    c_parts = craft_sfi.split(".")

    common = 0
    for ep, cp in zip(e_parts, c_parts):
        if ep.upper() == cp.upper():
            common += 1
        else:
            break

    if common >= 3:
        return 75
    elif common >= 2:
        return 50
    elif common >= 1:
        return 25
    return 0


def _extract_json(text: str) -> str | None:
    """从LLM返回中提取JSON"""
    match = re.search(r'```json\s*([\s\S]*?)\s*```', text)
    if match:
        return match.group(1)
    # 尝试 { ... } 或 [ ... ]
    match = re.search(r'[\[{][\s\S]*[\]}]', text)
    if match:
        return match.group(0)
    return None


def _handle_no_match(enquiry_item: dict) -> dict:
    """向量检索无结果时的处理"""
    suggested = _generate_suggested_entry(enquiry_item)
    return {
        "enquiry_sfi": enquiry_item.get("sfi_code"),
        "enquiry_title": enquiry_item.get("title", ""),
        "enquiry_description": enquiry_item.get("description", "")[:300],
        "quantity": enquiry_item.get("quantity"),
        "unit": enquiry_item.get("unit", ""),
        "matches": [],
        "best_match": None,
        "confidence": 0,
        "is_new_item": True,
        "suggested_entry": suggested,
    }
