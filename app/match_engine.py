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
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config
from app.craft_library import search_similar, batch_search_similar
from app.llm_client import get_llm_client

_audit_lock = threading.Lock()


def _append_match_llm_audit(record: dict) -> None:
    """若配置了 MATCH_LLM_AUDIT_JSONL，则追加一行 JSON，供事后分析门控与大模型调用。"""
    path = getattr(config, "MATCH_LLM_AUDIT_JSONL", "") or ""
    if not path:
        return
    try:
        line = json.dumps(record, ensure_ascii=False) + "\n"
        abs_path = os.path.abspath(path)
        d = os.path.dirname(abs_path)
        if d:
            os.makedirs(d, exist_ok=True)
        with _audit_lock:
            with open(abs_path, "a", encoding="utf-8") as f:
                f.write(line)
    except OSError:
        pass


def match_single_item(
    enquiry_item: dict,
    index,
    craft_items: list[dict],
    item_serial: int | None = None,
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
        return _handle_no_match(enquiry_item, item_serial=item_serial)

    return _evaluate_candidates(enquiry_item, candidates, item_serial=item_serial)


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
    max_workers = 8
    _log(f"Phase 2/2: LLM并发精排 (并发={max_workers})...")

    results = [None] * total
    done_count = 0
    lock = threading.Lock()
    t_phase2 = time.time()

    def _do_match(i, item, candidates):
        if candidates is not None:
            return i, _match_with_candidates(item, candidates, item_serial=i)
        else:
            return i, match_single_item(item, index, craft_items, item_serial=i)

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
            dp = result.get("decision_path", "")
            lc = result.get("llm_called", False)
            _log(f"  [{cnt}/{total}] {title_short} → {status} | path={dp} llm={lc}")
            if progress_callback:
                progress_callback(cnt, total)

    _log(f"  并发精排完成，耗时 {time.time()-t_phase2:.1f}s")
    return results


def _match_with_candidates(
    enquiry_item: dict, candidates: list[dict], item_serial: int | None = None
) -> dict:
    """使用预获取的候选列表执行匹配（跳过向量检索步骤）"""

    if not candidates:
        return _handle_no_match(enquiry_item, item_serial=item_serial)

    return _evaluate_candidates(enquiry_item, candidates, item_serial=item_serial)


def _evaluate_candidates(
    enquiry_item: dict, candidates: list[dict], item_serial: int | None = None
) -> dict:
    """根据门控策略决定是否调用LLM，并返回最终结果"""
    decision_path = "vector_only"
    decision_reason = "默认向量直出"
    llm_called = False

    exact_idx = _find_exact_sfi_candidate(enquiry_item, candidates)
    if exact_idx is not None:
        exact_vector = candidates[exact_idx]["score"] * 100
        if exact_vector >= config.LLM_RERANK_VECTOR_MIN:
            decision_path = "direct_sfi"
            decision_reason = f"SFI完全一致且向量分{exact_vector:.1f}，直接命中"
        else:
            decision_reason = f"SFI完全一致但向量分{exact_vector:.1f}偏低，继续门控判断"

    llm_results = None
    if decision_path != "direct_sfi":
        should_llm, reason = _should_call_llm_rerank(enquiry_item, candidates)
        if should_llm:
            llm_results = _llm_rerank(enquiry_item, candidates)
            llm_called = True
            decision_path = "llm_rerank"
            decision_reason = reason
        else:
            decision_path = "vector_only"
            decision_reason = reason

    matches = _compute_matches(
        enquiry_item,
        candidates,
        llm_results=llm_results,
        decision_path=decision_path,
        exact_idx=exact_idx,
    )

    best = matches[0] if matches else None
    v1 = candidates[0]["score"] * 100 if candidates else 0.0
    v2 = candidates[1]["score"] * 100 if len(candidates) > 1 else None
    gap = (v1 - v2) if v2 is not None else None
    audit: dict = {
        "ts": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
        "item_index": item_serial,
        "enquiry_sfi": enquiry_item.get("sfi_code"),
        "enquiry_title": (enquiry_item.get("title") or "")[:200],
        "decision_path": decision_path,
        "llm_called": llm_called,
        "gate_reason": decision_reason,
        "vector_top1": round(v1, 2),
        "vector_top2": round(v2, 2) if v2 is not None else None,
        "vector_gap": round(gap, 2) if gap is not None else None,
        "k_candidates": len(candidates),
    }
    if best:
        audit["best_craft_sfi"] = best.get("craft_sfi")
        audit["best_confidence"] = best.get("confidence")
        audit["best_vector_score"] = best.get("vector_score")
        audit["best_llm_score"] = best.get("llm_score")
        audit["best_sfi_score"] = best.get("sfi_score")
    if llm_results:
        scores = [float(x.get("score", 0) or 0) for x in llm_results]
        audit["llm_score_min"] = round(min(scores), 2)
        audit["llm_score_max"] = round(max(scores), 2)
        fail_mark = "LLM调用失败"
        audit["llm_all_failed"] = all(
            (x.get("reason") or "").startswith(fail_mark) for x in llm_results
        )
    _append_match_llm_audit(audit)

    return _build_result(
        enquiry_item,
        matches,
        decision_path=decision_path,
        decision_reason=decision_reason,
        llm_called=llm_called,
    )


def _compute_matches(enquiry_item, candidates, llm_results=None, decision_path: str = "vector_only", exact_idx: int | None = None):
    """对所有候选计算置信度 + 应用降级规则，返回排序后的matches列表"""
    matches = []
    has_sfi = bool(enquiry_item.get("sfi_code"))

    for idx, candidate in enumerate(candidates):
        llm_result = llm_results[idx] if llm_results and idx < len(llm_results) else {}
        craft = candidate["item"]
        vector_score = candidate["score"] * 100
        llm_score = llm_result.get("score", vector_score)
        sfi_score = _calc_sfi_match_score(
            enquiry_item.get("sfi_code"), craft.get("sfi_code", "")
        )

        # 权重计算
        if has_sfi:
            w1, w2, w3 = config.CONFIDENCE_WEIGHTS["vector_similarity"], \
                         config.CONFIDENCE_WEIGHTS["llm_score"], \
                         config.CONFIDENCE_WEIGHTS["sfi_match"]
        else:
            total_w = config.CONFIDENCE_WEIGHTS["vector_similarity"] + config.CONFIDENCE_WEIGHTS["llm_score"]
            w1 = config.CONFIDENCE_WEIGHTS["vector_similarity"] / total_w
            w2 = config.CONFIDENCE_WEIGHTS["llm_score"] / total_w
            w3 = 0

        confidence = w1 * vector_score + w2 * llm_score + w3 * sfi_score

        if decision_path == "direct_sfi" and exact_idx is not None and idx == exact_idx:
            confidence = max(confidence, 90)

        # === M4: 强制降级规则 ===
        # 规则1: 有SFI但完全不匹配 → 上限75
        if has_sfi and sfi_score == 0:
            confidence = min(confidence, 75)

        # 规则2: 向量分低于60 → 上限70
        if vector_score < 60:
            confidence = min(confidence, 70)

        # 规则3: AI分和向量分差异>25 → 取均值
        if abs(llm_score - vector_score) > 25:
            confidence = (llm_score + vector_score) / 2

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
            "llm_reason": llm_result.get("reason", "规则/向量直出"),
        })

    matches.sort(key=lambda x: x["confidence"], reverse=True)
    return matches


def _build_result(enquiry_item, matches, decision_path: str, decision_reason: str, llm_called: bool):
    """根据matches列表组装最终返回结果"""
    best = matches[0] if matches else None
    top_confidence = best["confidence"] if best else 0
    quality_flags = enquiry_item.get("quality_flags") or []
    if not isinstance(quality_flags, list):
        quality_flags = []

    # 解析阶段发现层级不完整时，强制降级到人工确认
    hierarchy_risk = any(
        f in quality_flags
        for f in ("parent_without_children", "possible_missing_children", "missing_sfi_hierarchy_uncertain")
    )
    if hierarchy_risk and matches:
        matches[0]["confidence"] = min(matches[0]["confidence"], 59)
        reason = matches[0].get("llm_reason", "")
        risk_note = "本条解析层级可能不完整，建议核对子项后再采用匹配结果。"
        matches[0]["llm_reason"] = f"{reason} {risk_note}".strip() if reason else risk_note
        top_confidence = matches[0]["confidence"]

    # === M3: Top-1/2 分差降级 ===
    if len(matches) >= 2:
        gap = matches[0]["confidence"] - matches[1]["confidence"]
        if gap < 5:
            matches[0]["confidence"] = min(matches[0]["confidence"], 75)
            base = (matches[0].get("llm_reason") or "").rstrip()
            extra = "另有工艺与当前首选接近，请核对哪一条最符合客户描述。"
            matches[0]["llm_reason"] = f"{base} {extra}".strip() if base else extra

    # 上述规则可能改变Top-1分值，需刷新
    top_confidence = matches[0]["confidence"] if matches else 0

    # 召回优先：有候选时永远保留 Top1 供人工核对；仅“无候选”视为新增项
    has_candidates = bool(matches)
    no_vector_hit = not has_candidates
    low_confidence = has_candidates and top_confidence < config.CONFIDENCE_LEVELS["low"]
    needs_human_review = (
        low_confidence
        or hierarchy_risk
        or (len(matches) >= 2 and matches[0]["confidence"] - matches[1]["confidence"] < 5)
    )

    if no_vector_hit:
        review_status = "PENDING_REVIEW"
        is_new_item = True
        best_for_display = None
        display_confidence = 0
    else:
        is_new_item = False
        best_for_display = matches[0]
        display_confidence = top_confidence
        review_status = "PENDING_REVIEW" if needs_human_review else "OK"

    result = {
        "enquiry_sfi": enquiry_item.get("sfi_code"),
        "enquiry_title": enquiry_item.get("title", ""),
        "enquiry_description": enquiry_item.get("description", "")[:800],
        "source_chunk": enquiry_item.get("source_chunk", ""),
        "parse_pass": enquiry_item.get("parse_pass", ""),
        "quality_flags": quality_flags,
        "quantity": enquiry_item.get("quantity"),
        "unit": enquiry_item.get("unit", ""),
        "matches": matches[:5],
        "best_match": best_for_display,
        "confidence": display_confidence,
        "is_new_item": is_new_item,
        "needs_human_review": bool(needs_human_review or no_vector_hit),
        "review_status": review_status,
        "must_keep": True,
        "suggested_entry": None,
        "decision_path": decision_path,
        "decision_reason": decision_reason,
        "llm_called": llm_called,
    }

    if is_new_item and config.ENABLE_AUTO_SUGGEST_NEW_ENTRY:
        result["suggested_entry"] = _generate_suggested_entry(enquiry_item)

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
    client = get_llm_client()

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

    prompt = f"""你是船舶维修工艺匹配专家。请严格判断客户询价项目与工艺库候选条目的匹配程度。

## 客户询价项目
- SFI编码: {enquiry_item.get('sfi_code') or '无'}
- 标题: {enquiry_item.get('title', '无')}
- 描述: {enquiry_desc}

## 工艺库候选条目
{candidates_text}

## 严格评分标准
- 90-100: SFI编码相同 + 标题高度一致（两个条件必须同时满足）
- 70-89: SFI编码前2段相同 或 标题关键词高度重叠
- 40-69: 同一大类但具体工作不同
- 0-39: 不相关或仅有表面词汇重叠

## 重要约束
- SFI编码不同但标题相似 → 不超过70分
- 仅凭个别词汇相似 → 不超过50分
- 不确定时宁愿打低分，不要虚高

## reason 字段（写给报价业务员看，与内部分数标准无关）
- 用**一句中文**说明为何给该候选打此分、或为何不确信（约 15～50 字，勿写长段落）
- **禁止**在 reason 中出现：向量、阈值、Top1、Top2、分差、分段、前两段、模型、embedding 等字样
- **不要**逐位对比 SFI 数字；可用「编码大类不同」「工作内容不一致」「字面相近但业务范围不同」等概括说法

如果所有候选都不匹配（都低于40分），请如实打低分。

返回JSON数组：
```json
[{{"index": 1, "score": 85, "reason": "客户询价与候选工艺工作内容一致，可直接对照报价。"}}]
```
只返回JSON。"""

    last_error = None
    for attempt in range(3):
        try:
            response = client.chat.completions.create(
                model=config.LLM_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.05,
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
            last_error = e
            if attempt < 2:
                wait = (2 ** attempt) * 0.5
                time.sleep(wait)

    print(f"[LLM精排] 调用失败(重试3次): {last_error}")
    return [{"score": 0, "reason": "LLM调用失败"} for _ in candidates]


def _generate_suggested_entry(enquiry_item: dict) -> dict:
    """
    当工艺库无匹配时，用LLM基于行业知识生成建议条目。
    人工确认后可回写到工艺库。
    """
    client = get_llm_client()

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


def _handle_no_match(enquiry_item: dict, item_serial: int | None = None) -> dict:
    """向量检索无结果时的处理"""
    _append_match_llm_audit({
        "ts": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
        "item_index": item_serial,
        "enquiry_sfi": enquiry_item.get("sfi_code"),
        "enquiry_title": (enquiry_item.get("title") or "")[:200],
        "decision_path": "no_candidate",
        "llm_called": False,
        "gate_reason": "向量检索无候选",
        "vector_top1": None,
        "vector_top2": None,
        "vector_gap": None,
        "k_candidates": 0,
    })
    suggested = _generate_suggested_entry(enquiry_item) if config.ENABLE_AUTO_SUGGEST_NEW_ENTRY else None
    return {
        "enquiry_sfi": enquiry_item.get("sfi_code"),
        "enquiry_title": enquiry_item.get("title", ""),
        "enquiry_description": enquiry_item.get("description", "")[:800],
        "source_chunk": enquiry_item.get("source_chunk", ""),
        "parse_pass": enquiry_item.get("parse_pass", ""),
        "quality_flags": enquiry_item.get("quality_flags") or [],
        "quantity": enquiry_item.get("quantity"),
        "unit": enquiry_item.get("unit", ""),
        "matches": [],
        "best_match": None,
        "confidence": 0,
        "is_new_item": True,
        "needs_human_review": True,
        "review_status": "PENDING_REVIEW",
        "must_keep": True,
        "suggested_entry": suggested,
        "decision_path": "no_candidate",
        "decision_reason": "向量检索无候选",
        "llm_called": False,
    }


def _find_exact_sfi_candidate(enquiry_item: dict, candidates: list[dict]) -> int | None:
    """在候选中查找SFI完全一致的条目下标"""
    enquiry_sfi = (enquiry_item.get("sfi_code") or "").strip().upper()
    if not enquiry_sfi:
        return None
    for idx, c in enumerate(candidates):
        craft_sfi = (c.get("item", {}).get("sfi_code", "") or "").strip().upper()
        if craft_sfi and craft_sfi == enquiry_sfi:
            return idx
    return None


def _should_call_llm_rerank(enquiry_item: dict, candidates: list[dict]) -> tuple[bool, str]:
    """仅在不确定场景触发LLM精排"""
    if not config.ENABLE_LLM_RERANK:
        return False, "已关闭LLM精排开关"
    if not candidates:
        return False, "无候选"

    top1 = candidates[0]["score"] * 100
    top2 = candidates[1]["score"] * 100 if len(candidates) > 1 else 0
    gap = top1 - top2

    if top1 < config.LLM_RERANK_TOP1_MIN:
        return True, f"Top1向量分{top1:.1f}低于阈值{config.LLM_RERANK_TOP1_MIN}"
    if len(candidates) > 1 and gap < config.LLM_RERANK_GAP_MAX:
        return True, f"Top1与Top2分差{gap:.1f}小于阈值{config.LLM_RERANK_GAP_MAX}"

    enquiry_sfi = (enquiry_item.get("sfi_code") or "").strip()
    top1_sfi = (candidates[0].get("item", {}).get("sfi_code") or "").strip()
    if enquiry_sfi and top1_sfi and _calc_sfi_match_score(enquiry_sfi, top1_sfi) == 0:
        return True, "询价SFI与Top1候选SFI冲突，需要LLM复核"

    return False, f"Top1向量分{top1:.1f}且分差{gap:.1f}，向量直出"
