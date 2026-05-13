"""
万邦船舶询价系统 — 优化验证测试套件
覆盖 16 项优化的代码级验证（无需 API 调用）

用法：
  python test_optimizations.py           # 运行全部测试
  python test_optimizations.py --verbose # 详细输出
"""
import sys
import os
import json
import tempfile
import math

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# ============================================================
# 测试框架（极简，不依赖 pytest）
# ============================================================
PASS, FAIL = 0, 0
TESTS = []

def test(name):
    """装饰器：注册一个测试用例"""
    def decorator(fn):
        TESTS.append((name, fn))
        return fn
    return decorator

def run_all(verbose=False):
    global PASS, FAIL
    PASS = FAIL = 0
    print("=" * 60)
    print("  万邦船舶询价系统 — 优化验证测试")
    print("=" * 60)
    for name, fn in TESTS:
        try:
            fn()
            PASS += 1
            print(f"  [PASS] {name}")
        except AssertionError as e:
            FAIL += 1
            print(f"  [FAIL] {name}")
            print(f"         {e}")
        except Exception as e:
            FAIL += 1
            print(f"  [ERROR] {name}: {e}")
    print(f"\n{'='*60}")
    print(f"  结果: {PASS} 通过 / {FAIL} 失败 / {PASS+FAIL} 总计")
    print(f"{'='*60}")
    return FAIL == 0


def assert_eq(a, b, msg=""):
    if a != b:
        raise AssertionError(f"期望 {b!r}, 实际 {a!r}" + (f" — {msg}" if msg else ""))

def assert_true(cond, msg=""):
    if not cond:
        raise AssertionError(msg or "条件不成立")

def assert_in(sub, container, msg=""):
    if sub not in container:
        raise AssertionError(f"{sub!r} 不在容器中" + (f" — {msg}" if msg else ""))


# ============================================================
# R1: 基础铺垫 + Bug修复
# ============================================================

@test("R1-A2-1: llm_client 单例模式")
def test_llm_client_singleton():
    from app.llm_client import get_llm_client, get_embed_client
    c1 = get_llm_client()
    c2 = get_llm_client()
    assert_true(c1 is c2, "get_llm_client 应返回同一实例(单例)")

@test("R1-A2-2: llm_client embed 单例模式")
def test_embed_client_singleton():
    from app.llm_client import get_embed_client
    e1 = get_embed_client()
    e2 = get_embed_client()
    assert_true(e1 is e2, "get_embed_client 应返回同一实例(单例)")

@test("R1-A2-3: llm_client timeout=120 已设置")
def test_llm_client_timeout():
    from app.llm_client import get_llm_client
    c = get_llm_client()
    assert_eq(c.timeout, 120, "LLM客户端timeout应为120")

@test("R1-A2-4: llm_client max_retries=2 已设置")
def test_llm_client_max_retries():
    from app.llm_client import get_llm_client
    c = get_llm_client()
    assert_eq(c.max_retries, 2, "LLM客户端max_retries应为2")

@test("R1-A2-5: craft_library 使用 get_embed_client 而非直接 OpenAI")
def test_craft_library_no_direct_openai():
    """验证 craft_library.py 不再直接 import OpenAI"""
    import app.craft_library as cl
    src = open(cl.__file__, encoding="utf-8").read()
    assert_true("from openai import OpenAI" not in src, "craft_library.py 不应再 import OpenAI")

@test("R1-A2-6: document_parser 使用 get_llm_client 而非直接 OpenAI")
def test_document_parser_no_direct_openai():
    import app.document_parser as dp
    src = open(dp.__file__, encoding="utf-8").read()
    assert_true("from openai import OpenAI" not in src, "document_parser.py 不应再 import OpenAI")

@test("R1-A2-7: match_engine 使用 get_llm_client 而非直接 OpenAI")
def test_match_engine_no_direct_openai():
    import app.match_engine as me
    src = open(me.__file__, encoding="utf-8").read()
    assert_true("from openai import OpenAI" not in src, "match_engine.py 不应再 import OpenAI")

@test("R1-B1-1: Excel 统计 low 范围为 40-59")
def test_excel_low_range_fixed():
    """B1: low 统计范围为 40-59"""
    import app.excel_generator as eg
    src = open(eg.__file__, encoding="utf-8").read()
    assert_in('40 <= r.get("confidence", 0) < 60', src, "excel_generator 中 low 统计应为 40-59")

@test("R1-B1-2: no_match 变量已删除")
def test_excel_no_match_removed():
    """B1: 未使用的 no_match 变量应已删除"""
    import app.excel_generator as eg
    src = open(eg.__file__, encoding="utf-8").read()
    assert_true("no_match" not in src, "excel_generator.py 中不应再有 no_match 变量")

@test("R1-A4-1: match_engine imports 在文件顶部")
def test_match_engine_imports_at_top():
    """A4: import time/threading/ThreadPoolExecutor 移到顶部"""
    import app.match_engine as me
    src = open(me.__file__, encoding="utf-8").read()
    # 检查在函数体内没有重复 import
    assert_true('    import time\n' not in src, "match_engine 函数内不应有 import time")
    assert_true('    from concurrent.futures import' not in src, "match_engine 函数内不应有 concurrent.futures import")
    assert_true('    import threading\n' not in src, "match_engine 函数内不应有 import threading")

@test("R1-A4-2: craft_library faiss 在顶部统一导入")
def test_craft_library_faiss_at_top():
    """A4: import faiss 移到文件顶部"""
    import app.craft_library as cl
    src = open(cl.__file__, encoding="utf-8").read()
    # 函数体内不应再有 import faiss
    assert_true('    import faiss\n' not in src, "craft_library 函数内不应有 import faiss")


# ============================================================
# R2: 匹配准确度
# ============================================================

@test("R2-A1-1: _compute_matches 函数存在")
def test_compute_matches_exists():
    from app.match_engine import _compute_matches
    assert_true(callable(_compute_matches), "_compute_matches 应存在且可调用")

@test("R2-A1-2: _build_result 函数存在")
def test_build_result_exists():
    from app.match_engine import _build_result
    assert_true(callable(_build_result), "_build_result 应存在且可调用")

@test("R2-A1-3: match_single_item 使用 _compute_matches + _build_result")
def test_match_single_item_refactored():
    """match_single_item 不再内联置信度计算，而是调用提取的函数"""
    import app.match_engine as me
    src = open(me.__file__, encoding="utf-8").read()
    # 函数体内应该调用 _compute_matches 和 _build_result
    assert_in("_compute_matches(", src, "match_engine 应调用 _compute_matches")
    assert_in("_build_result(", src, "match_engine 应调用 _build_result")

@test("R2-A1-4: _match_with_candidates 使用 _compute_matches + _build_result")
def test_match_with_candidates_refactored():
    import app.match_engine as me
    src = open(me.__file__, encoding="utf-8").read()
    # 确保老代码的重复置信度计算已移除
    assert_true('w1 * vector_score + w2 * llm_score' not in src or src.count('w1 * vector_score + w2 * llm_score') == 1,
                "置信度计算公式应只出现在 _compute_matches 中一次")

@test("R2-A1-5: _compute_matches 返回排序后的 matches 列表")
def test_compute_matches_returns_sorted():
    from app.match_engine import _compute_matches

    enquiry = {"sfi_code": "1.EH.1.1", "title": "CRANE SERVICE", "description": "Test"}
    candidates = [
        {"item": {"sfi_code": "1.EH.1.1", "title": "CRANE A", "detail": "X", "id": 0, "unit": "HR"}, "score": 0.95},
        {"item": {"sfi_code": "2.AB.3.4", "title": "PIPE B", "detail": "Y", "id": 1, "unit": "PC"}, "score": 0.60},
        {"item": {"sfi_code": "1.EH.1.1", "title": "CRANE C", "detail": "Z", "id": 2, "unit": "HR"}, "score": 0.90},
    ]
    llm_results = [
        {"score": 95, "reason": "完全匹配"},
        {"score": 30, "reason": "不相关"},
        {"score": 85, "reason": "高度匹配"},
    ]

    matches = _compute_matches(enquiry, candidates, llm_results)
    assert_eq(len(matches), 3)
    # 应降序排列
    for i in range(len(matches) - 1):
        assert_true(matches[i]["confidence"] >= matches[i+1]["confidence"],
                    f"matches 应按 confidence 降序排列，但 {matches[i]['confidence']} < {matches[i+1]['confidence']}")

@test("R2-A1-6: _build_result 正确组装返回结构")
def test_build_result_structure():
    from app.match_engine import _build_result

    enquiry = {"sfi_code": "1.A.1.1", "title": "TEST", "description": "Desc", "quantity": 3, "unit": "PC"}
    matches = [{
        "craft_sfi": "1.A.1.1", "craft_title": "MATCHED", "craft_detail": "X",
        "craft_id": 0, "unit": "PC", "vector_score": 95.0,
        "llm_score": 90, "sfi_score": 100, "confidence": 92, "llm_reason": "匹配"
    }]

    result = _build_result(enquiry, matches, "vector_only", "单元测试", False)

    assert_eq(result["enquiry_sfi"], "1.A.1.1")
    assert_eq(result["enquiry_title"], "TEST")
    assert_eq(result["quantity"], 3)
    assert_eq(result["unit"], "PC")
    assert_eq(len(result["matches"]), 1)
    assert_true(result["best_match"] is not None)
    assert_true(result["confidence"] > 0)
    assert_true(not result["is_new_item"])
    assert_eq(result.get("review_status"), "OK")

@test("R2-M4-1: SFI 完全不匹配 → 置信度上限 75")
def test_m4_sfi_mismatch_cap():
    from app.match_engine import _compute_matches

    enquiry = {"sfi_code": "1.EH.1.1", "title": "A", "description": "X"}
    candidates = [
        {"item": {"sfi_code": "9.ZZ.9.9", "title": "B", "detail": "Y", "id": 0, "unit": "PC"}, "score": 0.85},
    ]
    llm_results = [{"score": 95, "reason": "fake"}]

    matches = _compute_matches(enquiry, candidates, llm_results)
    assert_true(matches[0]["confidence"] <= 75,
                f"M4规则1: SFI完全不匹配应上限75, 实际 {matches[0]['confidence']}")

@test("R2-M4-2: 向量分 < 60 → 置信度上限 70")
def test_m4_low_vector_cap():
    from app.match_engine import _compute_matches

    enquiry = {"sfi_code": "1.EH.1.1", "title": "A", "description": "X"}
    candidates = [
        {"item": {"sfi_code": "1.EH.1.1", "title": "SAME", "detail": "Y", "id": 0, "unit": "PC"}, "score": 0.55},
    ]
    # vector_score=55, llm_score=65 差=10<25 不会触发规则3取均值
    llm_results = [{"score": 65, "reason": "fake"}]

    matches = _compute_matches(enquiry, candidates, llm_results)
    # 基础分 = 0.40*55 + 0.35*65 + 0.25*100 = 22+22.75+25 = 69.75 → round 70
    # 规则2: vector=55<60, min(70, 70)=70
    assert_true(matches[0]["confidence"] <= 70,
                f"M4规则2: 向量分<60应上限70, 实际 {matches[0]['confidence']}")

@test("R2-M4-3: AI分与向量分差异 > 25 → 取均值")
def test_m4_divergence_mean():
    from app.match_engine import _compute_matches

    enquiry = {"sfi_code": "1.EH.1.1", "title": "A", "description": "X"}
    candidates = [
        {"item": {"sfi_code": "1.EH.1.1", "title": "SAME", "detail": "Y", "id": 0, "unit": "PC"}, "score": 0.90},
    ]
    llm_results = [{"score": 40, "reason": "fake"}]  # 向量90 vs AI 40, 差=50 > 25

    matches = _compute_matches(enquiry, candidates, llm_results)
    expected_mean = round((90 + 40) / 2)  # = 65
    assert_eq(matches[0]["confidence"], expected_mean,
              f"M4规则3: 差异>25应取均值, 期望{expected_mean}, 实际{matches[0]['confidence']}")

@test("R2-M3-1: Top-1/2 分差 < 5 → 置信度上限 75 + 原因标注")
def test_m3_top_gap_downgrade():
    from app.match_engine import _build_result

    enquiry = {"sfi_code": None, "title": "X", "description": ""}
    matches = [
        {"craft_sfi": "", "craft_title": "A", "craft_detail": "", "craft_id": 0, "unit": "",
         "vector_score": 80.0, "llm_score": 80, "sfi_score": 0, "confidence": 80, "llm_reason": "原因1"},
        {"craft_sfi": "", "craft_title": "B", "craft_detail": "", "craft_id": 1, "unit": "",
         "vector_score": 79.0, "llm_score": 79, "sfi_score": 0, "confidence": 79, "llm_reason": "原因2"},
    ]

    result = _build_result(enquiry, matches, "vector_only", "单元测试", False)
    # 分差=1 < 5, 应降级
    assert_true(matches[0]["confidence"] <= 75,
                f"M3: 分差<5应将top-1上限75, 实际 {matches[0]['confidence']}")
    assert_in("另有工艺", matches[0]["llm_reason"],
              "M3: 分差过近时应提示核对相近工艺")

@test("R2-M1-1: Prompt 包含严格评分标准")
def test_m1_strict_prompt():
    """M1: LLM prompt 应包含新的严格评分标准"""
    import app.match_engine as me
    src = open(me.__file__, encoding="utf-8").read()
    assert_in("严格评分标准", src)
    assert_in("SFI编码相同 + 标题高度一致", src)
    assert_in("SFI编码不同但标题相似 → 不超过70分", src)
    assert_in("不确定时宁愿打低分", src)

@test("R2-M1-2: temperature 改为 0.05")
def test_m1_temperature():
    import app.match_engine as me
    src = open(me.__file__, encoding="utf-8").read()
    assert_in("temperature=0.05", src, "LLM精排 temperature 应为 0.05")

@test("R2-M2: 权重已调整为 0.40/0.35/0.25")
def test_m2_weights():
    import config
    w = config.CONFIDENCE_WEIGHTS
    assert_eq(w["vector_similarity"], 0.40, "向量相似度权重应为 0.40")
    assert_eq(w["llm_score"], 0.35, "LLM打分权重应为 0.35")
    assert_eq(w["sfi_match"], 0.25, "SFI匹配权重应为 0.25")


# ============================================================
# R3: 性能 + 工程化
# ============================================================

@test("R3-P1-1: 文档分块使用 config.PARSE_CHUNK_MAX_CHARS")
def test_p1_max_chars():
    import app.document_parser as dp
    src = open(dp.__file__, encoding="utf-8").read()
    assert_in("config.PARSE_CHUNK_MAX_CHARS", src, "分块应使用 config.PARSE_CHUNK_MAX_CHARS")

@test("R3-P1-2: max_workers 改为 min(8, total_chunks)")
def test_p1_max_workers():
    import app.document_parser as dp
    src = open(dp.__file__, encoding="utf-8").read()
    assert_in("min(8, total_chunks)", src, "max_workers 应改为 min(8, total_chunks)")

@test("R3-A3-1: main.py 有 init_session_state 函数")
def test_main_init_session_state():
    import main
    assert_true(callable(main.init_session_state), "main.py 应有 init_session_state 函数")

@test("R3-A3-2: main.py 有 render_sidebar 函数")
def test_main_render_sidebar():
    import main
    assert_true(callable(main.render_sidebar), "main.py 应有 render_sidebar 函数")

@test("R3-A3-3: main.py 有 run_parsing_pipeline 函数")
def test_main_parsing_pipeline():
    import main
    assert_true(callable(main.run_parsing_pipeline), "main.py 应有 run_parsing_pipeline 函数")

@test("R3-A3-4: main.py 有 run_matching_pipeline 函数")
def test_main_matching_pipeline():
    import main
    assert_true(callable(main.run_matching_pipeline), "main.py 应有 run_matching_pipeline 函数")

@test("R3-Q2-1: validator.py 存在 validate_config")
def test_validator_exists():
    from app.validator import validate_config
    assert_true(callable(validate_config), "validator.py 应提供 validate_config 函数")

@test("R3-Q2-2: validate_config 返回 list")
def test_validator_returns_list():
    from app.validator import validate_config
    errors = validate_config()
    assert_true(isinstance(errors, list), "validate_config 应返回 list")

@test("R3-Q2-3: validate_config 检查工艺库路径")
def test_validator_checks_craft_path():
    """用不存在的路径验证校验逻辑"""
    import app.validator as v
    import config
    orig = config.CRAFT_LIBRARY_PATH
    try:
        config.CRAFT_LIBRARY_PATH = "C:/nonexistent/path.xlsx"
        errors = v.validate_config()
        assert_true(any("工艺库" in e for e in errors), "应检测到工艺库路径不存在")
    finally:
        config.CRAFT_LIBRARY_PATH = orig

@test("R3-Q1-1: _llm_rerank 有重试逻辑 (3次尝试)")
def test_q1_retry_logic():
    """Q1: _llm_rerank 应有重试循环"""
    import app.match_engine as me
    src = open(me.__file__, encoding="utf-8").read()
    assert_in("for attempt in range(3)", src, "_llm_rerank 应有3次重试")
    assert_in("time.sleep", src, "_llm_rerank 重试间应有 sleep 退避")


# ============================================================
# R4: 结构性改进
# ============================================================

@test("R4-S4-1: 解析 prompt 包含多行合并指令")
def test_s4_multiline_prompt():
    """S4: prompt 应指导 LLM 用 | 合并多行描述"""
    import app.document_parser as dp
    src = open(dp.__file__, encoding="utf-8").read()
    assert_in("多行用 | 分隔", src, "S4: prompt应包含多行合并 | 指令")
    assert_in("保留所有关键技术细节和多行任务内容", src, "S4: prompt应保留所有技术细节")

@test("R4-S4-2: 结果描述截断改为 800 字符")
def test_s4_description_limit_800():
    """S4: match_engine 中 _build_result/_handle_no_match 的描述截断从 300 改为 800"""
    import app.match_engine as me
    src = open(me.__file__, encoding="utf-8").read()
    assert_in('[:800]', src, "match_engine 结果描述截断应为 800")
    # _build_query_text 保留 300 截断是正常的（用于 embedding 检索），只验证 _build_result 改了
    assert_in('enquiry_item.get("description", "")[:800]', src,
              "_build_result 中 enquiry_description 截断应为 800")

@test("R4-S1-1: 解析 prompt 包含 parent_sfi 字段")
def test_s1_parent_sfi_in_prompt():
    """S1: prompt 应要求 LLM 输出 parent_sfi"""
    import app.document_parser as dp
    src = open(dp.__file__, encoding="utf-8").read()
    assert_in("parent_sfi", src, "S1: prompt应包含 parent_sfi 字段")

@test("R4-S1-2: 解析 prompt 包含 is_range 字段")
def test_s1_is_range_in_prompt():
    """S1: prompt 应要求 LLM 输出 is_range"""
    import app.document_parser as dp
    src = open(dp.__file__, encoding="utf-8").read()
    assert_in("is_range", src, "S1: prompt应包含 is_range 字段")

@test("R4-S1-3: 输出验证包含 parent_sfi 和 is_range")
def test_s1_validation_includes_new_fields():
    """S1: 解析结果验证应包含 parent_sfi 和 is_range"""
    import app.document_parser as dp
    src = open(dp.__file__, encoding="utf-8").read()
    assert_in('"parent_sfi": item.get("parent_sfi")', src, "验证应包含 parent_sfi")
    assert_in('"is_range": item.get("is_range"', src, "验证应包含 is_range")


# ============================================================
# 集成场景测试
# ============================================================

@test("INT-1: 完整匹配流程（无 SFI 场景）")
def test_integration_no_sfi():
    """模拟无 SFI 编码的询价条目匹配流程"""
    from app.match_engine import _compute_matches, _build_result

    enquiry = {"sfi_code": None, "title": "GENERAL CLEANING", "description": "Clean all decks"}
    candidates = [
        {"item": {"sfi_code": "5.CL.1.1", "title": "DECK CLEANING", "detail": "General deck cleaning",
                  "id": 0, "unit": "M2"}, "score": 0.82},
    ]
    llm_results = [{"score": 78, "reason": "同一清洁工作范畴"}]

    matches = _compute_matches(enquiry, candidates, llm_results)
    result = _build_result(enquiry, matches, "vector_only", "单元测试", False)

    assert_true(result["confidence"] > 0)
    # matches 里每条 sfi_score 应为 0（因为无 SFI 编码）
    for m in result["matches"]:
        assert_eq(m["sfi_score"], 0, "无SFI时各条SFI分应都为0")
    # 无SFI时权重归一化：总权重 = 0.40+0.35 = 0.75
    import config
    total_w = config.CONFIDENCE_WEIGHTS["vector_similarity"] + config.CONFIDENCE_WEIGHTS["llm_score"]
    assert_true(abs(total_w - 0.75) < 0.01, f"无SFI时总权重应约0.75, 实际{total_w}")

@test("INT-2: 低置信有候选时保留 best_match 并标记待人工")
def test_integration_new_item_threshold():
    """召回优先：有候选且置信<40 不视为库无候选，应保留 Top1 并 PENDING_REVIEW"""
    from app.match_engine import _build_result

    enquiry = {"sfi_code": None, "title": "X", "description": ""}
    matches = [
        {"craft_sfi": "", "craft_title": "A", "craft_detail": "", "craft_id": 0, "unit": "",
         "vector_score": 30.0, "llm_score": 35, "sfi_score": 0, "confidence": 33, "llm_reason": ""},
    ]

    result = _build_result(enquiry, matches, "vector_only", "单元测试", False)
    assert_true(not result["is_new_item"], "有候选时不应为 is_new_item(库无候选)")
    assert_true(result["best_match"] is not None, "应保留 best_match 供人工核对")
    assert_eq(result["confidence"], 33)
    assert_true(result.get("needs_human_review"), "低置信应 needs_human_review")
    assert_eq(result.get("review_status"), "PENDING_REVIEW")
    assert_true(result.get("must_keep"))

@test("INT-3: 所有匹配字段完整性")
def test_integration_match_fields_complete():
    """验证 match 条目包含所有必要字段"""
    from app.match_engine import _compute_matches

    enquiry = {"sfi_code": "1.A.1.1", "title": "T", "description": "D"}
    candidates = [
        {"item": {"sfi_code": "1.A.1.1", "title": "T", "detail": "D", "id": 42, "unit": "PC"}, "score": 0.95},
    ]
    llm_results = [{"score": 95, "reason": "完全匹配"}]

    matches = _compute_matches(enquiry, candidates, llm_results)
    m = matches[0]

    required_fields = ["craft_sfi", "craft_title", "craft_detail", "craft_id",
                       "unit", "vector_score", "llm_score", "sfi_score",
                       "confidence", "llm_reason"]
    for f in required_fields:
        assert_true(f in m, f"match 缺少字段: {f}")

    assert_eq(m["craft_id"], 42)


@test("LH-1: learn_history 追加与按 file_hash 筛选读取")
def test_learn_history_jsonl():
    from app.learn_history import append_learn_event, read_recent_learn_events

    with tempfile.NamedTemporaryFile(mode="w", suffix=".jsonl", delete=False) as tf:
        p = tf.name
    try:
        append_learn_event({"source": "test", "x": 1}, path=p)
        append_learn_event({"source": "test", "file_hash": "abc", "x": 2}, path=p)
        all_rows = read_recent_learn_events(10, file_hash=None, path=p)
        assert_true(len(all_rows) >= 2, "应至少读到 2 行")
        fh_rows = read_recent_learn_events(10, file_hash="abc", path=p)
        assert_true(any(r.get("file_hash") == "abc" for r in fh_rows), "按 file_hash 应能筛到")
    finally:
        try:
            os.unlink(p)
        except OSError:
            pass


# ============================================================
# 主入口
# ============================================================
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="万邦船舶询价系统 优化验证测试")
    parser.add_argument("--verbose", action="store_true", help="详细输出")
    args = parser.parse_args()

    ok = run_all(verbose=args.verbose)
    sys.exit(0 if ok else 1)
