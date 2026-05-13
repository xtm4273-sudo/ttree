"""
万邦船舶询价系统 - Web Demo
上传文件(PDF/Word/Excel) → AI解析 → 智能匹配 → 页面核对 → 导出报价单
"""
import os
import sys
import json
import tempfile
import hashlib
from collections import defaultdict
from datetime import datetime

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import streamlit as st


def _apply_streamlit_secrets_to_environ() -> None:
    """Streamlit Community Cloud：Secrets 仅在 st.secrets；须在首次 import config 前写入 os.environ。"""
    try:
        sec = st.secrets
    except Exception:
        return
    try:
        for key in sec:
            val = sec[key]
            if not isinstance(val, (str, int, float, bool)):
                continue
            name = str(key)
            if name not in os.environ or os.environ.get(name, "") == "":
                os.environ[name] = str(val)
    except Exception:
        return


_apply_streamlit_secrets_to_environ()

from app.document_parser import parse_document
from app.craft_library import (
    add_to_library,
    build_vector_index,
    craft_entry_dedupe_key,
    load_craft_library,
    load_vector_index,
    merge_template_with_saved_user_entries,
)
from app.craft_excel_import import import_learning_rows_from_quotation_excel
from app.learn_history import append_learn_event, read_recent_learn_events
from app.match_engine import match_all_items, match_single_item
from app.excel_generator import (
    generate_quotation_excel,
    match_results_to_preview_dataframe,
    merge_preview_dataframe_into_match_results,
)
from app.enquiry_history import append_run, clear_all_runs, list_runs, load_run
import config

# === 辅助函数 ===


class JobProgress:
    """
    将解析 / 匹配映射到单一进度条（各阶段为真实计数，不做假动画）。
    权重：提取文本 0~20%，解析条目 20~55%，智能匹配 55~100%。
    """

    EXTRACT_END = 0.20
    PARSE_END = 0.55
    MATCH_END = 1.0

    def __init__(self, progress_bar, caption_holder):
        self.bar = progress_bar
        self.cap = caption_holder

    def _show(self, p: float, text: str):
        p = min(max(float(p), 0.0), 1.0)
        self.bar.progress(p, text=(text or "处理中…")[:80])
        self.cap.caption(text or "")

    def on_doc(self, event: str, cur: int, tot: int, msg: str | None = None):
        tot = max(tot, 1)
        cur = min(max(cur, 0), tot)
        if event == "extracting_text":
            p = self.EXTRACT_END * (cur / tot)
        elif event in ("rule_extract", "segmenting_chunks"):
            # 20% ~ 50%：规则识别或 AI 分块的真实进度
            p = self.EXTRACT_END + (0.50 - self.EXTRACT_END) * (cur / tot)
        elif event == "post_processing":
            # 50% ~ 55%：整理 / 去重
            p = 0.50 + (self.PARSE_END - 0.50) * (cur / tot)
        else:
            p = self.EXTRACT_END
        self._show(p, msg or "解析中…")

    def skip_to_after_parse(self):
        self._show(self.PARSE_END, "解析已完成，开始智能匹配…")

    def on_match(self, cur: int, tot: int):
        tot = max(tot, 1)
        cur = min(max(cur, 0), tot)
        p = self.PARSE_END + (self.MATCH_END - self.PARSE_END) * (cur / tot)
        self._show(p, f"智能匹配中 {cur}/{tot}…")

    def after_match(self, n_items: int):
        if n_items <= 0:
            self._show(self.MATCH_END, "未识别到条目，请在下方预览中核对后导出。")
        else:
            self._show(self.MATCH_END, f"已完成 {n_items} 条匹配，请在下方预览表格中核对后导出 Excel。")

    def on_excel_start(self):
        self._show(0.96, "正在生成 Excel 报价单…")

    def on_excel_done(self):
        self._show(1.0, "Excel 已生成，可下载")

    def skip_to_after_match(self):
        self._show(self.MATCH_END, "匹配已完成，请在下方预览表格中核对后导出。")

    def all_done(self):
        self._show(1.0, "匹配已完成，请在下方核对预览后导出 Excel。")


QUOTATION_PIPELINE_STEPS = (
    "上传询价",
    "解析项目",
    "工艺匹配",
    "核对报价",
    "导出 Excel",
)


def derive_quotation_pipeline_step_states(file_key: str | None) -> list[str]:
    """
    五步主链每步 UI 状态：done / current / upcoming（不包含自主学习）。
    """
    n = len(QUOTATION_PIPELINE_STEPS)
    if not file_key:
        return ["upcoming"] * n

    up = "upcoming"
    cur = "current"
    done = "done"

    def pack(i_active: int) -> list[str]:
        return [done if j < i_active else (cur if j == i_active else up) for j in range(n)]

    fk = file_key
    enquiry_items = st.session_state.get("enquiry_items")
    if not isinstance(enquiry_items, list):
        enquiry_items = []
    mr = st.session_state.get("match_results")
    mrfk = st.session_state.get("match_results_file_key")
    matched = isinstance(mr, list) and mrfk == fk

    if len(enquiry_items) == 0:
        if not matched:
            return pack(1)
        # 解析结果为零条但已跑完匹配：仍进入核对/导出阶段
    elif not matched:
        return pack(2)

    excel_ready = bool(st.session_state.get("excel_ready"))
    excel_path = st.session_state.get("excel_path") or ""
    path_ok = bool(excel_path) and os.path.isfile(excel_path)
    rev = int(st.session_state.get("results_revision", 0))
    export_rev = st.session_state.get("quotation_export_revision")
    export_matches_rev = export_rev is None or int(export_rev) == rev
    export_ok = excel_ready and path_ok and export_matches_rev

    if export_ok:
        return [done] * n

    if excel_ready and path_ok and not export_matches_rev:
        return pack(4)

    return pack(3)


def _quotation_pipeline_stepper_html(states: list[str]) -> str:
    titles = QUOTATION_PIPELINE_STEPS
    if len(states) != len(titles):
        states = (list(states) + ["upcoming"] * len(titles))[: len(titles)]

    def circle_inner(i: int, st: str) -> str:
        if st == "done":
            return "✓"
        return str(i + 1)

    def circle_class(st: str) -> str:
        if st == "done":
            return "ps-circle ps-done"
        if st == "current":
            return "ps-circle ps-current"
        return "ps-circle ps-wait"

    def title_class(st: str) -> str:
        if st == "done":
            return "ps-title ps-title-done"
        if st == "current":
            return "ps-title ps-title-current"
        return "ps-title ps-title-wait"

    parts: list[str] = ['<div class="pipeline-stepper"><div class="ps-row">']
    for i, (title, st) in enumerate(zip(titles, states, strict=True)):
        inner = circle_inner(i, st)
        parts.append('<div class="ps-node">')
        parts.append(f'<div class="{circle_class(st)}">{inner}</div>')
        parts.append(f'<div class="{title_class(st)}">{title}</div>')
        parts.append("</div>")
        if i < len(titles) - 1:
            line_done = states[i] == "done"
            ln = "ps-line ps-line-done" if line_done else "ps-line"
            parts.append(f'<div class="{ln}" aria-hidden="true"></div>')
    parts.append("</div></div>")
    return "".join(parts)


def render_quotation_pipeline_stepper(file_key: str | None):
    if not file_key:
        return
    states = derive_quotation_pipeline_step_states(file_key)
    st.markdown(_quotation_pipeline_stepper_html(states), unsafe_allow_html=True)
    rev = int(st.session_state.get("results_revision", 0))
    er = st.session_state.get("quotation_export_revision")
    if (
        er is not None
        and int(er) != rev
        and st.session_state.get("excel_ready")
        and st.session_state.get("excel_path")
        and os.path.isfile(st.session_state["excel_path"])
    ):
        st.caption("报价表已相对上次导出发生变更，建议重新点击「导出 Excel 报价单」。")


def init_session_state():
    """初始化 session state 中的索引和工艺库数据"""
    if "index_loaded" not in st.session_state:
        index, craft_items = load_vector_index("data")
        st.session_state["index_loaded"] = index is not None
        st.session_state["index"] = index
        st.session_state["craft_items"] = craft_items


def refresh_vector_index_session():
    """从磁盘重载索引到 session（学习入库或批量导入后调用）。"""
    index, craft_items = load_vector_index("data")
    st.session_state["index_loaded"] = index is not None
    st.session_state["index"] = index
    st.session_state["craft_items"] = craft_items


def render_sidebar():
    """渲染侧边栏配置，返回 (index, craft_items)"""
    with st.sidebar:
        st.markdown("### 系统配置")

        api_key = st.text_input("API Key", value=config.LLM_API_KEY, type="password")
        if api_key:
            config.LLM_API_KEY = api_key
            config.EMBED_API_KEY = api_key

        base_url = st.text_input("接口地址", value=config.LLM_BASE_URL)
        if base_url:
            config.LLM_BASE_URL = base_url
            config.EMBED_BASE_URL = base_url

        model = st.text_input("对话模型", value=config.LLM_MODEL)
        if model:
            config.LLM_MODEL = model

        st.divider()

        index, craft_items = load_vector_index("data")
        if index is not None:
            st.success(f"工艺库已加载: {len(craft_items)} 条")
        else:
            st.warning("向量索引未构建")

        craft_lib_path = st.text_input("工艺库路径", value=config.CRAFT_LIBRARY_PATH)

        if st.button("构建索引"):
            if not config.LLM_API_KEY:
                st.error("请先填入API Key")
            else:
                with st.spinner("正在构建..."):
                    try:
                        template_only = load_craft_library(craft_lib_path)
                        craft_items = merge_template_with_saved_user_entries(
                            template_only, "data"
                        )
                        index = build_vector_index(craft_items, "data")
                        st.success(
                            f"完成！共 {len(craft_items)} 条（已合并此前「学习入库」条目）"
                        )
                        st.rerun()
                    except Exception:
                        st.error("构建索引失败，请检查配置后重试。")

        with st.expander("从报价单 Excel 批量学习", expanded=False):
            st.caption(
                "使用本页导出后的报价单：在表格中填好「匹配工艺（首选）」等列后，"
                "将「学习入库」列填 **ADD**，导出 Excel 并保存，在此上传。"
            )
            bulk_xlsx = st.file_uploader(
                "上传已填报价单",
                type=["xlsx"],
                key="craft_bulk_learn_uploader",
            )
            if st.button("导入到工艺库", key="craft_bulk_learn_btn"):
                if not bulk_xlsx:
                    st.warning("请先选择 xlsx 文件")
                elif not config.LLM_API_KEY:
                    st.error("请先填入 API Key（向量化需要）")
                else:
                    with st.spinner("正在解析并写入…"):
                        try:
                            n, errs = import_learning_rows_from_quotation_excel(
                                bulk_xlsx.getvalue(), "data"
                            )
                            if errs:
                                st.text_area("导入说明", "\n".join(errs[:80]), height=160)
                            if n:
                                st.success(f"已写入工艺库 {n} 条")
                                refresh_vector_index_session()
                            else:
                                st.info("没有可写入的新增行（检查「学习入库」是否为 ADD）")
                            st.rerun()
                        except Exception as e:
                            st.error(f"导入失败: {e}")

        with st.expander("高级", expanded=False):
            st.caption("解析历史存储目录")
            st.code(os.path.normpath(config.ENQUIRY_HISTORY_DIR), language=None)
            if st.checkbox("确认删除全部解析历史（不可恢复）", key="chk_clear_enquiry_history"):
                if st.button("清空解析历史"):
                    n = clear_all_runs()
                    st.success(f"已清空，共删除 {n} 个 run 文件。")
                    st.rerun()

    return index, craft_items


def run_parsing_pipeline(uploaded_file, progress_callback=None) -> tuple:
    """解析上传的询价单文件，返回 (enquiry_items, tmp_path) 或 st.stop()"""
    suffix = os.path.splitext(uploaded_file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    if not config.LLM_API_KEY:
        st.error("请在左侧配置 API Key")
        st.stop()

    try:
        enquiry_items = parse_document(
            tmp_path,
            log_callback=None,
            progress_callback=progress_callback,
        )
    except Exception:
        st.error("文件解析失败，请检查文件内容后重试。")
        os.unlink(tmp_path)
        st.stop()

    return enquiry_items, tmp_path


def run_matching_pipeline(
    enquiry_items,
    index,
    craft_items,
    file_key: str,
    job_progress: JobProgress | None = None,
):
    """运行智能匹配并写入 session（Excel 在页面确认后由用户手动导出）。"""

    if index is None:
        st.error("向量索引未构建，请先在左侧点击【构建索引】")
        st.stop()

    match_was_cached = (
        st.session_state.get("match_results_file_key") == file_key
        and "match_results" in st.session_state
    )

    if not match_was_cached:
        try:
            match_results = match_all_items(
                enquiry_items,
                index,
                craft_items,
                progress_callback=(lambda c, t: job_progress.on_match(c, t)) if job_progress else None,
            )
        except Exception:
            st.error("智能匹配失败，请稍后重试。")
            st.stop()
        st.session_state["match_results"] = match_results
        st.session_state["match_results_file_key"] = file_key
        st.session_state["results_revision"] = st.session_state.get("results_revision", 0) + 1
        for k in (
            "excel_ready",
            "excel_path",
            "excel_name",
            "quotation_preview_stamp",
            "quotation_export_revision",
        ):
            st.session_state.pop(k, None)
        if job_progress:
            job_progress.after_match(len(enquiry_items))
    elif job_progress:
        job_progress.skip_to_after_match()


def render_quotation_preview_and_export(active_file_key: str | None):
    """匹配完成后：可编辑预览表，确认后导出 Excel。"""
    if not active_file_key:
        return
    results = st.session_state.get("match_results")
    if not results:
        return

    st.divider()
    st.subheader("报价单预览")
    st.caption(
        "请在本表核对或修改各列；确认无误后点击「导出 Excel 报价单」。"
        "导出文件可用于外发或侧栏「从报价单 Excel 批量学习」。"
    )

    rev = int(st.session_state.get("results_revision", 0))
    stamp = (active_file_key, rev)
    if st.session_state.get("quotation_preview_stamp") != stamp:
        st.session_state["quotation_preview_stamp"] = stamp
        st.session_state["quotation_edited_df"] = match_results_to_preview_dataframe(results)
        for k in ("excel_ready", "excel_path", "excel_name", "quotation_export_revision"):
            st.session_state.pop(k, None)

    from streamlit.column_config import TextColumn

    disabled_cols = [
        "置信度%",
        "状态",
        "匹配原因",
        "处理状态",
    ]
    edited_df = st.data_editor(
        st.session_state["quotation_edited_df"],
        key=f"quotation_editor_{active_file_key}_{rev}",
        num_rows="fixed",
        use_container_width=True,
        height=min(620, 140 + 32 * len(results)),
        disabled=disabled_cols,
        column_config={
            "匹配原因": TextColumn(width="large"),
            "询价标题": TextColumn(width="medium"),
        },
    )
    st.session_state["quotation_edited_df"] = edited_df

    if st.button("导出 Excel 报价单", type="primary", key=f"export_quotation_{active_file_key}_{rev}"):
        enq = st.session_state.get("enquiry_items")
        merged = merge_preview_dataframe_into_match_results(edited_df, results, enq)
        st.session_state["match_results"] = merged
        output_name = f"报价单_{datetime.now().strftime('%m%d_%H%M')}.xlsx"
        output_path = os.path.join(tempfile.gettempdir(), output_name)
        try:
            generate_quotation_excel(merged, output_path)
        except Exception:
            st.error("报价单生成失败，请稍后重试。")
            return
        st.session_state["excel_ready"] = True
        st.session_state["excel_path"] = output_path
        st.session_state["excel_name"] = output_name
        st.session_state["quotation_export_revision"] = int(st.session_state.get("results_revision", 0))
        st.session_state["quotation_edited_df"] = match_results_to_preview_dataframe(merged)
        st.success("已根据当前表格生成 Excel，可下载。")
        st.rerun()

    if st.session_state.get("excel_ready") and st.session_state.get("excel_path"):
        p = st.session_state["excel_path"]
        if os.path.isfile(p):
            with open(p, "rb") as f:
                st.download_button(
                    label="下载报价单 Excel",
                    data=f.read(),
                    file_name=st.session_state.get("excel_name") or "报价单.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="secondary",
                )
        else:
            st.caption("上次导出的临时文件已不存在，请重新点击「导出 Excel 报价单」。")


_LEARNING_HIERARCHY_FLAGS = frozenset(
    {
        "parent_without_children",
        "possible_missing_children",
        "missing_sfi_hierarchy_uncertain",
    }
)

# 分组展示顺序（与 learning_bucket 返回的 id 一致）
_LEARNING_BUCKET_ORDER = (
    "no_vector",
    "hierarchy_risk",
    "low_confidence",
    "tight_top2",
    "review_other",
)

_LEARNING_BUCKET_GROUP_TITLE = {
    "no_vector": "库无向量候选",
    "hierarchy_risk": "待确认：解析层级风险",
    "low_confidence": "待确认：置信度偏低",
    "tight_top2": "待确认：Top1/Top2 难分",
    "review_other": "待确认：建议人工核对",
}

_LEARNING_BUCKET_HELP = {
    "no_vector": "向量检索未命中任何工艺（与 Excel「库无候选」一致）。可在此补录标准工艺后重匹配。",
    "hierarchy_risk": "解析认为可能缺子级或层级不完整，系统已压置信度，请人工核对后再决定是否入库。",
    "low_confidence": "已有候选但综合置信度低于阈值，匹配可能不准。",
    "tight_top2": "前两名候选置信度接近，系统难以自动取舍。",
    "review_other": "系统建议人工核对匹配结果（如待审核状态等）。",
}


def learning_bucket(res: dict) -> tuple[str, str]:
    """
    自主学习面板用：互斥优先级分类。
    返回 (bucket_id, 短标签) ，短标签用于 Expander 标题前缀。
    """
    if (res.get("decision_path") or "") == "no_candidate":
        return ("no_vector", "库无向量候选")

    if not res.get("needs_human_review"):
        return ("review_other", "待确认")

    qf = res.get("quality_flags") or []
    if isinstance(qf, list) and _LEARNING_HIERARCHY_FLAGS.intersection(qf):
        return ("hierarchy_risk", "解析层级风险")

    conf = res.get("confidence")
    if conf is not None and conf < config.CONFIDENCE_LEVELS["low"]:
        return ("low_confidence", "置信度偏低")

    matches = res.get("matches") or []
    if len(matches) >= 2:
        c0 = float(matches[0].get("confidence") or 0)
        c1 = float(matches[1].get("confidence") or 0)
        if c0 - c1 < 5:
            return ("tight_top2", "Top1/Top2难分")

    return ("review_other", "建议人工核对")


def count_self_learning_eligible_rows() -> int:
    """与自主学习面板相同的「可关注」条目数，用于可选分支标题提示。"""
    md = st.session_state.get("match_results") or []
    eq = st.session_state.get("enquiry_items") or []
    if not md or not eq:
        return 0
    n = min(len(md), len(eq))
    c = 0
    for i in range(n):
        res = md[i]
        if (
            res.get("is_new_item")
            or res.get("needs_human_review")
            or res.get("decision_path") == "no_candidate"
        ):
            c += 1
    return c


def render_self_learning_panel(embedded: bool = False):
    """无匹配/待确认等条目：按类别分组展示，人工确认后可写入工艺库并重匹配。"""
    md = st.session_state.get("match_results") or []
    eq = st.session_state.get("enquiry_items") or []
    if not md or not eq:
        if embedded:
            st.caption("暂无匹配结果；完成上方解析与匹配后，如需补录工艺可再展开此处。")
        return
    n = min(len(md), len(eq))
    if not embedded:
        st.divider()
        st.subheader("工艺库自主学习（人机闭环）")
    st.caption(
        "下方按**类别**列出需要关注的条目：**库无向量候选**表示工艺库里搜不到相近项；"
        "**待确认**表示有候选但置信低、前两名太接近或解析层级有风险等，请按公司主数据填写工艺标题与 SFI 后入库。"
        "提交后写入 **data/craft_library.json** 并更新向量索引，**下次上传询价单或重新匹配**即可从工艺库命中新条目（与是否导出 Excel 无关）。"
        "「学习入库」列在上方报价单预览中显示 **已入库** 后，导出 Excel 可留档。"
    )

    learned_summary = []
    for i in range(n):
        r = md[i]
        if (r.get("quotation_learn_action") or "").strip() == "已入库":
            learned_summary.append(
                {
                    "条目序号": i + 1,
                    "询价标题": (r.get("enquiry_title") or "")[:120],
                    "入库工艺": (r.get("learnt_craft_title") or "")[:120],
                    "工艺库ID": r.get("learnt_craft_id"),
                    "入库时间": r.get("learnt_at") or "",
                }
            )
    with st.expander("已在本会话学习入库的条目", expanded=False):
        if learned_summary:
            st.dataframe(pd.DataFrame(learned_summary), use_container_width=True, hide_index=True)
        else:
            st.caption("本会话尚未通过下方表单完成单条入库（批量 Excel 导入见侧栏，记录在历史表中）。")

    st.markdown("##### 学习入库记录（本地）")
    only_current = st.checkbox(
        "仅显示当前询价单（按文件指纹筛选）",
        value=True,
        key="learn_history_filter_current_file",
    )
    fh = st.session_state.get("active_file_key") or None
    events = read_recent_learn_events(
        50, file_hash=fh if only_current and fh else None
    )
    st.caption(
        "审计文件：`data/learn_history.jsonl`（每行一条 JSON，便于追溯谁、何时、写入哪条工艺）。"
    )
    if events:
        st.dataframe(pd.DataFrame(events), use_container_width=True, height=min(360, 80 + 28 * len(events)))
    else:
        st.info("暂无记录；完成单条入库或侧栏批量导入后会出现。")

    st.divider()
    st.markdown("##### 待学习 / 待确认条目")

    rows: list[tuple[int, dict, dict, str, str]] = []
    for i in range(n):
        res, enq = md[i], eq[i]
        eligible = (
            res.get("is_new_item")
            or res.get("needs_human_review")
            or res.get("decision_path") == "no_candidate"
        )
        if not eligible:
            continue
        bid, badge = learning_bucket(res)
        rows.append((i, res, enq, bid, badge))

    if not rows:
        st.info("当前结果中暂无需要学习入库的条目（无「库无候选 / 待确认」类）。")
        return

    by_bucket: dict[str, list[tuple[int, dict, dict, str, str]]] = defaultdict(list)
    for row in rows:
        by_bucket[row[3]].append(row)

    for bid in _LEARNING_BUCKET_ORDER:
        group = by_bucket.get(bid)
        if not group:
            continue
        title = _LEARNING_BUCKET_GROUP_TITLE.get(bid, bid)
        st.markdown(f"#### {title}（{len(group)} 条）")
        st.caption(_LEARNING_BUCKET_HELP.get(bid, ""))
        for i, res, enq, _, badge in group:
            et = (enq.get("title") or "")[:60]
            with st.expander(f"[{badge}] 条目 {i + 1}：{et}", expanded=False):
                st.caption(
                    f"path={res.get('decision_path')} | llm_called={res.get('llm_called')} | "
                    f"review={res.get('review_status')} | conf={res.get('confidence')}"
                )
                with st.form(key=f"craft_learn_form_{i}"):
                    sfi_in = st.text_input(
                        "人工确认 SFI（选填）",
                        value=(enq.get("sfi_code") or res.get("enquiry_sfi") or ""),
                        key=f"learn_sfi_{i}",
                    )
                    title_in = st.text_input(
                        "工艺标题（必填，英文大写推荐）",
                        value=(enq.get("title") or res.get("enquiry_title") or "")[:500],
                        key=f"learn_title_{i}",
                    )
                    desc_in = st.text_area(
                        "工艺说明 / 描述（选填，默认用解析描述）",
                        value=(
                            (enq.get("description") or res.get("enquiry_description") or "")[
                                :1500
                            ]
                        ),
                        height=80,
                        key=f"learn_desc_{i}",
                    )
                    unit_in = st.text_input(
                        "单位", value=(enq.get("unit") or "LOT"), key=f"learn_unit_{i}"
                    )
                    submitted = st.form_submit_button("加入工艺库并重匹配本条")
                if submitted:
                    if not (title_in or "").strip():
                        st.error("请填写工艺标题")
                    elif not config.LLM_API_KEY:
                        st.error("请配置 API Key")
                    else:
                        ok, new_id = add_to_library(
                            {
                                "sfi_code": (sfi_in or "").strip(),
                                "title": (title_in or "").strip(),
                                "description": (desc_in or "").strip(),
                                "unit": (unit_in or "LOT").strip(),
                            },
                            "data",
                        )
                        if not ok:
                            st.warning("未写入（可能重复或索引不存在）。请先构建索引。")
                        else:
                            idx, cis = load_vector_index("data")
                            if idx is None:
                                st.error("索引加载失败")
                            else:
                                new_res = match_single_item(enq, idx, cis, item_serial=i)
                                t_title = (title_in or "").strip()
                                t_sfi = (sfi_in or "").strip()
                                new_res["quotation_learn_action"] = "已入库"
                                new_res["learnt_at"] = datetime.now().isoformat(
                                    timespec="seconds"
                                )
                                new_res["learnt_craft_id"] = new_id
                                new_res["learnt_craft_title"] = t_title
                                new_res["learnt_craft_sfi"] = t_sfi
                                st.session_state["match_results"][i] = new_res
                                st.session_state["results_revision"] = (
                                    st.session_state.get("results_revision", 0) + 1
                                )
                                append_learn_event(
                                    {
                                        "source": "ui_single",
                                        "file_hash": st.session_state.get("active_file_key")
                                        or "",
                                        "item_index": i,
                                        "enquiry_title": (enq.get("title") or "")[:500],
                                        "enquiry_sfi": (enq.get("sfi_code") or "")[:80],
                                        "craft_id": new_id,
                                        "craft_title": t_title[:500],
                                        "craft_sfi": t_sfi[:80],
                                        "dedupe_key": str(
                                            craft_entry_dedupe_key(
                                                {"sfi_code": t_sfi, "title": t_title}
                                            )
                                        ),
                                    }
                                )
                                for k in (
                                    "excel_ready",
                                    "excel_path",
                                    "excel_name",
                                    "quotation_preview_stamp",
                                    "quotation_export_revision",
                                ):
                                    st.session_state.pop(k, None)
                                st.success(
                                    "已入库并重匹配；上方报价单预览「学习入库」列为「已入库」，请导出 Excel 留档。"
                                )
                                refresh_vector_index_session()
                                st.rerun()


# 页面配置
st.set_page_config(
    page_title="万邦船舶询价匹配系统",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# === 自定义样式 ===
st.markdown("""
<style>
    /* 整体背景 - 浅色 */
    .stApp {
        background: linear-gradient(135deg, #f0f4f8 0%, #e8eef5 50%, #f5f7fa 100%);
    }

    /* 隐藏默认header和footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* 主标题 */
    .main-title {
        text-align: center;
        padding: 1.5rem 0 0.5rem 0;
        background: linear-gradient(90deg, #1565C0, #0D47A1, #1976D2);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.2rem;
        font-weight: 800;
        letter-spacing: 2px;
    }

    .sub-title {
        text-align: center;
        color: #546E7A;
        font-size: 0.95rem;
        margin-bottom: 2rem;
    }

    /* 卡片容器 */
    .card {
        background: #ffffff;
        border: 1px solid #e0e7ee;
        border-radius: 16px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }

    .card-header {
        color: #1565C0;
        font-size: 1.1rem;
        font-weight: 600;
        margin-bottom: 0.8rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* 指标卡片 */
    .metric-row {
        display: flex;
        gap: 1rem;
        margin: 1rem 0;
    }

    .metric-card {
        flex: 1;
        background: #ffffff;
        border: 1px solid #e0e7ee;
        border-radius: 12px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    }

    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        margin: 0.3rem 0;
    }

    .metric-label {
        font-size: 0.8rem;
        color: #607D8B;
        text-transform: uppercase;
    }

    .metric-green .metric-value { color: #2E7D32; }
    .metric-yellow .metric-value { color: #E65100; }
    .metric-red .metric-value { color: #C62828; }
    .metric-purple .metric-value { color: #6A1B9A; }
    .metric-blue .metric-value { color: #1565C0; }

    /* 步骤指示器 */
    .step-indicator {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        margin: 1.5rem 0 1rem 0;
    }

    .step-badge {
        background: linear-gradient(135deg, #1976D2, #1565C0);
        color: white;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.8rem;
        font-weight: 700;
    }

    .step-text {
        color: #37474F;
        font-size: 1rem;
        font-weight: 500;
    }

    /* 上传区域 */
    .upload-zone {
        border: 2px dashed rgba(21, 101, 192, 0.3);
        border-radius: 16px;
        padding: 2rem;
        text-align: center;
        background: rgba(21, 101, 192, 0.02);
        transition: all 0.3s;
    }

    .upload-zone:hover {
        border-color: rgba(21, 101, 192, 0.6);
        background: rgba(21, 101, 192, 0.05);
    }

    /* 置信度标签 */
    .conf-high {
        background: rgba(46, 125, 50, 0.1);
        color: #2E7D32;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
    }

    .conf-med {
        background: rgba(230, 81, 0, 0.1);
        color: #E65100;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
    }

    .conf-low {
        background: rgba(198, 40, 40, 0.1);
        color: #C62828;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
    }

    .conf-new {
        background: rgba(106, 27, 154, 0.1);
        color: #6A1B9A;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
    }

    /* 表格美化 */
    div[data-testid="stDataFrame"] {
        border: 1px solid #e0e7ee;
        border-radius: 12px;
        overflow: hidden;
    }

    /* 按钮 */
    .stButton > button {
        background: linear-gradient(135deg, #1976D2, #1565C0);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        font-weight: 600;
        transition: all 0.3s;
    }

    .stButton > button:hover {
        background: linear-gradient(135deg, #1E88E5, #1976D2);
        box-shadow: 0 4px 12px rgba(25, 118, 210, 0.4);
    }

    /* 侧边栏 */
    section[data-testid="stSidebar"] {
        background: #f8fafc;
        border-right: 1px solid #e0e7ee;
    }

    /* 进度条 */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #1976D2, #42A5F5);
    }

    /* expander */
    .streamlit-expanderHeader {
        background: #f5f7fa;
        border-radius: 8px;
    }

    /* 主流程五步 Stepper（与帆软式里程碑条类似） */
    .pipeline-stepper {
        background: #ffffff;
        border: 1px solid #e0e7ee;
        border-radius: 12px;
        padding: 0.75rem 1rem 1rem 1rem;
        margin: 0.5rem 0 1rem 0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
    }
    .pipeline-stepper .ps-row {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        width: 100%;
        flex-wrap: nowrap;
        gap: 0;
    }
    .pipeline-stepper .ps-node {
        display: flex;
        flex-direction: column;
        align-items: center;
        flex: 0 1 120px;
        min-width: 0;
        text-align: center;
    }
    .pipeline-stepper .ps-circle {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.85rem;
        font-weight: 700;
        flex-shrink: 0;
    }
    .pipeline-stepper .ps-done {
        background: #90CAF9;
        color: #ffffff;
    }
    .pipeline-stepper .ps-current {
        background: #1976D2;
        color: #ffffff;
    }
    .pipeline-stepper .ps-wait {
        background: #ECEFF1;
        color: #90A4AE;
        border: 1px solid #CFD8DC;
    }
    .pipeline-stepper .ps-title {
        margin-top: 6px;
        font-size: 0.78rem;
        line-height: 1.25;
        word-break: break-all;
    }
    .pipeline-stepper .ps-title-done,
    .pipeline-stepper .ps-title-current {
        color: #263238;
    }
    .pipeline-stepper .ps-title-current {
        font-weight: 700;
    }
    .pipeline-stepper .ps-title-wait {
        color: #B0BEC5;
    }
    .pipeline-stepper .ps-line {
        flex: 1 1 12px;
        height: 2px;
        margin-top: 15px;
        min-width: 6px;
        background: #E0E7EE;
        align-self: flex-start;
    }
    .pipeline-stepper .ps-line-done {
        background: #90CAF9;
    }
</style>
""", unsafe_allow_html=True)

# === 标题 ===
st.markdown('<div class="main-title">万邦船舶智能询价匹配系统</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">上传询价单 → 智能匹配 → 页面核对 → 导出 Excel 报价单</div>', unsafe_allow_html=True)

# === 主流程 ===

# 启动配置校验
from app.validator import validate_config
config_errors = validate_config()
if config_errors:
    for err in config_errors:
        st.warning(f"配置警告: {err}")

# 初始化 session state
init_session_state()

# 侧边栏
index, craft_items = render_sidebar()

tab_upload, tab_history = st.tabs(["上传处理", "历史解析"])

with tab_upload:
    uploaded_file = st.file_uploader(
        "上传询价单文件",
        type=["pdf", "docx", "xlsx", "xls"],
        help="支持 PDF、Word、Excel 格式。上传后将自动解析并匹配，在页面核对后可导出报价单。",
    )
    if uploaded_file:
        pass
    elif st.session_state.get("pipeline_from_history"):
        st.info("当前会话已从「历史解析」载入数据，进度与导出见页面下方。")
    else:
        st.markdown(
            """
    <div class="card" style="text-align:center; padding: 3rem;">
        <div style="font-size: 3rem; margin-bottom: 1rem;">📄 → 📊</div>
        <div style="color: #455A64; font-size: 1.1rem; margin-bottom: 2rem;">
            上传后将自动解析、匹配工艺库；在页面表格中核对确认后再导出 Excel。
        </div>
    </div>
    """,
            unsafe_allow_html=True,
        )

with tab_history:
    runs = list_runs()
    if not runs:
        st.info("暂无解析记录。在「上传处理」中成功解析一次后，记录会保存在此供查看。")
    else:
        id_to_label = {
            r["run_id"]: (
                f"{(r.get('created_at') or '')[:19].replace('T', ' ')} UTC | "
                f"{r.get('original_filename') or '(未命名)'} | "
                f"{r.get('item_count', 0)} 条 | "
                f"{(r.get('file_hash') or '')[:8]}…"
            )
            for r in runs
        }
        run_ids = [r["run_id"] for r in runs]
        pick = st.selectbox(
            "选择一条解析记录",
            options=run_ids,
            format_func=lambda rid: id_to_label.get(rid, rid),
            key="history_select_run_id",
        )
        rec = load_run(pick) if pick else None
        if rec:
            items = rec.get("enquiry_items") or []
            if items:
                st.dataframe(
                    pd.DataFrame(items),
                    use_container_width=True,
                    height=min(480, 120 + 28 * len(items)),
                )
            else:
                st.warning("该记录无解析条目。")
            payload = json.dumps(rec, ensure_ascii=False, indent=2).encode("utf-8")
            st.download_button(
                label="下载完整 JSON",
                data=payload,
                file_name=f"enquiry_run_{rec.get('run_id', 'export')}.json",
                mime="application/json",
            )
            if st.button("载入并生成报价单", key="history_load_match"):
                st.session_state["enquiry_items"] = list(items)
                st.session_state["active_file_key"] = rec.get("file_hash") or ""
                st.session_state["pipeline_from_history"] = True
                st.session_state["from_history_run_id"] = rec.get("run_id") or ""
                for key in [
                    "match_results",
                    "match_results_file_key",
                    "excel_ready",
                    "excel_path",
                    "excel_name",
                    "quotation_preview_stamp",
                    "quotation_edited_df",
                    "quotation_export_revision",
                ]:
                    st.session_state.pop(key, None)
                st.rerun()

# --- 统一管线：本地上传 或 历史载入 ---
if uploaded_file:
    st.session_state.pop("from_history_run_id", None)
    st.session_state.pop("pipeline_from_history", None)

    file_bytes = uploaded_file.getvalue()
    file_key = hashlib.md5(file_bytes).hexdigest()

    if st.session_state.get("active_file_key") != file_key:
        st.session_state["active_file_key"] = file_key
        for key in [
            "enquiry_items",
            "match_results",
            "match_results_file_key",
            "excel_ready",
            "excel_path",
            "excel_name",
            "quotation_preview_stamp",
            "quotation_edited_df",
            "quotation_export_revision",
        ]:
            st.session_state.pop(key, None)

elif st.session_state.get("pipeline_from_history"):
    file_key = st.session_state.get("active_file_key") or ""
    if not file_key:
        st.error("历史记录缺少 file_hash，无法继续匹配。")
        st.stop()
else:
    file_key = None

if uploaded_file or st.session_state.get("pipeline_from_history"):
    if st.session_state.get("from_history_run_id") and not uploaded_file:
        st.caption(f"当前数据来源：历史记录 `{st.session_state['from_history_run_id']}`")

    progress_bar = st.progress(0.0, text="准备处理…")
    caption_el = st.empty()
    jp = JobProgress(progress_bar, caption_el)

    if uploaded_file:
        fully_done = (
            st.session_state.get("match_results_file_key") == file_key
            and st.session_state.get("match_results") is not None
        )

        if fully_done:
            jp.all_done()
            enquiry_items = st.session_state.get("enquiry_items", [])
            run_matching_pipeline(enquiry_items, index, craft_items, file_key, jp)
        else:
            if "enquiry_items" not in st.session_state:
                enquiry_items, parse_tmp_path = run_parsing_pipeline(
                    uploaded_file,
                    progress_callback=jp.on_doc,
                )
                st.session_state["enquiry_items"] = enquiry_items
                try:
                    append_run(uploaded_file.name, file_key, enquiry_items)
                except Exception:
                    st.caption("（解析历史写入失败，可检查目录权限）")
                try:
                    os.unlink(parse_tmp_path)
                except OSError:
                    pass
            else:
                enquiry_items = st.session_state["enquiry_items"]
                jp.skip_to_after_parse()

            run_matching_pipeline(enquiry_items, index, craft_items, file_key, jp)
    else:
        enquiry_items = st.session_state.get("enquiry_items") or []
        file_key = st.session_state.get("active_file_key")
        fully_done = (
            st.session_state.get("match_results_file_key") == file_key
            and st.session_state.get("match_results") is not None
        )
        if fully_done:
            jp.all_done()
        else:
            jp.skip_to_after_parse()
        run_matching_pipeline(enquiry_items, index, craft_items, file_key, jp)

    render_quotation_pipeline_stepper(file_key)
    render_quotation_preview_and_export(file_key)

n_learn = count_self_learning_eligible_rows()
_learn_title = "工艺库自主学习（可选）"
if n_learn:
    _learn_title += f" — 本单有 {n_learn} 条可关注"
with st.expander(_learn_title, expanded=False):
    render_self_learning_panel(embedded=True)
