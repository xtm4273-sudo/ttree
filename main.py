"""
万邦船舶询价系统 - Web Demo
上传文件(PDF/Word/Excel) → AI解析 → 智能匹配 → 在线查看 → 下载报价单
"""
import os
import sys
import time
import tempfile
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd

from app.document_parser import parse_document
from app.craft_library import load_craft_library, build_vector_index, load_vector_index, add_to_library
from app.match_engine import match_all_items
from app.excel_generator import generate_quotation_excel
import config

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
</style>
""", unsafe_allow_html=True)

# === 标题 ===
st.markdown('<div class="main-title">万邦船舶智能询价匹配系统</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">AI驱动 · 询价单解析 · 工艺库匹配 · 报价单生成</div>', unsafe_allow_html=True)

# === 侧边栏配置 ===
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
                    craft_items = load_craft_library(craft_lib_path)
                    index = build_vector_index(craft_items, "data")
                    st.success(f"完成！共 {len(craft_items)} 条")
                    st.rerun()
                except Exception as e:
                    st.error(f"错误: {e}")


# === 主界面 ===
# 加载索引状态
if "index_loaded" not in st.session_state:
    index, craft_items = load_vector_index("data")
    st.session_state["index_loaded"] = index is not None
    st.session_state["index"] = index
    st.session_state["craft_items"] = craft_items
else:
    index = st.session_state.get("index")
    craft_items = st.session_state.get("craft_items")

# Step 1: 上传
st.markdown("""
<div class="step-indicator">
    <div class="step-badge">1</div>
    <div class="step-text">上传询价单文件</div>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "拖拽或选择询价单文件",
    type=["pdf", "docx", "xlsx", "xls"],
    help="支持 PDF、Word、Excel 格式",
    label_visibility="collapsed",
)

if uploaded_file:
    suffix = os.path.splitext(uploaded_file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    # Step 2: 解析
    st.markdown("""
    <div class="step-indicator">
        <div class="step-badge">2</div>
        <div class="step-text">AI 智能解析</div>
    </div>
    """, unsafe_allow_html=True)

    if not config.LLM_API_KEY:
        st.error("请在左侧配置 API Key")
        st.stop()

    with st.status("正在解析文档...", expanded=True) as parse_status:
        def parse_log(msg):
            st.write(f"`{datetime.now().strftime('%H:%M:%S')}` {msg}")

        t_parse_start = time.time()
        parse_log(f"开始解析: {uploaded_file.name}")

        try:
            enquiry_items = parse_document(tmp_path, log_callback=parse_log)
        except Exception as e:
            parse_log(f"❌ 解析失败: {e}")
            import traceback
            st.code(traceback.format_exc(), language="text")
            os.unlink(tmp_path)
            st.stop()

        t_parse = time.time() - t_parse_start
        parse_log(f"解析完成! 共 {len(enquiry_items)} 条，耗时 {t_parse:.1f}s")
        parse_status.update(label=f"解析完成: {len(enquiry_items)} 条 ({t_parse:.1f}s)", state="complete")

    with st.expander("查看解析结果（前15条）", expanded=False):
        preview_data = []
        for item in enquiry_items[:15]:
            preview_data.append({
                "SFI编码": item.get("sfi_code") or "-",
                "标题": item["title"][:55],
                "数量": item.get("quantity") or "-",
                "单位": item.get("unit") or "-",
            })
        st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)

    # Step 3: 匹配
    st.markdown("""
    <div class="step-indicator">
        <div class="step-badge">3</div>
        <div class="step-text">AI 智能匹配</div>
    </div>
    """, unsafe_allow_html=True)

    if index is None:
        st.error("向量索引未构建，请先在左侧点击【构建索引】")
        st.stop()

    if st.button("开始匹配", type="primary"):
        progress_bar = st.progress(0, text="准备中...")

        def update_progress(current, total):
            pct = current / total
            eta = ""
            if current > 0 and hasattr(update_progress, '_start_time'):
                elapsed = time.time() - update_progress._start_time
                remaining = elapsed / current * (total - current)
                eta = f" | 预计剩余 {remaining:.0f}s"
            progress_bar.progress(pct, text=f"匹配中... {current}/{total}{eta}")
        update_progress._start_time = time.time()

        with st.status("正在匹配...", expanded=True) as match_status:
            def match_log(msg):
                st.write(f"`{datetime.now().strftime('%H:%M:%S')}` {msg}")

            t_match_start = time.time()

            try:
                match_results = match_all_items(
                    enquiry_items, index, craft_items,
                    progress_callback=update_progress,
                    log_callback=match_log,
                )
            except Exception as e:
                match_log(f"❌ 匹配失败: {e}")
                import traceback
                st.code(traceback.format_exc(), language="text")
                st.stop()

            t_match = time.time() - t_match_start
            match_log(f"全部匹配完成! 耗时 {t_match:.1f}s")
            match_status.update(label=f"匹配完成 ({t_match:.1f}s)", state="complete")

        progress_bar.progress(1.0, text="匹配完成！")
        st.session_state["match_results"] = match_results

    # Step 4: 结果
    if "match_results" in st.session_state:
        results = st.session_state["match_results"]

        st.markdown("""
        <div class="step-indicator">
            <div class="step-badge">4</div>
            <div class="step-text">匹配结果 & 导出</div>
        </div>
        """, unsafe_allow_html=True)

        # 统计
        total = len(results)
        high = sum(1 for r in results if r["confidence"] >= 80 and not r.get("is_new_item"))
        med = sum(1 for r in results if 60 <= r["confidence"] < 80 and not r.get("is_new_item"))
        low = sum(1 for r in results if 40 <= r["confidence"] < 60 and not r.get("is_new_item"))
        new_items = sum(1 for r in results if r.get("is_new_item"))

        st.markdown(f"""
        <div class="metric-row">
            <div class="metric-card metric-blue">
                <div class="metric-label">总条目</div>
                <div class="metric-value">{total}</div>
            </div>
            <div class="metric-card metric-green">
                <div class="metric-label">高置信</div>
                <div class="metric-value">{high}</div>
                <div class="metric-label">{high*100//max(total,1)}% 自动可用</div>
            </div>
            <div class="metric-card metric-yellow">
                <div class="metric-label">建议复核</div>
                <div class="metric-value">{med}</div>
            </div>
            <div class="metric-card metric-red">
                <div class="metric-label">需确认</div>
                <div class="metric-value">{low}</div>
            </div>
            <div class="metric-card metric-purple">
                <div class="metric-label">新增项</div>
                <div class="metric-value">{new_items}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # 结果表格
        table_data = []
        for r in results:
            best = r.get("best_match")
            is_new = r.get("is_new_item", False)
            suggested = r.get("suggested_entry")

            if is_new and suggested:
                matched = suggested.get("title", "")[:45]
                status = "新增"
            elif best:
                matched = best["craft_title"][:45]
                conf = r["confidence"]
                if conf >= 80:
                    status = "自动"
                elif conf >= 60:
                    status = "复核"
                else:
                    status = "确认"
            else:
                matched = "-"
                status = "无"

            table_data.append({
                "SFI编码": r.get("enquiry_sfi") or "-",
                "询价标题": r["enquiry_title"][:45],
                "匹配工艺": matched,
                "置信度": r["confidence"] if not is_new else 0,
                "状态": status,
            })

        df = pd.DataFrame(table_data)

        def highlight_status(val):
            colors = {
                "自动": "background-color: rgba(46,125,50,0.15); color: #2E7D32",
                "复核": "background-color: rgba(230,81,0,0.15); color: #E65100",
                "确认": "background-color: rgba(198,40,40,0.15); color: #C62828",
                "新增": "background-color: rgba(106,27,154,0.15); color: #6A1B9A",
                "无": "background-color: rgba(150,150,150,0.15); color: #666",
            }
            return colors.get(val, "")

        styled = df.style.map(highlight_status, subset=["状态"])
        st.dataframe(styled, use_container_width=True, height=450, hide_index=True)

        # 导出
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 1, 2])

        with col1:
            output_name = f"报价单_{datetime.now().strftime('%m%d_%H%M')}.xlsx"
            output_path = os.path.join(tempfile.gettempdir(), output_name)

            if st.button("生成Excel报价单"):
                with st.spinner("生成中..."):
                    generate_quotation_excel(results, output_path)
                st.session_state["excel_ready"] = True
                st.session_state["excel_path"] = output_path
                st.session_state["excel_name"] = output_name

        if st.session_state.get("excel_ready"):
            with col2:
                with open(st.session_state["excel_path"], "rb") as f:
                    st.download_button(
                        label="下载报价单",
                        data=f.read(),
                        file_name=st.session_state["excel_name"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    # 清理
    try:
        os.unlink(tmp_path)
    except:
        pass

else:
    # 欢迎界面
    st.markdown("""
    <div class="card" style="text-align:center; padding: 3rem;">
        <div style="font-size: 3rem; margin-bottom: 1rem;">📄 → 🤖 → 📊</div>
        <div style="color: #455A64; font-size: 1.1rem; margin-bottom: 2rem;">
            上传客户询价单文件，开始AI智能匹配
        </div>
        <div style="color: #546E7A; font-size: 0.9rem; text-align: left; max-width: 500px; margin: 0 auto;">
            <p><strong style="color:#1565C0">支持格式：</strong>PDF、Word (.docx)、Excel (.xlsx)</p>
            <p><strong style="color:#1565C0">工作流程：</strong></p>
            <ol style="padding-left: 1.2rem;">
                <li>AI 解析文档，识别维修项目条目</li>
                <li>每条项目与工艺库进行语义匹配</li>
                <li>标注置信度评分</li>
                <li>导出 Excel 报价单供人工审核</li>
            </ol>
            <p style="margin-top:1rem;"><strong style="color:#1565C0">置信度说明：</strong></p>
            <p><span class="conf-high">≥80 自动</span> 可直接使用</p>
            <p><span class="conf-med">60-79 复核</span> 建议人工复核</p>
            <p><span class="conf-low">40-59 确认</span> 需人工确认</p>
            <p><span class="conf-new">新增</span> 工艺库中不存在，AI生成建议</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
