"""
配置文件 - API密钥和模型参数
"""
import os
try:
    from dotenv import load_dotenv
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    if os.path.exists(env_path):
        load_dotenv(env_path)
except ImportError:
    pass  # dotenv not installed. use system env vars or defaults.

# === LLM API 配置 ===
LLM_API_KEY = os.getenv("LLM_API_KEY", "")
LLM_BASE_URL = os.getenv("LLM_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")
LLM_MODEL = os.getenv("LLM_MODEL", "qwen-plus")

# === Embedding 配置 ===
EMBED_API_KEY = os.getenv("EMBED_API_KEY", "") or LLM_API_KEY
EMBED_BASE_URL = os.getenv("EMBED_BASE_URL", "") or LLM_BASE_URL
EMBED_MODEL = os.getenv("EMBED_MODEL", "text-embedding-v3")

# === 匹配参数 ===
TOP_K = 10  # 向量检索返回候选数

# 匹配阶段审计：每条匹配追加一行 JSON（含是否调用大模型精排、门控原因、向量 Top1/2 等）
# 设为非空路径则启用，例如 MATCH_LLM_AUDIT_JSONL=logs/match_llm_audit.jsonl
MATCH_LLM_AUDIT_JSONL = os.getenv("MATCH_LLM_AUDIT_JSONL", "").strip()

# === 策略开关（准确度优先，减少冗余LLM调用） ===
ENABLE_RULE_FIRST = True
ENABLE_LLM_RERANK = True
ENABLE_AUTO_SUGGEST_NEW_ENTRY = False

# LLM精排触发门槛（仅在不确定时触发）
LLM_RERANK_TOP1_MIN = 70
LLM_RERANK_GAP_MAX = 5
LLM_RERANK_VECTOR_MIN = 72

CONFIDENCE_WEIGHTS = {
    "vector_similarity": 0.40,  # 向量相似度权重
    "llm_score": 0.35,          # LLM打分权重
    "sfi_match": 0.25,          # SFI编码匹配权重（无编码时自动归零）
}

# === 置信度分级阈值 ===
CONFIDENCE_LEVELS = {
    "high": 80,      # >=80 高置信，自动可用
    "medium": 60,    # 60-79 中等，建议复核
    "low": 40,       # 40-59 低，需人工确认
    # <40 无匹配/不可信，触发AI生成建议条目
}

# === 文件路径 ===
_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
_REPO_CRAFT_TEMPLATE = os.path.join(_PROJECT_ROOT, "data", "craft_template.xlsx")
_LEGACY_CRAFT_TEMPLATE = r"C:\Users\24895\Desktop\万邦船舶询价\万邦提供的报价单模板.xlsx"


def _default_craft_library_path() -> str:
    """优先仓库内模板（GitHub/Streamlit），否则回退本机历史路径。"""
    if os.path.isfile(_REPO_CRAFT_TEMPLATE):
        return _REPO_CRAFT_TEMPLATE
    if os.path.isfile(_LEGACY_CRAFT_TEMPLATE):
        return _LEGACY_CRAFT_TEMPLATE
    return _REPO_CRAFT_TEMPLATE


CRAFT_LIBRARY_PATH = os.getenv("CRAFT_LIBRARY_PATH", "").strip() or _default_craft_library_path()

# 询价解析历史（Streamlit「历史解析」Tab，仅 JSON，不含原文件二进制）
ENQUIRY_HISTORY_DIR = os.getenv(
    "ENQUIRY_HISTORY_DIR",
    os.path.join(_PROJECT_ROOT, "data", "enquiry_runs"),
)
ENQUIRY_HISTORY_MAX_RUNS = int(os.getenv("ENQUIRY_HISTORY_MAX_RUNS", "200"))

# 报价单归档（「我的报价」Tab，含匹配快照与可选 Excel 副本）
QUOTATION_STORE_DIR = os.getenv(
    "QUOTATION_STORE_DIR",
    os.path.join(_PROJECT_ROOT, "data", "quotations"),
)
QUOTATION_STORE_MAX_RECORDS = int(os.getenv("QUOTATION_STORE_MAX_RECORDS", "500"))

# === 解析召回优先（宁可多不可少） ===
PARSE_CHUNK_MAX_CHARS = int(os.getenv("PARSE_CHUNK_MAX_CHARS", "12000"))
PARSE_CHUNK_FALLBACK_CHARS = int(os.getenv("PARSE_CHUNK_FALLBACK_CHARS", "6000"))
PARSE_MAX_TOKENS = int(os.getenv("PARSE_MAX_TOKENS", "12000"))
# 单块疑似截断或条目过少时，按更小子块重试
PARSE_MIN_ITEMS_PER_CHUNK_HEURISTIC = int(os.getenv("PARSE_MIN_ITEMS_PER_CHUNK_HEURISTIC", "1"))
