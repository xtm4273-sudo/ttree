"""
配置文件 - API密钥和模型参数
"""
import os

# === LLM API 配置 ===
LLM_API_KEY = os.getenv("LLM_API_KEY", "sk-c4c68e444f5f43fe918759f998d35884")
LLM_BASE_URL = os.getenv("LLM_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")
LLM_MODEL = os.getenv("LLM_MODEL", "qwen-plus")

# === Embedding 配置 ===
EMBED_API_KEY = os.getenv("EMBED_API_KEY", "") or LLM_API_KEY
EMBED_BASE_URL = os.getenv("EMBED_BASE_URL", "") or LLM_BASE_URL
EMBED_MODEL = os.getenv("EMBED_MODEL", "text-embedding-v3")

# === 匹配参数 ===
TOP_K = 10  # 向量检索返回候选数

CONFIDENCE_WEIGHTS = {
    "vector_similarity": 0.25,  # 向量相似度权重
    "llm_score": 0.55,          # LLM打分权重
    "sfi_match": 0.20,          # SFI编码匹配权重（无编码时自动归零）
}

# === 置信度分级阈值 ===
CONFIDENCE_LEVELS = {
    "high": 80,      # >=80 高置信，自动可用
    "medium": 60,    # 60-79 中等，建议复核
    "low": 40,       # 40-59 低，需人工确认
    # <40 无匹配/不可信，触发AI生成建议条目
}

# === 文件路径 ===
CRAFT_LIBRARY_PATH = os.getenv(
    "CRAFT_LIBRARY_PATH",
    r"C:\Users\24895\Desktop\万邦船舶询价\万邦提供的报价单模板.xlsx"
)
