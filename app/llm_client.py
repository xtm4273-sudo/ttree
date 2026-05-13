"""
统一的OpenAI客户端管理
避免各模块重复创建连接
"""
from openai import OpenAI
import config

_llm_client = None
_llm_client_sig: tuple[str, str] | None = None
_embed_client = None
_embed_client_sig: tuple[str, str] | None = None


def get_llm_client():
    global _llm_client, _llm_client_sig
    sig = (config.LLM_API_KEY or "", config.LLM_BASE_URL or "")
    if _llm_client is None or _llm_client_sig != sig:
        _llm_client = OpenAI(
            api_key=config.LLM_API_KEY,
            base_url=config.LLM_BASE_URL,
            timeout=120,
            max_retries=2,
        )
        _llm_client_sig = sig
    return _llm_client


def get_embed_client():
    global _embed_client, _embed_client_sig
    sig = (config.EMBED_API_KEY or "", config.EMBED_BASE_URL or "")
    if _embed_client is None or _embed_client_sig != sig:
        _embed_client = OpenAI(
            api_key=config.EMBED_API_KEY,
            base_url=config.EMBED_BASE_URL,
            timeout=120,
            max_retries=2,
        )
        _embed_client_sig = sig
    return _embed_client
