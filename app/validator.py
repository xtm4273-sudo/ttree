"""
启动配置校验模块
在系统启动时检查关键配置是否就绪
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config


def validate_config() -> list[str]:
    """
    校验启动配置，返回错误列表。

    Returns:
        list[str]: 错误信息列表，空列表表示配置正常
    """
    errors = []

    # 检查 API Key
    if not config.LLM_API_KEY:
        errors.append("LLM_API_KEY 未设置，请在 .env 文件或环境变量中配置")

    # 检查工艺库路径
    craft_path = config.CRAFT_LIBRARY_PATH
    if not craft_path or not os.path.exists(craft_path):
        errors.append(f"工艺库文件不存在: {craft_path}")

    # 检查 FAISS 索引文件
    index_path = os.path.join("data", "craft_library.index")
    data_path = os.path.join("data", "craft_library.json")
    if not os.path.exists(index_path):
        errors.append(f"FAISS索引未构建: {index_path} (请点击侧边栏「构建索引」)")
    if not os.path.exists(data_path):
        errors.append(f"工艺库数据缺失: {data_path}")

    return errors
