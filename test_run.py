"""
MVP验证脚本 - 跑完整流程: 解析PDF → 向量检索 → LLM精排 → 生成Excel
"""
import time
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.document_parser import parse_document
from app.craft_library import load_vector_index
from app.match_engine import match_all_items
from app.excel_generator import generate_quotation_excel

PDF_PATH = "test_enquiry_50pages.pdf"
OUTPUT_EXCEL = "test_output_quotation.xlsx"


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")


def main():
    total_start = time.time()

    # 1. 加载向量索引
    log("加载FAISS索引...")
    index, craft_items = load_vector_index("data")
    if index is None:
        log("❌ 索引未构建，退出")
        return
    log(f"工艺库: {len(craft_items)} 条, 索引维度: {index.d}")

    # 2. 解析文档
    log(f"开始解析: {PDF_PATH}")
    t0 = time.time()
    items = parse_document(PDF_PATH, log_callback=log)
    log(f"解析完成: {len(items)} 条, 耗时 {time.time()-t0:.1f}s")

    # 打印前5条看看
    log("--- 解析结果预览(前5条) ---")
    for i, item in enumerate(items[:5]):
        sfi = item.get("sfi_code") or "无"
        title = item["title"][:60]
        log(f"  [{i+1}] SFI={sfi} | {title}")

    # 3. 匹配
    log(f"开始匹配 {len(items)} 条...")
    t0 = time.time()

    def progress(cur, total):
        if cur % 5 == 0 or cur == total:
            elapsed = time.time() - t0
            eta = elapsed / cur * (total - cur) if cur > 0 else 0
            log(f"  进度: {cur}/{total} ({cur*100//total}%) ETA={eta:.0f}s")

    results = match_all_items(items, index, craft_items,
                              progress_callback=progress, log_callback=log)
    t_match = time.time() - t0
    log(f"匹配完成, 耗时 {t_match:.1f}s")

    # 4. 统计
    total = len(results)
    high = sum(1 for r in results if r["confidence"] >= 80 and not r.get("is_new_item"))
    med = sum(1 for r in results if 60 <= r["confidence"] < 80 and not r.get("is_new_item"))
    low = sum(1 for r in results if 40 <= r["confidence"] < 60 and not r.get("is_new_item"))
    new_items = sum(1 for r in results if r.get("is_new_item"))
    pending = sum(1 for r in results if r.get("review_status") == "PENDING_REVIEW")

    log("=== 匹配统计 ===")
    log(f"  总条目: {total}")
    log(f"  高置信(>=80): {high} ({high*100//max(total,1)}%)")
    log(f"  中等(60-79):  {med} ({med*100//max(total,1)}%)")
    log(f"  低(40-59):    {low} ({low*100//max(total,1)}%)")
    log(f"  库无向量候选: {new_items} ({new_items*100//max(total,1)}%)")
    log(f"  待人工(PENDING): {pending} ({pending*100//max(total,1)}%)")

    # 打印部分匹配结果
    log("=== 匹配结果示例 ===")
    for r in results[:8]:
        best = r.get("best_match")
        is_new = r.get("is_new_item")
        enq = r["enquiry_title"][:40]
        if is_new:
            sug = r.get("suggested_entry", {}).get("title", "")[:40]
            log(f"  [NEW] {enq} → AI建议: {sug}")
        elif best:
            craft = best["craft_title"][:40]
            log(f"  [{r['confidence']:3d}] {enq} → {craft}")
        else:
            log(f"  [  0] {enq} → 无匹配")

    # 5. 生成Excel
    log(f"生成报价单: {OUTPUT_EXCEL}")
    generate_quotation_excel(results, OUTPUT_EXCEL)
    log(f"报价单已保存: {OUTPUT_EXCEL}")

    total_time = time.time() - total_start
    log(f"=== 全流程完成, 总耗时 {total_time:.1f}s ===")


if __name__ == "__main__":
    main()
