"""
连续页抽样压测脚本：
1) 从大PDF中抽取连续N页（默认50页）
2) 跑完整流程（解析 -> 匹配 -> 生成Excel）
3) 输出耗时和决策路径统计
"""
import argparse
import json
import os
import random
import time
from datetime import datetime

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    from PyPDF2 import PdfReader, PdfWriter

from app.document_parser import parse_document
from app.craft_library import load_vector_index
from app.excel_generator import generate_quotation_excel
from app.match_engine import match_all_items


def sample_contiguous_pages(src_pdf: str, out_pdf: str, pages: int, seed: int):
    reader = PdfReader(src_pdf)
    total_pages = len(reader.pages)
    if total_pages == 0:
        raise ValueError("源PDF无页面")

    sample_pages = min(pages, total_pages)
    random.seed(seed)
    max_start = total_pages - sample_pages
    start = random.randint(0, max_start) if max_start > 0 else 0
    end = start + sample_pages - 1

    writer = PdfWriter()
    for i in range(start, end + 1):
        writer.add_page(reader.pages[i])

    with open(out_pdf, "wb") as f:
        writer.write(f)

    return {
        "total_pages": total_pages,
        "sample_pages": sample_pages,
        "start_page_1based": start + 1,
        "end_page_1based": end + 1,
    }


def main():
    parser = argparse.ArgumentParser(description="连续50页抽样端到端测试")
    parser.add_argument("--src-pdf", required=True, help="源PDF路径")
    parser.add_argument("--pages", type=int, default=50, help="连续抽样页数")
    parser.add_argument("--seed", type=int, default=20260512, help="随机种子（决定起始页）")
    parser.add_argument(
        "--out-dir",
        default=os.path.join("test_datasets", "benchmark_outputs"),
        help="输出目录（默认 test_datasets/benchmark_outputs）",
    )
    args = parser.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    sampled_pdf = os.path.join(args.out_dir, f"sample_contiguous_{args.pages}p_{ts}.pdf")
    out_excel = os.path.join(args.out_dir, f"quotation_sample_{args.pages}p_{ts}.xlsx")
    out_report = os.path.join(args.out_dir, f"report_sample_{args.pages}p_{ts}.json")

    total_t0 = time.time()
    sample_meta = sample_contiguous_pages(args.src_pdf, sampled_pdf, args.pages, args.seed)

    index, craft_items = load_vector_index("data")
    if index is None or not craft_items:
        raise RuntimeError("FAISS索引未加载，请先构建索引")

    parse_logs = []
    match_logs = []

    def parse_log(msg):
        parse_logs.append(msg)
        print(f"[PARSE] {msg}")

    def match_log(msg):
        match_logs.append(msg)
        print(f"[MATCH] {msg}")

    parse_t0 = time.time()
    enquiry_items = parse_document(sampled_pdf, log_callback=parse_log)
    parse_elapsed = time.time() - parse_t0
    parsed_count = len(enquiry_items)

    match_t0 = time.time()
    results = match_all_items(enquiry_items, index, craft_items, log_callback=match_log)
    match_elapsed = time.time() - match_t0
    exported_count = len(results)

    excel_t0 = time.time()
    generate_quotation_excel(results, out_excel)
    excel_elapsed = time.time() - excel_t0

    decision_counts = {}
    for r in results:
        p = r.get("decision_path", "unknown")
        decision_counts[p] = decision_counts.get(p, 0) + 1

    llm_called = sum(1 for r in results if r.get("llm_called"))
    total_items = len(results)
    high = sum(1 for r in results if r.get("confidence", 0) >= 80 and not r.get("is_new_item"))
    med = sum(1 for r in results if 60 <= r.get("confidence", 0) < 80 and not r.get("is_new_item"))
    low = sum(1 for r in results if 40 <= r.get("confidence", 0) < 60 and not r.get("is_new_item"))
    new_items = sum(1 for r in results if r.get("is_new_item"))

    total_elapsed = time.time() - total_t0
    report = {
        "source_pdf": args.src_pdf,
        "sample_pdf": sampled_pdf,
        "sample_meta": sample_meta,
        "output_excel": out_excel,
        "timings_sec": {
            "parse": round(parse_elapsed, 2),
            "match": round(match_elapsed, 2),
            "excel": round(excel_elapsed, 2),
            "total": round(total_elapsed, 2),
        },
        "integrity": {
            "parsed_items": parsed_count,
            "match_result_rows": exported_count,
            "rows_match": parsed_count == exported_count,
        },
        "result_stats": {
            "total_items": total_items,
            "high_confidence": high,
            "medium_confidence": med,
            "low_confidence": low,
            "new_items": new_items,
            "llm_called_items": llm_called,
            "llm_called_ratio": round((llm_called / total_items), 4) if total_items else 0,
            "decision_path_counts": decision_counts,
        },
    }

    with open(out_report, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print("\n=== BENCHMARK SUMMARY ===")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print(f"\nREPORT_FILE={out_report}")
    print(f"EXCEL_FILE={out_excel}")


if __name__ == "__main__":
    main()
