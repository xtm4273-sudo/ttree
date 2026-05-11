"""
万邦船舶询价系统 - 测试评估脚本

功能：
1. 跑完整流程（解析->匹配->出结果）
2. 分段计时，定位瓶颈
3. 逐条展示匹配结果，让用户评判正确/错误
4. 自动计算准确率、置信度校准等指标
5. 保存测试报告，方便后续对比

使用方法：
  python test_evaluation.py               # 完整模式：跑流程+逐条评判
  python test_evaluation.py --quick        # 快速模式：只跑流程+统计，不评判
  python test_evaluation.py --report       # 只看上一次的测试报告
  python test_evaluation.py --review       # 加载上次结果再审核
"""
import sys
import os
import time
import json
import argparse
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.document_parser import parse_document
from app.craft_library import load_vector_index
from app.match_engine import match_all_items
from app.excel_generator import generate_quotation_excel
import config

# === 配置 ===
TEST_PDF = "test_enquiry_50pages.pdf"
REPORT_FILE = "test_report.json"
HISTORY_FILE = "test_history.json"


# ============================================================
# 工具：用纯ASCII画表格边框（兼容GBK终端）
# ============================================================
def t_border(width=50):
    return "+" + "-" * (width - 2) + "+"

def t_header(text, width=50):
    return "| " + text.ljust(width - 4) + " |"

def t_row(key, val, width=50):
    line = f"| {key}: {val}".ljust(width - 1) + "|"
    return line


# ============================================================
# 第一步：跑流程
# ============================================================
def run_pipeline():
    """执行完整流程，返回匹配结果和时间统计"""
    timings = {}

    bar = "=" * 60
    print(bar)
    print("  万邦船舶询价系统 - 测试评估")
    print(f"  时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(bar)

    # 1. 加载索引
    print("\n[1/4] 加载FAISS索引...")
    t0 = time.time()
    index, craft_items = load_vector_index("data")
    if index is None:
        print("  [失败] 索引未构建，请先在Streamlit界面构建索引")
        return None, None
    t_load = time.time() - t0
    timings["load_index"] = round(t_load, 2)
    print(f"  [OK] 工艺库: {len(craft_items)} 条, 维度: {index.d}, 耗时: {t_load:.2f}s")

    # 2. 解析文档
    print(f"\n[2/4] 解析文档: {TEST_PDF}")
    t0 = time.time()

    def doc_log(msg):
        print(f"  {msg}")

    try:
        enquiry_items = parse_document(TEST_PDF, log_callback=doc_log)
    except Exception as e:
        print(f"  [失败] 解析失败: {e}")
        import traceback
        traceback.print_exc()
        return None, None

    t_parse = time.time() - t0
    timings["parse_document"] = round(t_parse, 2)
    print(f"\n  [OK] 解析完成: {len(enquiry_items)} 条, 耗时: {t_parse:.1f}s")

    # 3. 匹配
    print(f"\n[3/4] AI匹配 ({len(enquiry_items)} 条)...")
    t0 = time.time()

    phase_times = {"phase1": 0, "phase2": 0}

    def match_progress(cur, total):
        pass

    def match_log(msg):
        if "Phase 1" in msg:
            print(f"  {msg}")
        elif "Phase 2" in msg:
            print(f"  {msg}")
        elif "向量检索完成" in msg:
            elapsed = time.time() - t0
            phase_times["phase1"] = round(elapsed, 1)
            print(f"    {msg}")
        elif "并发精排完成" in msg:
            elapsed = time.time() - t0
            phase_times["phase2"] = round(elapsed - phase_times["phase1"], 1)
            print(f"    {msg}")
        elif "完成" in msg:
            print(f"  {msg}")

    results = match_all_items(
        enquiry_items, index, craft_items,
        progress_callback=match_progress,
        log_callback=match_log,
    )

    t_match = time.time() - t0
    timings["match_total"] = round(t_match, 1)
    timings["phase1_vector_search"] = phase_times["phase1"]
    timings["phase2_llm_rerank"] = phase_times["phase2"]

    print(f"  [OK] 匹配完成, 总耗时: {t_match:.1f}s")
    print(f"    Phase 1 (向量检索): {phase_times['phase1']}s")
    print(f"    Phase 2 (LLM精排):  {phase_times['phase2']}s")

    # 4. 生成Excel
    print(f"\n[4/4] 生成报价单...")
    t0 = time.time()
    output_name = f"test_output_{datetime.now().strftime('%m%d_%H%M')}.xlsx"
    generate_quotation_excel(results, output_name)
    t_excel = time.time() - t0
    timings["generate_excel"] = round(t_excel, 2)
    timings["total"] = round(sum(timings.values()), 1)
    print(f"  [OK] 报价单已保存: {output_name}, 耗时: {t_excel:.1f}s")

    return results, timings


# ============================================================
# 第二步：算统计指标
# ============================================================
def compute_statistics(results):
    """计算匹配结果的统计指标"""
    total = len(results)
    if total == 0:
        return {}

    high = sum(1 for r in results if r["confidence"] >= 80 and not r.get("is_new_item"))
    med = sum(1 for r in results if 60 <= r["confidence"] < 80 and not r.get("is_new_item"))
    low = sum(1 for r in results if 40 <= r["confidence"] < 60 and not r.get("is_new_item"))
    no_match = sum(1 for r in results if r.get("is_new_item"))

    stats = {
        "total_items": total,
        "high_confidence_auto": high,
        "medium_review": med,
        "low_confirm": low,
        "new_item_suggested": no_match,
        "auto_available_rate": round(high * 100 / max(total, 1), 1),
        "need_manual_rate": round((total - high) * 100 / max(total, 1), 1),
    }

    # 置信度分布分析
    confs = [r["confidence"] for r in results if not r.get("is_new_item")]
    if confs:
        stats["avg_confidence"] = round(sum(confs) / len(confs), 1)
        stats["median_confidence"] = round(sorted(confs)[len(confs) // 2], 1)
        stats["min_confidence"] = min(confs)
        stats["max_confidence"] = max(confs)

    # 如果串行执行，每条约2.5秒
    stats["estimated_serial_time"] = round(len(confs) * 2.5, 1) if confs else 0

    # 置信度分布
    stats["confidence_distribution"] = {}
    for r in results:
        if not r.get("is_new_item"):
            c = r["confidence"]
            band = f"{(c//10)*10}-{(c//10)*10+9}"
            stats["confidence_distribution"][band] = stats["confidence_distribution"].get(band, 0) + 1

    return stats


# ============================================================
# 第三步：逐条审核（人工打标签）
# ============================================================
def review_results(results):
    """
    逐条展示匹配结果，让用户评判。

    评分标准：
      y = 正确（匹配正确，可以直接用）
      p = 部分正确（同类但不够精确，需要调整）
      n = 错误（匹配错了）
      s = 跳过（不确定或不想评）
    """
    print("\n" + "=" * 60)
    print("  逐条审核模式")
    print("  评分标准: y=正确  p=部分正确  n=错误  s=跳过  q=退出")
    print("=" * 60)

    reviews = []
    total = len(results)

    for i, r in enumerate(results):
        is_new = r.get("is_new_item", False)
        conf = r["confidence"]
        title = r["enquiry_title"][:60]
        best = r.get("best_match")
        suggested = r.get("suggested_entry")

        dash = "-" * 50
        print(f"\n{dash}")
        print(f"[{i+1}/{total}] 询价: {title}")
        if r.get("enquiry_sfi"):
            print(f"       SFI: {r['enquiry_sfi']}")

        # 显示匹配结果
        if is_new:
            sug_title = suggested.get("title", "")[:50] if suggested else "(无)"
            print(f"       [NEW] 新增项 (<40分, 工艺库无匹配)")
            print(f"       AI建议标题: {sug_title}")
        elif best:
            print(f"       匹配工艺: {best['craft_title'][:55]}")
            print(f"       置信度: {conf}分 | 向量分={best['vector_score']} AI分={best['llm_score']} SFI分={best['sfi_score']}")
            print(f"       原因: {best.get('llm_reason', '')[:80]}")
        else:
            print(f"       无匹配")

        # 等用户输入
        while True:
            try:
                tag = input(f"       评分 [y/p/n/s/q]: ").strip().lower()
                if tag in ("y", "p", "n", "s", "q"):
                    break
                print("       请输入 y, p, n, s 或 q")
            except (EOFError, KeyboardInterrupt):
                tag = "q"
                break

        if tag == "q":
            break

        if tag != "s":
            reviews.append({
                "index": i,
                "enquiry_title": r["enquiry_title"],
                "confidence": conf,
                "is_new_item": is_new,
                "matched_title": best["craft_title"] if best else (suggested.get("title", "") if suggested else ""),
                "human_score": tag,
            })
        else:
            reviews.append({
                "index": i,
                "enquiry_title": r["enquiry_title"],
                "confidence": conf,
                "is_new_item": is_new,
                "human_score": "skipped",
            })

    reviewed = len([r for r in reviews if r['human_score'] != 'skipped'])
    print(f"\n  [OK] 已审核 {reviewed} 条")
    return reviews


# ============================================================
# 第四步：算精准度指标
# ============================================================
def compute_accuracy(reviews, stats):
    """根据人工评分计算精准度指标"""
    rated = [r for r in reviews if r['human_score'] in ("y", "p", "n")]
    if not rated:
        print("\n  [提示] 没有评级记录，无法计算准确率")
        return stats

    correct = sum(1 for r in rated if r['human_score'] == "y")
    partial = sum(1 for r in rated if r['human_score'] == "p")
    wrong = sum(1 for r in rated if r['human_score'] == "n")

    stats["human_reviewed_count"] = len(rated)
    stats["human_correct"] = correct
    stats["human_partial"] = partial
    stats["human_wrong"] = wrong
    stats["accuracy_strict"] = round(correct * 100 / len(rated), 1)   # 严格：只有y算对
    stats["accuracy_loose"] = round((correct + partial) * 100 / len(rated), 1)  # 宽松：y+p算对

    # 置信度校准分析
    print(f"\n  [置信度校准] 各分数段的实际准确率:")
    print(f"  {'分数段':<14} {'数量':>4} {'准确率':>8} {'平均分':>8}")
    print(f"  {'-'*40}")
    bands = [(80, 100, "高置信(80-100)"), (60, 79, "中置信(60-79)"), (0, 59, "低置信(0-59)")]
    for lo, hi, label in bands:
        band = [r for r in rated if lo <= r["confidence"] <= hi and not r["is_new_item"]]
        if band:
            band_correct = sum(1 for r in band if r["human_score"] == "y")
            avg_conf = sum(r["confidence"] for r in band) / len(band)
            acc = band_correct * 100 / len(band)
            print(f"  {label:<14} {len(band):>4} {acc:>7.1f}% {avg_conf:>8.1f}")

    # 也给新增项（NEW）算一个准确率
    new_band = [r for r in rated if r.get("is_new_item")]
    if new_band:
        new_acc = sum(1 for r in new_band if r["human_score"] == "y") * 100 / len(new_band)
        print(f"  {'新增项(NEW)':<14} {len(new_band):>4} {new_acc:>7.1f}% {'N/A':>8}")

    return stats


# ============================================================
# 报告输出
# ============================================================
def print_report(stats, timings, results, reviews):
    """打印完整测试报告"""
    W = 52  # 表格宽度

    print("\n\n")
    print("=" * 60)
    print("  测 试 报 告")
    print("=" * 60)
    print(f"  测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  测试文件: {TEST_PDF}")

    # 速度指标
    print(f"\n  -------- 速度指标 --------")
    if timings:
        print(f"  索引加载         {timings.get('load_index', 0):>8.2f}s")
        print(f"  文档解析         {timings.get('parse_document', 0):>8.2f}s")
        print(f"  Phase1 向量检索  {timings.get('phase1_vector_search', 0):>8.1f}s")
        print(f"  Phase2 LLM精排   {timings.get('phase2_llm_rerank', 0):>8.1f}s")
        print(f"  生成Excel        {timings.get('generate_excel', 0):>8.2f}s")
        print(f"  -----------------------------------")
        print(f"  总耗时           {timings.get('total', 0):>8.1f}s")

    # 匹配统计
    print(f"\n  -------- 匹配统计 --------")
    if stats:
        print(f"  总条目数:           {stats['total_items']}")
        print(f"  高置信 >=80 (自动): {stats['high_confidence_auto']}")
        print(f"  中等 60-79 (复核):  {stats['medium_review']}")
        print(f"  低 <60 (需确认):     {stats['low_confirm']}")
        print(f"  新增项 (紫色):       {stats['new_item_suggested']}")
        print(f"  --------------------------------")
        print(f"  自动可用率:          {stats.get('auto_available_rate', 0)}%")
        print(f"  需人工处理:          {stats.get('need_manual_rate', 0)}%")
        if "avg_confidence" in stats:
            print(f"  平均置信度:          {stats['avg_confidence']}分")
            print(f"  中位置信度:          {stats['median_confidence']}分")
            print(f"  置信度范围:          {stats['min_confidence']}-{stats['max_confidence']}分")

    # 置信度分布（详细）
    if "confidence_distribution" in stats:
        print(f"\n  -------- 置信度分布 --------")
        for band in sorted(stats["confidence_distribution"].keys()):
            count = stats["confidence_distribution"][band]
            bar = "#" * count
            print(f"  {band}分: {count:>3}  {bar}")

    # 精准度指标
    if "human_reviewed_count" in stats and stats["human_reviewed_count"] > 0:
        print(f"\n  -------- 精准度指标 --------")
        print(f"  人工审核条目:       {stats['human_reviewed_count']}")
        print(f"  正确:               {stats['human_correct']}")
        print(f"  部分正确:           {stats['human_partial']}")
        print(f"  错误:               {stats['human_wrong']}")
        print(f"  --------------------------------")
        print(f"  严格准确率(y):      {stats.get('accuracy_strict', 0)}%")
        print(f"  宽松准确率(y+p):    {stats.get('accuracy_loose', 0)}%")

    # 性能评估
    print(f"\n  -------- 性能评估 --------")
    if timings and stats:
        phase2 = timings.get("phase2_llm_rerank", 0)
        total_items = stats["total_items"]
        if phase2 > 0 and total_items > 0:
            avg_per_item = phase2 / total_items
            print(f"  单条LLM精排平均:    {avg_per_item:.2f}s/条")
            if stats.get("estimated_serial_time"):
                saved = stats["estimated_serial_time"] - timings.get("match_total", 0)
                print(f"  并发优化节省:       约 {saved:.0f}s")
                print(f"  如果串行执行估:     约 {stats['estimated_serial_time']}s")

    # 优化建议
    print(f"\n  -------- 优化建议 --------")
    suggestions = []
    if timings and timings.get("phase2_llm_rerank", 0) > 30:
        suggestions.append("  * LLM精排耗时过长，建议增大并发数(当前8)")
    if stats and stats.get("auto_available_rate", 0) < 30:
        suggestions.append("  * 自动可用率偏低，检查Prompt和权重配置")
    if stats and stats.get("new_item_suggested", 0) > stats["total_items"] * 0.3:
        suggestions.append("  * 新增项比例过高，工艺库可能缺失关键条目")
    if not suggestions:
        suggestions.append("  * 暂无显著优化建议")
    for s in suggestions:
        print(s)

    print(f"\n  Excel报价单已生成，可打开查看详情")
    print(f"  测试报告已保存: {REPORT_FILE}")
    print("=" * 60)


# ============================================================
# 主入口
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="万邦船舶询价系统测试评估")
    parser.add_argument("--quick", action="store_true", help="快速模式：只跑流程+统计，不逐条审核")
    parser.add_argument("--report", action="store_true", help="只看上一次的测试报告")
    parser.add_argument("--review", action="store_true", help="对已有结果进行审核（不重新跑流程）")
    args = parser.parse_args()

    # 查看历史报告
    if args.report:
        if os.path.exists(REPORT_FILE):
            with open(REPORT_FILE, "r", encoding="utf-8") as f:
                report = json.load(f)
            print_report(
                report.get("stats", {}),
                report.get("timings", {}),
                None,
                report.get("reviews", []),
            )
        else:
            print("[提示] 没有找到历史报告，请先运行测试")
        return

    # 加载已有结果进行审核
    if args.review:
        if os.path.exists("_last_results.json"):
            with open("_last_results.json", "r", encoding="utf-8") as f:
                saved = json.load(f)
            results = saved["results"]
            timings = saved["timings"]
            print(f"已加载之前的结果: {len(results)} 条")
        else:
            print("[提示] 没有找到之前的结果，请先运行完整测试")
            return
    else:
        # 跑完整流程
        results, timings = run_pipeline()
        if results is None:
            return

        # 保存结果供后续审核
        with open("_last_results.json", "w", encoding="utf-8") as f:
            json.dump({
                "results": results,
                "timings": timings,
                "timestamp": datetime.now().isoformat(),
            }, f, ensure_ascii=False, indent=2)

    # 算统计指标
    stats = compute_statistics(results)

    # 如果不是快速模式，逐条审核
    reviews = []
    if not args.quick and not args.report:
        print(f"\n共 {len(results)} 条匹配结果，是否开始逐条审核？")
        resp = input("回车开始审核，输入 n 跳过: ").strip().lower()
        if resp != "n":
            reviews = review_results(results)
            stats = compute_accuracy(reviews, stats)

    # 打印报告
    print_report(stats, timings, results, reviews)

    # 保存报告
    report = {
        "timestamp": datetime.now().isoformat(),
        "stats": stats,
        "timings": timings,
        "reviews": reviews,
    }
    with open(REPORT_FILE, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    # 追加到历史记录
    history = []
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            try:
                history = json.load(f)
            except:
                pass
    history.append({
        "timestamp": report["timestamp"],
        "stats": stats,
        "timings": timings,
    })
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history[-10:], f, ensure_ascii=False, indent=2)

    print(f"\n历史记录已保存: {HISTORY_FILE}")


if __name__ == "__main__":
    main()
