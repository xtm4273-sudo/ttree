"""
Excel报价单生成模块
将匹配结果按万邦模板格式输出Excel
"""
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config


# 样式
HEADER_FONT = Font(bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
CAT_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

FILL_HIGH = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
FILL_MED = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
FILL_LOW = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FILL_NONE = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
FILL_NEW = PatternFill(start_color="E2BFFF", end_color="E2BFFF", fill_type="solid")  # 紫色-新增项


def generate_quotation_excel(match_results: list[dict], output_path: str) -> str:
    """
    生成报价单Excel。
    """
    wb = Workbook()
    _build_quotation_sheet(wb, match_results)
    _build_review_sheet(wb, match_results)
    _build_new_items_sheet(wb, match_results)
    _build_summary_sheet(wb, match_results)
    wb.save(output_path)
    return output_path


def _build_quotation_sheet(wb: Workbook, results: list[dict]):
    """主报价表"""
    ws = wb.active
    ws.title = "Quotation"

    headers = [
        ("ITEM No.", 14),
        ("ENQUIRY TITLE", 40),
        ("MATCHED CRAFT / SUGGESTED", 45),
        ("UNIT", 8),
        ("UNIT PRICE", 12),
        ("QTY", 7),
        ("TOTAL", 12),
        ("CONF%", 7),
        ("STATUS", 12),
        ("MATCH REASON", 35),
    ]

    for col, (name, width) in enumerate(headers, 1):
        cell = ws.cell(1, col, name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"

    row = 2
    for result in results:
        best = result.get("best_match")
        conf = result.get("confidence", 0)
        is_new = result.get("is_new_item", False)
        suggested = result.get("suggested_entry")

        # 选择填充色
        if is_new:
            fill = FILL_NEW
        else:
            fill = _get_confidence_fill(conf)

        # A: SFI编码
        sfi = result.get("enquiry_sfi") or ""
        ws.cell(row, 1, sfi).border = THIN_BORDER

        # B: 询价标题
        ws.cell(row, 2, result["enquiry_title"]).border = THIN_BORDER

        # C: 匹配工艺或建议条目
        if is_new and suggested:
            craft_text = suggested.get("title", result["enquiry_title"])
        elif best:
            craft_text = best["craft_title"]
        else:
            craft_text = "(未匹配)"
        ws.cell(row, 3, craft_text).border = THIN_BORDER

        # D: 单位
        if is_new and suggested:
            unit = suggested.get("unit", "")
        elif best:
            unit = best["unit"]
        else:
            unit = result.get("unit", "")
        ws.cell(row, 4, unit).border = THIN_BORDER

        # E: 单价（留空，客户真实库中有）
        ws.cell(row, 5, "").border = THIN_BORDER

        # F: 数量
        qty = result.get("quantity")
        ws.cell(row, 6, qty if qty else "").border = THIN_BORDER

        # G: 总价
        ws.cell(row, 7, "").border = THIN_BORDER

        # H: 置信度
        conf_cell = ws.cell(row, 8, conf if not is_new else "NEW")
        conf_cell.border = THIN_BORDER
        conf_cell.alignment = Alignment(horizontal="center")
        if conf < 60 and not is_new:
            conf_cell.font = Font(bold=True, color="CC0000")

        # I: 状态
        if is_new:
            status = "NEW - 待确认"
        elif conf >= 80:
            status = "AUTO"
        elif conf >= 60:
            status = "建议复核"
        else:
            status = "需确认"
        ws.cell(row, 9, status).border = THIN_BORDER

        # J: 匹配原因
        if is_new and suggested:
            reason = f"工艺库无此项，AI建议: {suggested.get('description', '')[:50]}"
        elif best:
            reason = best.get("llm_reason", "")
        else:
            reason = ""
        ws.cell(row, 10, reason).border = THIN_BORDER

        # 整行背景色
        for c in range(1, len(headers) + 1):
            ws.cell(row, c).fill = fill

        row += 1

    # 图例
    row += 2
    ws.cell(row, 1, "图例:").font = Font(bold=True, size=9)
    row += 1
    legends = [
        (">=80 高置信，可直接使用", FILL_HIGH),
        ("60-79 中等，建议复核", FILL_MED),
        ("40-59 低，需人工确认", FILL_LOW),
        ("<40 无匹配/不可信", FILL_NONE),
        ("NEW 工艺库新增项（AI建议）", FILL_NEW),
    ]
    for text, fill in legends:
        ws.cell(row, 1, text).font = Font(size=9)
        ws.cell(row, 1).fill = fill
        row += 1


def _build_review_sheet(wb: Workbook, results: list[dict]):
    """需审核条目"""
    ws = wb.create_sheet("需审核")

    headers = ["SFI", "询价标题", "最佳匹配", "置信度", "原因", "建议操作"]
    for col, name in enumerate(headers, 1):
        cell = ws.cell(1, col, name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 30
    ws.column_dimensions["F"].width = 20

    row = 2
    for result in results:
        conf = result.get("confidence", 0)
        is_new = result.get("is_new_item", False)
        if conf >= config.CONFIDENCE_LEVELS["high"] and not is_new:
            continue

        best = result.get("best_match")
        ws.cell(row, 1, result.get("enquiry_sfi") or "")
        ws.cell(row, 2, result["enquiry_title"])
        ws.cell(row, 3, best["craft_title"] if best else "(无)")
        ws.cell(row, 4, conf)
        ws.cell(row, 5, best["llm_reason"] if best else "无匹配候选")

        if is_new:
            action = "新增项：确认AI建议或手动定义"
        elif conf < config.CONFIDENCE_LEVELS["low"]:
            action = "匹配可能有误，请手动选择"
        else:
            action = "建议复核匹配是否正确"
        ws.cell(row, 6, action)
        row += 1


def _build_new_items_sheet(wb: Workbook, results: list[dict]):
    """工艺库新增项建议"""
    ws = wb.create_sheet("新增项建议")

    headers = ["询价标题", "AI建议标题", "AI建议描述", "建议单位", "操作"]
    for col, name in enumerate(headers, 1):
        cell = ws.cell(1, col, name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 25

    row = 2
    for result in results:
        if not result.get("is_new_item"):
            continue

        suggested = result.get("suggested_entry", {})
        ws.cell(row, 1, result["enquiry_title"])
        ws.cell(row, 2, suggested.get("title", ""))
        ws.cell(row, 3, suggested.get("description", ""))
        ws.cell(row, 4, suggested.get("unit", ""))
        ws.cell(row, 5, "确认后加入工艺库")
        row += 1

    if row == 2:
        ws.cell(2, 1, "（无新增项）")


def _build_summary_sheet(wb: Workbook, results: list[dict]):
    """统计汇总"""
    ws = wb.create_sheet("统计")

    total = len(results)
    high = sum(1 for r in results if r.get("confidence", 0) >= 80 and not r.get("is_new_item"))
    med = sum(1 for r in results if 60 <= r.get("confidence", 0) < 80 and not r.get("is_new_item"))
    low = sum(1 for r in results if 40 <= r.get("confidence", 0) < 80 and not r.get("is_new_item"))
    new_items = sum(1 for r in results if r.get("is_new_item"))
    no_match = sum(1 for r in results if r.get("confidence", 0) < 40 and not r.get("is_new_item"))

    data = [
        ["指标", "数量", "占比"],
        ["总条目数", total, "100%"],
        ["高置信 (>=80)", high, f"{high*100//max(total,1)}%"],
        ["中等 (60-79)", med, f"{med*100//max(total,1)}%"],
        ["低 (40-59)", low, f"{low*100//max(total,1)}%"],
        ["新增项（库中无）", new_items, f"{new_items*100//max(total,1)}%"],
        [],
        ["自动可用率", f"{high*100//max(total,1)}%", "无需人工"],
        ["需人工处理", f"{(total-high)*100//max(total,1)}%", ""],
    ]

    for r, row_data in enumerate(data, 1):
        for c, val in enumerate(row_data, 1):
            cell = ws.cell(r, c, val)
            if r == 1:
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 15


def _get_confidence_fill(confidence: int) -> PatternFill:
    if confidence >= config.CONFIDENCE_LEVELS["high"]:
        return FILL_HIGH
    elif confidence >= config.CONFIDENCE_LEVELS["medium"]:
        return FILL_MED
    elif confidence >= config.CONFIDENCE_LEVELS["low"]:
        return FILL_LOW
    else:
        return FILL_NONE
