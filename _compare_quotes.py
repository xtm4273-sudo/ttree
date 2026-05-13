# -*- coding: utf-8 -*-
"""Compare standard template vs generated quotation Excel."""
import pandas as pd
from pathlib import Path

path_std = Path(r"C:\Users\24895\Desktop\万邦船舶询价\万邦提供的报价单模板.xlsx")
path_gen = Path(r"C:\software\报价单_0511_1817.xlsx")
out = Path(r"C:\Users\24895\Desktop\ttree\_quote_diff_report.txt")

lines: list[str] = []


def log(msg: str) -> None:
    lines.append(msg)


def norm_cell(v):
    if pd.isna(v):
        return ""
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return str(round(v, 6)).rstrip("0").rstrip(".")
    return str(v).strip()


def df_to_matrix(df: pd.DataFrame) -> list[list[str]]:
    out_rows: list[list[str]] = []
    for _, row in df.iterrows():
        out_rows.append([norm_cell(x) for x in row.tolist()])
    return out_rows


def compare_sheets(name_std: str, name_gen: str, max_rows: int | None = None) -> None:
    df_std = pd.read_excel(path_std, sheet_name=name_std, header=None)
    df_gen = pd.read_excel(path_gen, sheet_name=name_gen, header=None)
    log(f"\n=== Sheet: 标准[{name_std}] vs 识别[{name_gen}] ===")
    log(f"标准: {df_std.shape[0]} 行 x {df_std.shape[1]} 列")
    log(f"识别: {df_gen.shape[0]} 行 x {df_gen.shape[1]} 列")

    r_max = max(df_std.shape[0], df_gen.shape[0])
    c_max = max(df_std.shape[1], df_gen.shape[1])
    if max_rows is not None:
        r_max = min(r_max, max_rows)

    diff_count = 0
    for r in range(r_max):
        row_diffs: list[str] = []
        for c in range(c_max):
            v_std = norm_cell(df_std.iloc[r, c]) if r < df_std.shape[0] and c < df_std.shape[1] else ""
            v_gen = norm_cell(df_gen.iloc[r, c]) if r < df_gen.shape[0] and c < df_gen.shape[1] else ""
            if v_std != v_gen:
                col_letter = chr(65 + c) if c < 26 else f"C{c}"
                row_diffs.append(f"  列{c+1}({col_letter}): 标准={v_std!r} | 识别={v_gen!r}")
                diff_count += 1
        if row_diffs:
            log(f"--- 第 {r+1} 行 (Excel行号 {r+1}) 不一致 ---")
            log("\n".join(row_diffs))

    if diff_count == 0:
        log("(本页范围内无单元格差异)")
    else:
        log(f"\n本 Sheet 不一致单元格数: {diff_count}")


def main():
    xl_std = pd.ExcelFile(path_std)
    xl_gen = pd.ExcelFile(path_gen)
    log("===== 文件概览 =====")
    log(f"标准模板: {path_std}")
    log(f"识别结果: {path_gen}")
    log(f"标准 Sheet 列表: {xl_std.sheet_names}")
    log(f"识别 Sheet 列表: {xl_gen.sheet_names}")

    std_set = set(xl_std.sheet_names)
    gen_set = set(xl_gen.sheet_names)
    only_std = std_set - gen_set
    only_gen = gen_set - std_set
    log("\n===== Sheet 结构差异 =====")
    if only_std:
        log(f"仅在标准模板中存在: {sorted(only_std)}")
    if only_gen:
        log(f"仅在识别结果中存在: {sorted(only_gen)}")
    common = sorted(std_set & gen_set)
    log(f"两边同名的 Sheet: {common}")

    gen_main = None
    if "Quotation" in gen_set:
        gen_main = "Quotation"
    elif "报价单" in gen_set:
        gen_main = "报价单"
    if gen_main and "Quotation" in std_set:
        compare_sheets("Quotation", gen_main)
    else:
        log("\n(无法对比主表：标准需含「Quotation」，识别结果需含「Quotation」或「报价单」)")

    # 尝试对比识别文件中的其他 sheet 与标准（若有同名）
    skip_gen = {"Quotation", "报价单"}
    for s in xl_gen.sheet_names:
        if s not in skip_gen and s in std_set:
            compare_sheets(s, s)

    out.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
