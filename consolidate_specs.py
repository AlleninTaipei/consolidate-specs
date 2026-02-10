#!/usr/bin/env python3
"""
consolidate_specs.py – 將多個機種規格 Excel 檔案彙整為一份 Excel。

Usage:
    python consolidate_specs.py ./spec_files
    python consolidate_specs.py ./spec_files -o output.xlsx
    python consolidate_specs.py ./spec_files --force
"""

import sys
import os
import glob
import argparse
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False


# ---------------------------------------------------------------------------
# 2a. parse_spec_file
# ---------------------------------------------------------------------------
def _parse_rows(rows_iter):
    """共用解析邏輯：從 (b_val, c_val) 迭代器建構 OrderedDict。"""
    categories = OrderedDict()
    current_cat = None
    for b_val, c_val in rows_iter:
        b_str = str(b_val).strip() if b_val is not None else ""
        c_str = str(c_val).strip() if c_val is not None else ""

        if b_str:
            current_cat = b_str
            if current_cat not in categories:
                categories[current_cat] = []
            if c_str:
                categories[current_cat].append(c_str)
        elif c_str and current_cat:
            categories[current_cat].append(c_str)
    return categories


def _parse_xlsx(filepath):
    """解析 .xlsx 檔案。"""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)

    sheet = None
    for name in ("Specification", "Specifications"):
        if name in wb.sheetnames:
            sheet = wb[name]
            break
    if sheet is None:
        wb.close()
        raise ValueError('找不到 "Specification" 工作表')

    model_name = sheet["B1"].value
    if not model_name:
        wb.close()
        raise ValueError("B1 無機種名稱")
    model_name = str(model_name).strip()

    def row_iter():
        for row in sheet.iter_rows(min_row=2, max_col=3, values_only=True):
            b = row[1] if len(row) > 1 else None
            c = row[2] if len(row) > 2 else None
            yield b, c

    categories = _parse_rows(row_iter())
    wb.close()
    return model_name, categories


def _parse_xls(filepath):
    """解析 .xls 檔案（需要 xlrd）。"""
    if not HAS_XLRD:
        raise ValueError("需要安裝 xlrd 才能讀取 .xls 檔案 (pip install xlrd)")
    wb = xlrd.open_workbook(filepath)

    sheet = None
    for name in ("Specification", "Specifications"):
        if name in wb.sheet_names():
            sheet = wb.sheet_by_name(name)
            break
    if sheet is None:
        raise ValueError('找不到 "Specification" 工作表')

    model_name = sheet.cell_value(0, 1) if sheet.ncols > 1 else None
    if not model_name:
        raise ValueError("B1 無機種名稱")
    model_name = str(model_name).strip()

    def row_iter():
        for r in range(1, sheet.nrows):
            b = sheet.cell_value(r, 1) if sheet.ncols > 1 else None
            c = sheet.cell_value(r, 2) if sheet.ncols > 2 else None
            # xlrd 回傳空字串而非 None
            yield (b if b != "" else None), (c if c != "" else None)

    categories = _parse_rows(row_iter())
    return model_name, categories


def parse_spec_file(filepath):
    """解析單一 Excel 檔案，回傳 (model_name, OrderedDict{類別名: [描述行]})。"""
    if filepath.lower().endswith(".xls") and not filepath.lower().endswith(".xlsx"):
        return _parse_xls(filepath)
    return _parse_xlsx(filepath)


# ---------------------------------------------------------------------------
# 2b. merge_category_orders
# ---------------------------------------------------------------------------
def merge_category_orders(all_data):
    """依各檔案出現順序合併類別，保持相對順序一致。"""
    seen = set()
    order = []
    for _model, cats in all_data:
        for cat in cats:
            if cat not in seen:
                seen.add(cat)
                order.append(cat)
    return order


# ---------------------------------------------------------------------------
# 2c. compute_max_rows
# ---------------------------------------------------------------------------
def compute_max_rows(all_data, category_order):
    """計算每個類別跨所有機種的最大描述行數（最少 1）。"""
    max_rows = OrderedDict()
    for cat in category_order:
        max_rows[cat] = 1  # 至少 1 列
    for _model, cats in all_data:
        for cat in category_order:
            if cat in cats:
                max_rows[cat] = max(max_rows[cat], len(cats[cat]))
    return max_rows


# ---------------------------------------------------------------------------
# 2d. write_output
# ---------------------------------------------------------------------------
def write_output(all_data, category_order, max_rows, output_path):
    """寫入彙整 Excel。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consolidated Specs"

    bold = Font(bold=True)
    thin_bottom = Border(bottom=Side(style="thin"))
    wrap_align = Alignment(wrap_text=True, vertical="top")

    num_models = len(all_data)

    # --- Header row (Row 1) ---
    ws.cell(row=1, column=2, value="Model").font = bold
    for idx, (model_name, _cats) in enumerate(all_data):
        col = 3 + idx
        ws.cell(row=1, column=col, value=model_name).font = bold

    # --- 設定欄寬 ---
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 22
    for idx in range(num_models):
        ws.column_dimensions[get_column_letter(3 + idx)].width = 40

    # --- 資料列 (Row 2+) ---
    current_row = 2
    for cat in category_order:
        rows_for_cat = max_rows[cat]

        # 類別名稱只出現在第一列
        cell_b = ws.cell(row=current_row, column=2, value=cat)
        cell_b.font = bold
        cell_b.alignment = wrap_align

        # 填入各機種資料
        for idx, (_model, cats) in enumerate(all_data):
            col = 3 + idx
            lines = cats.get(cat, [])
            for r_offset, line in enumerate(lines):
                ws.cell(
                    row=current_row + r_offset, column=col, value=line
                ).alignment = wrap_align

        # 類別區塊最後一列加底線分隔
        last_row_of_block = current_row + rows_for_cat - 1
        for col in range(2, 3 + num_models):
            ws.cell(row=last_row_of_block, column=col).border = thin_bottom

        current_row += rows_for_cat

    # --- 凍結窗格 C2 ---
    ws.freeze_panes = "C2"

    wb.save(output_path)
    print(f"✔ 已產出彙整檔案: {output_path}")


# ---------------------------------------------------------------------------
# 2e. load_existing_models – 讀取已存在的彙整檔
# ---------------------------------------------------------------------------
def load_existing_models(output_path):
    """讀取已存在的 consolidated 輸出檔，回傳已匯入的 model 資料。

    Returns:
        list[(model_name, OrderedDict{category: [lines]})]
        若檔案不存在或無法讀取 → 回傳空 list
    """
    if not os.path.isfile(output_path):
        return []

    try:
        wb = openpyxl.load_workbook(output_path, read_only=True, data_only=True)
    except Exception as e:
        print(f"  ⚠ 無法讀取既有檔案 {output_path}: {e}")
        print("    將視為首次執行，重新匯入全部。")
        return []

    try:
        ws = wb.active

        # 讀取 Row 1 的 C 欄起所有 model 名稱
        model_names = []
        col = 3
        while True:
            val = ws.cell(row=1, column=col).value
            if val is None:
                break
            model_names.append(str(val).strip())
            col += 1

        if not model_names:
            wb.close()
            return []

        # 初始化每個 model 的 OrderedDict
        model_data = [OrderedDict() for _ in model_names]

        # 走訪 Row 2+ 還原每個 model 的資料
        current_cat = None
        row = 2
        while True:
            b_val = ws.cell(row=row, column=2).value
            # 檢查是否已超出資料範圍（B 欄和所有 model 欄都為空）
            all_empty = b_val is None
            if all_empty:
                has_data = False
                for i in range(len(model_names)):
                    if ws.cell(row=row, column=3 + i).value is not None:
                        has_data = True
                        break
                if not has_data:
                    break

            # B 欄有值 → 新類別開始
            if b_val is not None:
                current_cat = str(b_val).strip()
                for md in model_data:
                    if current_cat not in md:
                        md[current_cat] = []

            # 各 model 欄位有值 → 加入該類別的描述行
            if current_cat:
                for i in range(len(model_names)):
                    cell_val = ws.cell(row=row, column=3 + i).value
                    if cell_val is not None:
                        cell_str = str(cell_val).strip()
                        if cell_str:
                            model_data[i][current_cat].append(cell_str)

            row += 1

        wb.close()

        # 清理空類別
        result = []
        for name, data in zip(model_names, model_data):
            cleaned = OrderedDict()
            for cat, lines in data.items():
                if lines:
                    cleaned[cat] = lines
            result.append((name, cleaned))

        return result

    except Exception as e:
        print(f"  ⚠ 解析既有檔案時發生錯誤: {e}")
        print("    將視為首次執行，重新匯入全部。")
        wb.close()
        return []


# ---------------------------------------------------------------------------
# 2f. main – CLI 入口
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="將多個機種規格 Excel 檔案彙整為一份 Excel。"
    )
    parser.add_argument("folder", help="包含規格檔案的資料夾路徑")
    parser.add_argument(
        "-o", "--output", default=None,
        help="輸出檔案路徑（預設: <folder>/consolidated_specs.xlsx）"
    )
    parser.add_argument(
        "-f", "--force", action="store_true",
        help="忽略既有檔案，從頭全部重建"
    )
    args = parser.parse_args()

    folder = args.folder
    if not os.path.isdir(folder):
        print(f"錯誤: 資料夾不存在 – {folder}")
        sys.exit(1)

    output_path = args.output if args.output else os.path.join(folder, "consolidated_specs.xlsx")
    output_abs = os.path.abspath(output_path)

    # 搜尋 .xlsx 及 .xls（排除暫存檔與輸出檔本身）
    excel_files = sorted(
        glob.glob(os.path.join(folder, "*.xlsx"))
        + glob.glob(os.path.join(folder, "*.xls"))
    )
    # 去重（*.xls 會匹配到 .xlsx）並排序
    excel_files = sorted(set(excel_files))
    excel_files = [
        f
        for f in excel_files
        if not os.path.basename(f).startswith("~$")
        and os.path.abspath(f) != output_abs
    ]

    if not excel_files:
        print("錯誤: 資料夾中無可用的 Excel 檔案 (.xlsx / .xls)")
        sys.exit(1)

    # --- 載入既有資料（增量匯入）---
    existing_data = []
    existing_names = set()
    if not args.force and os.path.isfile(output_path):
        print(f"偵測到既有檔案: {output_path}")
        existing_data = load_existing_models(output_path)
        existing_names = {name for name, _ in existing_data}
        if existing_names:
            print(f"已匯入的機種 ({len(existing_names)}): {', '.join(sorted(existing_names))}\n")

    if args.force:
        print("使用 --force 模式，將重新匯入全部機種。\n")

    print(f"找到 {len(excel_files)} 個檔案，開始解析…\n")

    new_data = []
    skipped = 0
    for fpath in excel_files:
        fname = os.path.basename(fpath)
        try:
            model_name, cats = parse_spec_file(fpath)
            if model_name in existing_names:
                print(f"  ⏭ 跳過 {model_name}（已匯入）")
                skipped += 1
            else:
                new_data.append((model_name, cats))
                total_lines = sum(len(v) for v in cats.values())
                print(f"  ✔ {fname}  →  {model_name}  ({len(cats)} 類別, {total_lines} 描述行)")
        except Exception as e:
            print(f"  ⚠ 跳過 {fname}: {e}")

    # 若無新資料
    if not new_data:
        if skipped > 0:
            print(f"\n所有機種皆已匯入，無需更新。（共 {skipped} 個機種已跳過）")
        else:
            print("\n錯誤: 沒有成功解析任何檔案")
            sys.exit(1)
        return

    # 合併：既有資料 + 新資料
    all_data = existing_data + new_data

    category_order = merge_category_orders(all_data)
    max_rows_map = compute_max_rows(all_data, category_order)

    total_output_rows = sum(max_rows_map.values()) + 1  # +1 for header
    print(f"\n新增機種: {len(new_data)}, 總機種數: {len(all_data)}")
    print(f"類別數: {len(category_order)}, 輸出總列數: {total_output_rows}")

    write_output(all_data, category_order, max_rows_map, output_path)


if __name__ == "__main__":
    main()
