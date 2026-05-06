#!/usr/bin/env python3
"""
出成率批量计算脚本
输入月份，自动遍历该月所有 SKU，计算出成率并生成合并表
支持两种口径：按耗用材料 / 按原料大类
"""
import argparse
import openpyxl
from pathlib import Path
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

BOM_DIR = Path("/Users/taojing/Documents/OpenClaw_Docs/Bom/")
MAPPING_DIR = Path("/Users/taojing/.openclaw-dev/media/inbound/")


def get_bom_path(month: str) -> Path:
    year, month_num = month.split("-")
    file_name = f"{year}年{int(month_num)}月bom表.xlsx"
    path = BOM_DIR / file_name
    if not path.exists():
        raise FileNotFoundError(f"BOM 文件不存在: {path}")
    return path


def find_latest_mapping(month: str) -> Path:
    """查找最新的原料大类映射表
    
    匹配策略：
    1. 找到 `media/inbound/原料大类映射表*.xlsx` 所有候选文件
    2. 取 st_mtime 最新的文件
    3. 适用于当月及过往所有月份（映射表按最新期间编制，兼容历史）
    """
    candidates = list(MAPPING_DIR.glob("原料大类映射表*.xlsx"))
    if not candidates:
        raise FileNotFoundError(f"未找到任何原料大类映射表，搜索路径：{MAPPING_DIR}")
    
    # 取最新文件
    latest = max(candidates, key=lambda p: p.stat().st_mtime)
    return latest


def load_bom(path: Path):
    return openpyxl.load_workbook(path, data_only=True)


def parse_product_info(sheet):
    """解析产品维度信息（列从 Col 3 开始，每产品占2列）"""
    products = []
    max_col = sheet.max_column
    col = 3

    while col + 1 <= max_col:
        code = sheet.cell(row=1, column=col).value
        name = sheet.cell(row=2, column=col).value
        spec = sheet.cell(row=3, column=col).value
        weight = sheet.cell(row=4, column=col).value

        if not code or (isinstance(code, str) and not code.strip()):
            break

        products.append({
            "code": str(code).strip(),
            "name": str(name).strip() if name else "",
            "spec": str(spec).strip() if spec else "",
            "inbound_weight": float(weight) if weight and str(weight).strip() else 0,
            "qty_col": col,
            "amt_col": col + 1
        })
        col += 2

    return products


def is_raw_material(material_code: str) -> bool:
    return material_code is not None and str(material_code).upper().startswith("Y")


def calc_all_yield_rates(sheet, products):
    """计算所有 SKU 的出成率，返回 list"""
    results = []  # [{sku_code, sku_name, sku_spec, inbound_weight, material_code, material_name, consumption, yield_rate}]

    for p in products:
        inbound = p["inbound_weight"]
        if inbound <= 0:
            continue

        for row in range(6, sheet.max_row + 1):
            material_code = sheet.cell(row=row, column=1).value
            material_name = sheet.cell(row=row, column=2).value

            if material_code is None or material_name is None:
                continue

            if not is_raw_material(str(material_code)):
                continue

            qty = sheet.cell(row=row, column=p["qty_col"]).value
            if qty and str(qty).strip() and float(qty) > 0:
                consumption = float(qty)
                results.append({
                    "sku_code": p["code"],
                    "sku_name": p["name"],
                    "sku_spec": p["spec"],
                    "inbound_weight": inbound,
                    "material_code": str(material_code).strip(),
                    "material_name": str(material_name).strip(),
                    "consumption": consumption,
                    "yield_rate": inbound / consumption
                })

    return results


def load_material_mapping(mapping_path):
    """加载原料大类映射表，返回 {原料名称: 原料大类} dict"""
    wb = openpyxl.load_workbook(mapping_path, data_only=True)
    sheet = wb.active

    mapping = {}
    for row in range(2, sheet.max_row + 1):
        material_name = sheet.cell(row=row, column=1).value
        category = sheet.cell(row=row, column=2).value

        if material_name and category:
            mapping[str(material_name).strip()] = str(category).strip()

    return mapping


def apply_mapping(yield_rates, mapping):
    """将 yield_rates 中的 material_name 替换为原料大类（如果存在映射）"""
    for r in yield_rates:
        r["material_category"] = mapping.get(r["material_name"], r["material_name"])
    return yield_rates


def write_excel(data_by_category, month, output_path):
    """写入 Excel（两个 sheet）"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # ---- Sheet 1: 按耗用材料口径 ----
    ws1 = wb.create_sheet("出成率（按耗用材料）")
    headers1 = ["SKU", "品名", "规格", "入库重量(kg)", "原料名称", "耗用数量(kg)", "出成率"]

    for col_idx, h in enumerate(headers1, 1):
        cell = ws1.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')

    row_idx = 2
    for item in data_by_category.get("material", []):
        values = [
            item["sku_code"],
            item["sku_name"],
            item["sku_spec"],
            item["inbound_weight"],
            item["material_name"],
            item["consumption"],
            round(item["yield_rate"], 4)
        ]
        for col_idx, v in enumerate(values, 1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=v)
            cell.border = thin_border
        row_idx += 1

    # Auto width for sheet 1
    for col in ws1.columns:
        max_len = max(len(str(cell.value) if cell.value else "") for cell in col)
        ws1.column_dimensions[col[0].column_letter].width = max(max_len + 2, 12)

    # ---- Sheet 2: 按原料大类口径 ----
    ws2 = wb.create_sheet("出成率（按原料大类）")
    headers2 = ["SKU", "品名", "规格", "入库重量(kg)", "原料大类", "出成率"]

    for col_idx, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')

    row_idx = 2
    for item in data_by_category.get("category", []):
        values = [
            item["sku_code"],
            item["sku_name"],
            item["sku_spec"],
            item["inbound_weight"],
            item["material_category"],
            round(item["yield_rate"], 4)
        ]
        for col_idx, v in enumerate(values, 1):
            cell = ws2.cell(row=row_idx, column=col_idx, value=v)
            cell.border = thin_border
        row_idx += 1

    for col in ws2.columns:
        max_len = max(len(str(cell.value) if cell.value else "") for cell in col)
        ws2.column_dimensions[col[0].column_letter].width = max(max_len + 2, 12)

    # Summary sheet
    ws3 = wb.create_sheet("汇总")
    ws3.cell(row=1, column=1, value="月份").font = header_font
    ws3.cell(row=1, column=2, value=month)
    ws3.cell(row=2, column=1, value="生成时间").font = header_font
    ws3.cell(row=2, column=2, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    ws3.cell(row=3, column=1, value="SKU 总数").font = header_font
    ws3.cell(row=3, column=2, value=len(set(item["sku_code"] for item in data_by_category.get("material", []))))
    ws3.cell(row=4, column=1, value="按耗用材料行数").font = header_font
    ws3.cell(row=4, column=2, value=len(data_by_category.get("material", [])))
    ws3.cell(row=5, column=1, value="按原料大类行数").font = header_font
    ws3.cell(row=5, column=2, value=len(data_by_category.get("category", [])))

    wb.save(output_path)
    return output_path


def aggregate_by_category(yield_rates):
    """按原料大类聚合（对同一 SKU + 原料大类下的多物料取加权平均出成率）"""
    # 按 SKU + 原料大类聚合
    agg = {}  # {(sku_code, material_category): {"sku_name":..., "sku_spec":..., "inbound_weight":..., "material_category":..., "items": [...]}}
    for r in yield_rates:
        key = (r["sku_code"], r["material_category"])
        if key not in agg:
            agg[key] = {
                "sku_code": r["sku_code"],
                "sku_name": r["sku_name"],
                "sku_spec": r["sku_spec"],
                "inbound_weight": r["inbound_weight"],
                "material_category": r["material_category"],
                "items": []
            }
        agg[key]["items"].append(r)

    result = []
    for v in agg.values():
        # 正确做法：入库重量 ÷ 该大类下所有原料的汇总耗用
        # ⚠️ 不能用加权平均！出成率是非线性指标，
        #    加权平均会产生错误结果（如入库90622却算出出成率5.03）
        total_consumption = sum(i["consumption"] for i in v["items"])
        if total_consumption > 0:
            correct_yield = v["inbound_weight"] / total_consumption
        else:
            correct_yield = v["items"][0]["yield_rate"]
        result.append({
            "sku_code": v["sku_code"],
            "sku_name": v["sku_name"],
            "sku_spec": v["sku_spec"],
            "inbound_weight": v["inbound_weight"],
            "material_category": v["material_category"],
            "yield_rate": correct_yield
        })
    return result


def main():
    parser = argparse.ArgumentParser(description="批量计算出成率并生成合并表")
    parser.add_argument("--month", required=True, help="月份，如 2026-04")
    parser.add_argument("--output", help="输出 xlsx 路径，默认保存到 BOM 同目录")
    parser.add_argument("--no-mapping", action="store_true", help="跳过原料大类映射（仅生成按耗用材料口径）")
    args = parser.parse_args()

    print(f"📖 读取 BOM 表: {args.month}")
    bom_path = get_bom_path(args.month)
    wb = load_bom(bom_path)

    # 只解析小包装 sheet
    sheet = wb.worksheets[0]
    print(f"   Sheet: {sheet.title}")
    products = parse_product_info(sheet)
    print(f"   产品数: {len(products)}")

    print(f"🔢 计算出成率...")
    yield_rates = calc_all_yield_rates(sheet, products)
    print(f"   原料明细行数: {len(yield_rates)}")

    if not args.no_mapping:
        print(f"🗺️  加载原料大类映射...")
        try:
            mapping_path = find_latest_mapping(args.month)
            print(f"   映射表: {mapping_path.name}")
            mapping = load_material_mapping(mapping_path)
            print(f"   映射条目: {len(mapping)}")
            yield_rates = apply_mapping(yield_rates, mapping)
        except FileNotFoundError as e:
            print(f"   ⚠️  {e}")
            print(f"   跳过原料大类口径")
            yield_rates = apply_mapping(yield_rates, {})

    data_by_category = {
        "material": yield_rates,
        "category": aggregate_by_category(yield_rates)
    }

    if args.output:
        output_path = Path(args.output)
    else:
        year, month_num = args.month.split("-")
        output_path = BOM_DIR / f"{year}年{int(month_num)}月出成率合并.xlsx"

    print(f"💾 生成 Excel: {output_path}")
    write_excel(data_by_category, args.month, output_path)

    print()
    print(f"✅ 完成！")
    print(f"   按耗用材料: {len(data_by_category['material'])} 行")
    print(f"   按原料大类: {len(data_by_category['category'])} 行")

    return output_path


if __name__ == "__main__":
    main()