#!/usr/bin/env python3
"""
出成率计算脚本
计算指定 SKU 在小包装工序中的出成率
仅适用于原料（Y开头），按耗用材料口径汇总
"""
import argparse
import openpyxl
from pathlib import Path

BOM_DIR = Path("/Users/taojing/Documents/OpenClaw_Docs/Bom/")


def get_bom_path(month: str) -> Path:
    year, month_num = month.split("-")
    file_name = f"{year}年{int(month_num)}月bom表.xlsx"
    path = BOM_DIR / file_name
    if not path.exists():
        raise FileNotFoundError(f"BOM 文件不存在: {path}")
    return path


def load_bom(path: Path):
    return openpyxl.load_workbook(path, data_only=True)


def parse_product_info(sheet):
    """解析 Row 1-4 的产品维度信息（列从 Col 3 开始，每产品占2列）"""
    products = []
    max_col = sheet.max_column
    col = 3  # 产品从第3列开始
    
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
    """判断是否为原料（Y开头）"""
    return material_code is not None and str(material_code).upper().startswith("Y")


def calc_yield_rate(sheet, sku_code):
    """计算某 SKU 的出成率（仅原料）"""
    # 小包装 SKU 带 "-1" 后缀
    search_code = sku_code if sku_code.endswith("-1") else sku_code + "-1"
    
    products = parse_product_info(sheet)
    sku = None
    for p in products:
        if p["code"] == search_code:
            sku = p
            break
    
    if not sku:
        raise ValueError(f"SKU {search_code} 在小包装 Sheet 中不存在")
    
    inbound_weight = sku["inbound_weight"]
    results = []
    
    # Row 6+ 为物料明细
    for row in range(6, sheet.max_row + 1):
        material_code = sheet.cell(row=row, column=1).value
        material_name = sheet.cell(row=row, column=2).value
        
        if material_code is None or material_name is None:
            continue
        
        # 仅计算原料
        if not is_raw_material(str(material_code)):
            continue
        
        qty = sheet.cell(row=row, column=sku["qty_col"]).value
        if qty and float(qty) > 0:
            consumption = float(qty)
            yield_rate = inbound_weight / consumption
            results.append({
                "material_code": str(material_code).strip(),
                "material_name": str(material_name).strip(),
                "consumption": consumption,
                "yield_rate": yield_rate
            })
    
    return {
        "sku_code": search_code,
        "sku_name": sku["name"],
        "sku_spec": sku["spec"],
        "inbound_weight": inbound_weight,
        "materials": results
    }


def output_table(data, mode="material"):
    """输出飞书表格格式"""
    print(f"| SKU | {data['sku_code']} |")
    print(f"| 品名 | {data['sku_name']} |")
    print(f"| 规格 | {data['sku_spec']} |")
    print(f"| 入库重量 | {data['inbound_weight']:.2f} kg |")
    print()
    print("| 物料编码 | 耗用材料 | 耗用数量(kg) | 出成率 |")
    print("| --- | --- | --- | --- |")
    
    for m in data["materials"]:
        print(f"| {m['material_code']} | {m['material_name']} | {m['consumption']:.4f} | {m['yield_rate']:.4f} |")
    
    print()
    print(f"共 {len(data['materials'])} 种原料")
    print(f"出成率 = 入库重量 ÷ 耗用数量")
    print(f"数据来源：{data['month']} 小包装 BOM 表")


def main():
    parser = argparse.ArgumentParser(description="计算出成率")
    parser.add_argument("--month", required=True, help="月份，如 2026-01")
    parser.add_argument("--sku", required=True, help="产品代码，如 YQ00109XS")
    parser.add_argument("--mode", default="material", choices=["material", "large_category"], 
                       help="口径：material=按耗用材料，large_category=按原料大类（需额外映射表）")
    args = parser.parse_args()
    
    path = get_bom_path(args.month)
    wb = load_bom(path)
    
    # 仅小包装 sheet
    result = calc_yield_rate(wb.worksheets[0], args.sku)
    result["month"] = args.month
    
    output_table(result, args.mode)


if __name__ == "__main__":
    main()