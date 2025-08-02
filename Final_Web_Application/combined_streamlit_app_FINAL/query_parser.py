
import pandas as pd
import re

def process_query(query: str, df: pd.DataFrame):
    filtered_df = df.copy()
    chart_data = None

    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    filtered_df.columns = df.columns

    heat_match = re.search(r"heat[\s-]*([a-z0-9]+)", query, re.IGNORECASE)
    material_match = re.search(r"material\s*[=|:]*\s*([\w\s%.-]+)", query, re.IGNORECASE)
    qty_gt = re.search(r"quantit(?:y|ies)\s*>=*\s*(\d+\.?\d*)", query, re.IGNORECASE)
    qty_lt = re.search(r"quantit(?:y|ies)\s*<=*\s*(\d+\.?\d*)", query, re.IGNORECASE)
    qty_exact = re.search(r"quantit(?:y|ies)\s*=\s*(\d+\.?\d*)", query, re.IGNORECASE)
    price_gt = re.search(r"price.*>=\s*(\d+\.?\d*)", query, re.IGNORECASE)
    price_lt = re.search(r"price.*<=\s*(\d+\.?\d*)", query, re.IGNORECASE)
    section_match = re.search(r"section\s*=\s*([\w]+)", query, re.IGNORECASE)
    summary_match = "summary" in query.lower()
    top_match = re.search(r"top\s*(\d+)", query, re.IGNORECASE)
    yield_match = re.search(r"yield\s*[>=<]*\s*(\d+\.?\d*)", query, re.IGNORECASE)

    if heat_match:
        heat = heat_match.group(1).upper()
        filtered_df = filtered_df[filtered_df["heat_no."].astype(str).str.upper() == heat]

    if material_match:
        mat = material_match.group(1).strip().lower()
        filtered_df = filtered_df[filtered_df["material"].str.lower().str.contains(mat)]

    if qty_gt:
        val = float(qty_gt.group(1))
        filtered_df = filtered_df[filtered_df["actual_quantity"] >= val]
    if qty_lt:
        val = float(qty_lt.group(1))
        filtered_df = filtered_df[filtered_df["actual_quantity"] <= val]
    if qty_exact:
        val = float(qty_exact.group(1))
        filtered_df = filtered_df[filtered_df["actual_quantity"] == val]

    if price_gt:
        val = float(price_gt.group(1))
        filtered_df = filtered_df[filtered_df["price_per_mt"] >= val]
    if price_lt:
        val = float(price_lt.group(1))
        filtered_df = filtered_df[filtered_df["price_per_mt"] <= val]

    if section_match:
        val = section_match.group(1).strip().lower()
        filtered_df = filtered_df[filtered_df["section"].str.lower() == val]

    if yield_match:
        val = float(yield_match.group(1))
        filtered_df = filtered_df[filtered_df["yield"] >= val]

    if summary_match:
        group_data = filtered_df.groupby("material")["actual_quantity"].sum().reset_index().sort_values(by="actual_quantity", ascending=False)
        return group_data, {"type": "bar", "data": group_data}

    if top_match:
        top_n = int(top_match.group(1))
        top_data = filtered_df.sort_values(by="actual_quantity", ascending=False).head(top_n)
        return top_data, None

    return filtered_df, None
