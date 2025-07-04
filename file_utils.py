import re
import pandas as pd
from collections import defaultdict

def merge_cp_files_by_keyword(cp_dataframes: dict) -> dict:
    grouped = defaultdict(list)

    # 将 DB, DB2, DB3... 聚合成同一组
    for key, df in cp_dataframes.items():
        for kw in ["华虹", "先进", "DB", "上华1厂", "上华2厂", "上华5厂"]:
            if key.startswith(kw):
                grouped[kw].append(df)
                break

    # 合并
    merged_cp_dataframes = {}
    for kw, df_list in grouped.items():
        # 排除空 DataFrame
        df_list = [df for df in df_list if df is not None and not df.empty]
        if df_list:
            merged_cp_dataframes[kw] = pd.concat(df_list, ignore_index=True)
        else:
            merged_cp_dataframes[kw] = pd.DataFrame()

    return merged_cp_dataframes



def extract_month_week(s):
    match = re.match(r"(\d{1,2})月WK(\d)", s)
    if match:
        month = int(match.group(1))
        week = int(match.group(2))
        return (month, week)
    return (99, 99)  # 放最后

def generate_fab_summary(cp_dataframes: dict) -> pd.DataFrame:
    import pandas as pd

    fab_rules = {
        "上华1厂": {"key": "上华", "fab": "CSMC-1", "part": "CUST_PARTNAME", "qty": "CURRENT_QTY", "date": "FORECAST_FAB_OUT_DATE"},
        "上华2厂": {"key": "上华", "fab": "CSMC-2", "part": "CUST_PARTNAME", "qty": "CURRENT_QTY", "date": "FORECAST_FAB_OUT_DATE"},
        "上华5厂": {"key": "上华", "fab": "CSMC-5", "part": "CUST_PARTNAME", "qty": "CURRENT_QTY", "date": "FORECAST_FAB_OUT_DATE"},
        "DB":     {"key": "DB", "fab": "DB", "part": "Customer Device", "qty": "Cur Wfs", "date": "Confirmed Date"},
        "华虹":    {"key": "华虹", "fab": "HHG", "part": "客户品名", "qty": "当前数量", "date": "最终确定交货日期"},
        "先进积塔": {"key": "先进", "fab": "ASMC-GTA", "part": "Device ID", "qty": "End Qty", "date": "Estimate Out Date"},
    }

    def get_week_label(dt: pd.Timestamp) -> str:
        if pd.isnull(dt): return None
        day = dt.day
        month = dt.strftime("%m月")
        if 1 <= day <= 7:
            return f"{month}WK1(1–7)"
        elif 8 <= day <= 15:
            return f"{month}WK2(8–15)"
        elif 16 <= day <= 22:
            return f"{month}WK3(16–22)"
        else:
            return f"{month}WK4(23–end)"

    all_rows = []

    for label, rule in fab_rules.items():
        fab_key = rule["key"]
        fab_value = rule["fab"]
        part_col = rule["part"]
        qty_col = rule["qty"]
        date_col = rule["date"]

        # 匹配所有相关表
        for sheet_name, df in cp_dataframes.items():
            if fab_key in sheet_name:
                if not all(col in df.columns for col in [part_col, qty_col, date_col]):
                    continue  # 跳过缺列的表

                df_temp = df[[part_col, qty_col, date_col]].copy()
                df_temp.columns = ["晶圆型号", "数量", "出货日期"]
                df_temp["出货日期"] = pd.to_datetime(df_temp["出货日期"], errors="coerce")
                df_temp["FAB"] = fab_value
                df_temp["周"] = df_temp["出货日期"].apply(get_week_label)
                all_rows.append(df_temp)

    df_all = pd.concat(all_rows, ignore_index=True)
    df_all = df_all.dropna(subset=["周", "数量", "晶圆型号"])

    # 透视表：晶圆型号 + FAB 行，列为每周，值为数量之和
    result = pd.pivot_table(
        df_all,
        index=["晶圆型号", "FAB"],
        columns="周",
        values="数量",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # 列顺序排序
    week_cols = sorted(
        [col for col in result.columns if isinstance(col, str) and "WK" in col],
        key=extract_month_week
    )
    result = result[["晶圆型号", "FAB"] + week_cols]


    return result


