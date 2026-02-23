import pandas as pd
import io


def clean_dataset(df):

    df = df.copy()

    # ---------- 1 移除全空列 ----------
    df = df.dropna(how='all')

    # ---------- 2 雙層表頭處理 ----------
    first_row = df.iloc[0].astype(str)

    if not first_row.str.contains(r'\d{4}', regex=True).any():
        df.columns = first_row
        df = df[1:]

    # ---------- 3 找日期欄 ----------
    date_col = None
    for col in df.columns:
        converted = pd.to_datetime(df[col], errors='coerce')
        if converted.notna().sum() > len(df) * 0.5:
            df[col] = converted
            date_col = col
            break

    # ---------- 4 文字保留 (N.D. / n.a. / 空白) ----------
    for col in df.columns:

        series = df[col].astype(str)

        # 保留空白
        series = series.replace(['nan', 'None'], '')

        # 統一 n.a.
        series = series.replace(['N/A', 'NA', 'na', 'n/a'], 'n.a.')

        # 保留 N.D.
        series = series.replace(['nd', 'ND', 'n.d'], 'N.D.')

        df[col] = series

    # ---------- 5 數值欄四捨五入 ----------
    for col in df.columns:

        # 嘗試轉數字
        numeric = pd.to_numeric(df[col], errors='coerce')

        # 如果此欄大部分是數字 → 視為數值欄
        if numeric.notna().sum() > len(df) * 0.5:

            df[col] = numeric.round(2)

            # 轉回字串，保持 Excel 輸出格式
            df[col] = df[col].map(
                lambda x: "" if pd.isna(x) else f"{x:.2f}"
            )

    # ---------- 6 日期排序 ----------
    if date_col:
        df = df.sort_values(by=date_col)

    df.reset_index(drop=True, inplace=True)

    return df


def dataframe_to_excel_response(df):

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)

    return output.getvalue()





