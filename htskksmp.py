import streamlit as st
import pandas as pd
import datetime
import re
import html
import zipfile
from io import BytesIO

# ---------------------------------------------------------
# 定数・設定
# ---------------------------------------------------------
st.set_page_config(
    page_title="発注データ集計アプリ",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 統一フォーマットのカラム名定義
COL_DATE = "date"
COL_DEPT = "department"
COL_JAN = "jan"
COL_NAME = "product_name"
COL_QTY = "quantity"
COL_PRICE = "unit_price"
COL_PROMO = "promotion"
COL_AMOUNT = "total_amount"

# ---------------------------------------------------------
# ユーティリティ関数
# ---------------------------------------------------------

def clean_jan(jan_val):
    """JANコードのクリーニング"""
    s = str(jan_val).strip()
    # Excel特有の '49... や "49..." を除去
    s = re.sub(r"^['\"]", "", s)
    s = re.sub(r"['\"]$", "", s)
    # 小数点以下(.0)の削除
    s = re.sub(r'\.0$', '', s)
    return s

def clean_dept(dept_val):
    try:
        return str(int(float(dept_val))).zfill(3)
    except (ValueError, TypeError):
        return "000"

def parse_date_str(date_str, default_year=None):
    if default_year is None:
        default_year = datetime.date.today().year
        
    s = str(date_str).strip()
    if not s or s.lower() == 'nan': return None
    
    # 8桁数値 (YYYYMMDD)
    if re.match(r'^\d{8}$', s):
        try: return datetime.datetime.strptime(s, '%Y%m%d').date()
        except ValueError: pass
        
    # M/D 形式 (01/21, 1/21, 01/21(水) など)
    # 括弧などの不要な文字を除去して先頭の M/D を取得
    m = re.match(r'(\d{1,2})/(\d{1,2})', s)
    if m:
        month, day = map(int, m.groups())
        try: return datetime.date(default_year, month, day)
        except ValueError: pass
        
    # 標準形式
    try: return pd.to_datetime(s).date()
    except: pass
    return None

# ---------------------------------------------------------
# データ処理ロジック
# ---------------------------------------------------------

def find_column(df_cols, candidates):
    """カラム名の候補リストから、存在するものを探して返す"""
    for c in candidates:
        if c in df_cols:
            return c
    return None

def process_format_1(df: pd.DataFrame) -> pd.DataFrame:
    """形式1: ODR_RES形式 (リスト形式 / 1行ヘッダー)"""
    
    col_mappings = {
        COL_DATE: ['納品日', '納入日', '発注日', '日付'],
        COL_DEPT: ['部門', '部門コード'],
        COL_JAN: ['商品コード', 'JANコード', 'JAN', 'JanCode'],
        COL_NAME: ['商品名', '商品名称', '品名'],
        COL_QTY: ['発注数量', '数量', '数'],
        COL_PRICE: ['売単価', '売価', '単価'],
        COL_PROMO: ['発注区分', '販促', '特売区分']
    }
    
    rename_dict = {}
    for unified_name, candidates in col_mappings.items():
        found = find_column(df.columns, candidates)
        if found:
            rename_dict[found] = unified_name
    
    # 必須カラムチェック
    if COL_DATE not in rename_dict.values() or COL_JAN not in rename_dict.values():
        return pd.DataFrame()

    df = df.rename(columns=rename_dict)
    
    # 欠損カラム補完
    for c in [COL_DEPT, COL_NAME, COL_QTY, COL_PRICE, COL_PROMO]:
        if c not in df.columns:
            df[c] = "" if c == COL_PROMO else 0

    # データ変換
    df[COL_DATE] = df[COL_DATE].apply(lambda x: parse_date_str(x))
    df[COL_DEPT] = df[COL_DEPT].apply(clean_dept)
    df[COL_JAN] = df[COL_JAN].apply(clean_jan)
    df[COL_QTY] = pd.to_numeric(df[COL_QTY], errors='coerce').fillna(0)
    df[COL_PRICE] = pd.to_numeric(df[COL_PRICE], errors='coerce').fillna(0)
    df[COL_PROMO] = df[COL_PROMO].fillna("").astype(str).replace(['nan', 'None'], '')
    
    cols = [COL_DATE, COL_DEPT, COL_JAN, COL_NAME, COL_QTY, COL_PRICE, COL_PROMO]
    return df[cols].dropna(subset=[COL_DATE])

def process_format_2_from_df(df: pd.DataFrame, year_hint=None) -> pd.DataFrame:
    """形式2: OrderCheckList形式 (マトリックス / 2行ヘッダー)"""
    new_cols = []
    last_top = None
    
    # マルチインデックスの整形 (Unnamedの補完)
    # 上段ヘッダー(top)が結合セルで、Pandas読み込み時にUnnamedになっている箇所を埋める
    for top, bottom in df.columns:
        s_top = str(top)
        # "Unnamed"が含まれておらず、かつ"週合計"のような集計列でない場合、それを現在の上段見出しとする
        if "Unnamed" not in s_top and "週合計" not in s_top:
            last_top = top
        
        # 上段がUnnamedなら、直前の有効な見出し(last_top)を使う
        final_top = last_top if "Unnamed" in s_top else top
        new_cols.append((final_top, bottom))
    
    df.columns = pd.MultiIndex.from_tuples(new_cols)
    
    fixed_col_map = {}
    date_cols = []
    
    # --- カラム解析 (修正箇所) ---
    # TopとBottomの両方をチェックして、固定列（JAN、部門、商品名）を特定する
    # 理由: 2行ヘッダーの場合、"部門"などのラベルはTopにあり、Bottomが"Unnamed"になることがあるため
    for top, bottom in new_cols:
        s_top = str(top)
        s_bottom = str(bottom)
        
        # 判定用文字列（両方チェック）
        # JANコード列の特定
        if ("JAN" in s_top or "JAN" in s_bottom) or ("商品コード" in s_top or "商品コード" in s_bottom):
            fixed_col_map[(top, bottom)] = 'JANコード'
        
        # 部門列の特定
        elif "部門" in s_top or "部門" in s_bottom:
            fixed_col_map[(top, bottom)] = '部門'
            
        # 商品名列の特定
        elif "商品名" in s_top or "商品名" in s_bottom:
            fixed_col_map[(top, bottom)] = '商品名'
        
        # 日付列の判定 (M/D 形式が含まれるか)
        # 日付はTopヘッダーにある ("01/21(水)" など)
        elif top is not None and "週合計" not in s_top and re.search(r'\d{1,2}/\d{1,2}', s_top):
            if top not in date_cols: date_cols.append(top)
    
    records = []
    for _, row in df.iterrows():
        # 固定情報の抽出
        jan_key = next((k for k, v in fixed_col_map.items() if v == 'JANコード'), None)
        dept_key = next((k for k, v in fixed_col_map.items() if v == '部門'), None)
        name_key = next((k for k, v in fixed_col_map.items() if v == '商品名'), None)
        
        jan = row[jan_key] if jan_key else None
        if pd.isna(jan): continue
        
        dept = row[dept_key] if dept_key else "000"
        name = row[name_key] if name_key else ""

        for date_str in date_cols:
            if not date_str or date_str == "nan": continue
            
            # 日付解析
            date_obj = parse_date_str(date_str, default_year=year_hint)
            if not date_obj: continue

            # マトリックスデータの取得
            try:
                # 数量 (Bottomヘッダーが"数量")
                # カラムキーは (top=日付文字列, bottom="数量")
                qty_val = row.get((date_str, '数量'))
                
                # 数量が取得できない、またはNaN/0の場合はスキップ
                if pd.isna(qty_val): continue
                qty = pd.to_numeric(qty_val, errors='coerce')
                if pd.isna(qty) or qty == 0: continue 

                # 売価・販促 (Bottomヘッダーが"売価", "販促")
                price = pd.to_numeric(row.get((date_str, '売価'), 0), errors='coerce')
                promo_val = row.get((date_str, '販促'))
                promo_str = str(promo_val) if not pd.isna(promo_val) else ""
                
                record = {
                    COL_DATE: date_obj,
                    COL_DEPT: clean_dept(dept),
                    COL_JAN: clean_jan(jan),
                    COL_NAME: name,
                    COL_QTY: qty,
                    COL_PRICE: price,
                    COL_PROMO: promo_str
                }
                records.append(record)
            except:
                continue
            
    return pd.DataFrame(records)

def load_data(uploaded_file) -> pd.DataFrame:
    """
    形式判定ロジック強化版
    """
    if uploaded_file is None: return pd.DataFrame()
    
    encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
    
    for enc in encodings:
        uploaded_file.seek(0)
        try:
            # 先頭の数行を読んで構造を解析
            raw_lines = []
            for _ in range(15):
                line = uploaded_file.readline()
                if not line: break
                try:
                    raw_lines.append(line.decode(enc).strip())
                except:
                    pass
            
            if not raw_lines: continue
            
            full_text = "\n".join(raw_lines)
            
            # --- ヘッダー行と形式の特定 ---
            header_row_idx = -1
            is_matrix_format = False
            
            for i, line in enumerate(raw_lines):
                # 「部門」と「JAN」(または商品コード) がある行をヘッダーとみなす
                if "部門" in line and ("JAN" in line or "商品コード" in line):
                    header_row_idx = i
                    
                    # 形式の分岐: ヘッダー行に「納品日」や「発注日」があればリスト形式(Format1)
                    if "納品日" in line or "発注日" in line:
                        is_matrix_format = False
                    # なければマトリックス形式(Format2)とみなす
                    else:
                        is_matrix_format = True
                    break
            
            # ヘッダーが見つからなかった場合、次へ
            if header_row_idx == -1: continue

            # --- 読み込み実行 ---
            uploaded_file.seek(0)
            
            if not is_matrix_format:
                # Format 1 (ODR_RES)
                df = pd.read_csv(uploaded_file, header=header_row_idx, encoding=enc, on_bad_lines='skip', dtype=str)
                processed = process_format_1(df)
                if not processed.empty: return processed

            else:
                # Format 2 (OrderCheckList)
                # 年号の抽出 (メタデータ行から 20xx を探す)
                year_hint = None
                m_year = re.search(r'20\d{2}', full_text)
                if m_year:
                    year_hint = int(m_year.group(0))
                
                # ヘッダーは2行分として読み込む
                # header=[Topの行, Bottomの行]
                df = pd.read_csv(uploaded_file, header=[header_row_idx, header_row_idx+1], encoding=enc, on_bad_lines='skip', dtype=str)
                processed = process_format_2_from_df(df, year_hint=year_hint)
                if not processed.empty: return processed

        except Exception:
            continue

    return pd.DataFrame()

# ---------------------------------------------------------
# CSV生成・POP生成
# ---------------------------------------------------------

def create_matrix_csv(df: pd.DataFrame) -> bytes:
    if df.empty: return b""
    
    meta_df = df.groupby(COL_JAN).agg({
        COL_DEPT: 'first',
        COL_NAME: 'first',
        COL_PRICE: 'max',
        COL_PROMO: 'first'
    })

    pivot_df = df.pivot_table(
        index=COL_JAN,
        columns=COL_DATE, 
        values=COL_QTY, 
        aggfunc='sum', 
        fill_value=0
    )
    
    result_df = pd.concat([meta_df, pivot_df], axis=1).reset_index()

    date_cols = sorted([c for c in result_df.columns if isinstance(c, (datetime.date, datetime.datetime))])
    
    result_df['合計数量'] = result_df[date_cols].sum(axis=1)
    result_df['合計金額'] = result_df['合計数量'] * result_df[COL_PRICE]

    col_map = {COL_DEPT: '部門', COL_JAN: 'JAN', COL_NAME: '商品名', COL_PRICE: '単価', COL_PROMO: '販促'}
    date_col_map = {d: d.strftime('%Y/%m/%d') for d in date_cols}
    result_df = result_df.rename(columns={**col_map, **date_col_map})
    
    base_cols = ['部門', 'JAN', '商品名', '単価']
    date_str_cols = [d.strftime('%Y/%m/%d') for d in date_cols]
    final_cols = base_cols + date_str_cols + ['合計数量', '合計金額', '販促']
    
    existing_cols = [c for c in final_cols if c in result_df.columns]
    result_df = result_df[existing_cols]
    result_df['JAN'] = "'" + result_df['JAN'].astype(str)

    csv_buffer = BytesIO()
    result_df.to_csv(csv_buffer, index=False, encoding='utf_8_sig')
    return csv_buffer.getvalue()

def generate_svg(row, daily_qty_map, start_date):
    dept = row[COL_DEPT]
    jan = row[COL_JAN]
    name = html.escape(str(row[COL_NAME]))
    price = row[COL_PRICE]
    total_qty = row[COL_QTY]
    total_amount = row[COL_AMOUNT]
    promo = str(row[COL_PROMO]) if row[COL_PROMO] else ""
    
    fc = "1F"
    if total_amount >= 100000: fc = "5F"
    elif total_amount >= 50000: fc = "4F"
    elif total_amount >= 20000: fc = "3F"
    elif total_amount >= 5000: fc = "2F"
    
    is_sale = False
    if promo and ("特売" in promo or "セール" in promo or "スポ" in promo):
        is_sale = True
    
    clr = "#ef4444" if is_sale else "#334155"
    bg = "#fef2f2" if is_sale else "#f8fafc"
    
    calendar_svg_parts = []
    current_d = start_date
    for i in range(7):
        d_str = f"{current_d.month}/{current_d.day}"
        qty = daily_qty_map.get(current_d, 0)
        fill_col = "#fff" if i % 2 == 0 else "#f9fafb"
        text_fill = "#000" if qty > 0 else "#d1d5db"
        x_pos = 5 + (i * 84)
        part = f"""<g transform="translate({x_pos}, 355)"><rect width="84" height="80" fill="{fill_col}" stroke="#e2e8f0"/><text x="42" y="20" font-family="sans-serif" font-size="12" fill="#64748b" text-anchor="middle">{d_str}</text><text x="42" y="60" font-family="sans-serif" font-size="26" fill="{text_fill}" font-weight="bold" text-anchor="middle">{int(qty)}</text></g>"""
        calendar_svg_parts.append(part)
        current_d += datetime.timedelta(days=1)
    
    calendar_svg = "".join(calendar_svg_parts)
    svg_content = f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 600 440" style="background:#fff;"><rect x="5" y="5" width="590" height="430" fill="white" stroke="{clr}" stroke-width="6"/><rect x="5" y="5" width="590" height="65" fill="{bg}"/><line x1="200" y1="5" x2="200" y2="70" stroke="{clr}" stroke-width="2"/><line x1="400" y1="5" x2="400" y2="70" stroke="{clr}" stroke-width="2"/><line x1="5" y1="70" x2="595" y2="70" stroke="{clr}" stroke-width="2"/><text x="102" y="45" font-family="sans-serif" font-size="28" font-weight="900" text-anchor="middle" fill="{clr}">{promo if promo else '通常'}</text><text x="215" y="25" font-family="sans-serif" font-size="12" fill="#64748b">部門</text><text x="215" y="55" font-family="sans-serif" font-size="24" font-weight="bold" fill="#1e293b">{dept}</text><text x="415" y="25" font-family="sans-serif" font-size="12" fill="#64748b">フェイス数</text><text x="500" y="55" font-family="sans-serif" font-size="40" font-weight="900" text-anchor="middle" fill="{clr}">{fc}</text><text x="25" y="105" font-family="sans-serif" font-size="12" fill="#64748b">JAN CODE</text><text x="25" y="145" font-family="monospace" font-size="40" font-weight="bold" letter-spacing="4" fill="#1e293b">{jan}</text><text x="25" y="185" font-family="sans-serif" font-size="34" font-weight="900" fill="#000">{name}</text><line x1="5" y1="205" x2="595" y2="205" stroke="#e2e8f0" stroke-width="2"/><text x="25" y="235" font-family="sans-serif" font-size="12" fill="#64748b">単価</text><text x="25" y="275" font-family="sans-serif" font-size="32" font-weight="bold">¥ {int(price):,}</text><text x="25" y="315" font-family="sans-serif" font-size="12" fill="#64748b">合計見込額</text><text x="25" y="345" font-family="sans-serif" font-size="28" font-weight="bold" fill="{clr}">¥ {int(total_amount):,}</text><rect x="340" y="215" width="240" height="130" rx="8" fill="#f1f5f9"/><text x="360" y="245" font-family="sans-serif" font-size="14" font-weight="bold" fill="#475569">合計点数</text><text x="460" y="325" font-family="sans-serif" font-size="90" font-weight="900" text-anchor="middle" fill="#000">{int(total_qty)}</text><text x="560" y="325" font-family="sans-serif" font-size="20" font-weight="bold" text-anchor="end" fill="#475569">点</text>{calendar_svg}</svg>"""
    return svg_content

def create_pop_zip(agg_df, raw_df, start_date) -> bytes:
    zip_buffer = BytesIO()
    daily_map = {}
    temp_df = raw_df[[COL_JAN, COL_DATE, COL_QTY]].copy()
    temp_df[COL_QTY] = pd.to_numeric(temp_df[COL_QTY], errors='coerce').fillna(0)
    grouped = temp_df.groupby([COL_JAN, COL_DATE])[COL_QTY].sum().reset_index()
    for _, r in grouped.iterrows():
        j = r[COL_JAN]; d = r[COL_DATE]; q = r[COL_QTY]
        if j not in daily_map: daily_map[j] = {}
        daily_map[j][d] = q

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, row in agg_df.iterrows():
            jan = row[COL_JAN]; dept = row[COL_DEPT]
            item_daily_map = daily_map.get(jan, {})
            svg_str = generate_svg(row, item_daily_map, start_date)
            safe_jan = re.sub(r'[\\/:*?"<>|]', '', str(jan))
            safe_dept = re.sub(r'[\\/:*?"<>|]', '', str(dept))
            zf.writestr(f"{safe_dept}_{safe_jan}.svg", svg_str.encode("utf-8"))
    return zip_buffer.getvalue()

# ---------------------------------------------------------
# アプリケーション本体
# ---------------------------------------------------------

def main():
    st.title("📦 発注データ集計アプリ")

    # ----------------------------------------
    # LINEブラウザ対策
    # ----------------------------------------
    st.markdown("""
    <style>
    .line-warning {
        background-color: #f0f2f6; 
        padding: 15px; 
        border-radius: 8px; 
        margin-bottom: 25px; 
        border-left: 6px solid #ff4b4b;
    }
    .line-warning h4 { margin: 0 0 10px 0; color: #ff4b4b; }
    .line-warning p { margin: 0; font-size: 14px; color: #31333F; line-height: 1.6; }
    </style>
    <div class="line-warning">
        <h4>⚠️ LINEから開いている方へ</h4>
        <p><b>LINE内蔵ブラウザではアップロードが反応しない場合があります。</b><br>
        反応しない場合は、右上のメニュー（︙または↗️）から「ブラウザで開く」を選択してください。</p>
    </div>
    """, unsafe_allow_html=True)

    # ----------------------------------------
    # 1. データ読込とフィルタ設定
    # ----------------------------------------
    with st.expander("🛠️ データ読込・フィルタ設定", expanded=True):
        st.caption("Step 1: データのアップロード")
        uploaded_files = st.file_uploader(
            "CSVファイルをドロップ", 
            type=["csv", "txt"], 
            accept_multiple_files=True
        )
        
        all_data = []
        if uploaded_files:
            for f in uploaded_files:
                df = load_data(f)
                if not df.empty:
                    all_data.append(df)
                    st.success(f"OK: {f.name} ({len(df)}行)")
                else:
                    st.error(f"NG: {f.name} (形式不明またはデータなし)")

        if all_data:
            master_df = pd.concat(all_data, ignore_index=True)
            master_df[COL_AMOUNT] = master_df[COL_QTY] * master_df[COL_PRICE]
            
            st.markdown("---")
            st.caption("Step 2: フィルタリング")

            # 期間設定
            min_date = master_df[COL_DATE].min()
            max_date = master_df[COL_DATE].max()
            if pd.isna(min_date): min_date = datetime.date.today()
            if pd.isna(max_date): max_date = datetime.date.today()
            
            date_range = st.slider(
                "期間を指定",
                min_value=min_date,
                max_value=max_date,
                value=(min_date, max_date),
                format="MM/DD"
            )
            start_d, end_d = date_range

            # 部門設定
            dept_options = sorted(master_df[COL_DEPT].unique())
            if 'selected_depts' not in st.session_state:
                st.session_state.selected_depts = dept_options

            def select_all_depts(): st.session_state.selected_depts = dept_options
            
            st.button("全部門を選択", on_click=select_all_depts, use_container_width=True)
            selected_depts = st.multiselect("部門を指定", dept_options, key="selected_depts")

            # 販促設定
            unique_promos = sorted(list(set(master_df[COL_PROMO].astype(str).unique())))
            promo_options = [p for p in unique_promos if p.strip()]
            if "" in unique_promos or "nan" in unique_promos:
                if "" not in promo_options: promo_options.append("")
            
            if 'selected_promos' not in st.session_state:
                st.session_state.selected_promos = promo_options
                
            def select_all_promos(): st.session_state.selected_promos = promo_options

            st.button("全販促タイプを選択", on_click=select_all_promos, use_container_width=True)
            selected_promos = st.multiselect("販促タイプを指定", promo_options, key="selected_promos")

            # キーワード検索
            search_text = st.text_area("キーワード検索 (商品名・JAN)", height=68, placeholder="スペース区切りで複数可")
            
        else:
            st.info("まずはファイルをアップロードしてください👆")
            st.stop()

    # -------------------------------------------------
    # 結果表示
    # -------------------------------------------------
    mask = (
        (master_df[COL_DATE] >= start_d) & 
        (master_df[COL_DATE] <= end_d) &
        (master_df[COL_DEPT].isin(selected_depts))
    )
    filtered_df = master_df[mask].copy()

    if selected_promos:
        filtered_df = filtered_df[filtered_df[COL_PROMO].astype(str).isin(selected_promos)]
    elif len(promo_options) > 0:
        filtered_df = filtered_df.iloc[0:0]

    if search_text:
        keywords = [k for k in re.split(r'[,\s\n\u3000]+', search_text) if k]
        if keywords:
            match_condition = pd.Series([False] * len(filtered_df), index=filtered_df.index)
            for k in keywords:
                if k.isdigit():
                    match_condition |= (filtered_df[COL_JAN] == k)
                match_condition |= filtered_df[COL_JAN].astype(str).str.contains(k, na=False)
                match_condition |= filtered_df[COL_NAME].astype(str).str.contains(k, na=False)
            filtered_df = filtered_df[match_condition]

    # 集計
    agg_view = filtered_df.groupby(COL_JAN, as_index=False).agg({
        COL_DEPT: 'first', COL_NAME: 'first', COL_PRICE: 'max', 
        COL_QTY: 'sum', COL_AMOUNT: 'sum', COL_PROMO: 'first'
    }).sort_values(by=COL_QTY, ascending=False)

    st.subheader("📊 集計結果")

    m1, m2, m3 = st.columns(3)
    m1.metric("合計金額", f"¥{agg_view[COL_AMOUNT].sum():,.0f}")
    m2.metric("総数量", f"{agg_view[COL_QTY].sum():,.0f}")
    m3.metric("アイテム", f"{len(agg_view)}品")

    st.dataframe(
        agg_view,
        column_config={
            COL_DEPT: "部門", COL_JAN: "JAN", COL_NAME: "商品名",
            COL_PRICE: st.column_config.NumberColumn("単価", format="¥%d"),
            COL_QTY: st.column_config.NumberColumn("数量", format="%d"),
            COL_AMOUNT: st.column_config.NumberColumn("金額", format="¥%d"),
            COL_PROMO: "販促"
        },
        use_container_width=True, hide_index=True, height=300
    )

    st.markdown("---")
    st.subheader("📤 ダウンロード")
    
    csv = create_matrix_csv(filtered_df)
    if csv:
        st.download_button(
            label="📄 マトリックスCSVを保存",
            data=csv,
            file_name=f"Order_{datetime.datetime.now():%Y%m%d}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    if not agg_view.empty:
        pop = create_pop_zip(agg_view, filtered_df, start_d)
        st.download_button(
            label="🎨 POP画像を一括保存 (ZIP)",
            data=pop,
            file_name=f"POP_{datetime.datetime.now():%Y%m%d}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
