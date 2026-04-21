import re
import streamlit as st
import pandas as pd
import requests
from io import BytesIO, StringIO, BytesIO as BIO
from datetime import date, datetime, timedelta
from requests.auth import HTTPBasicAuth

from google.oauth2 import service_account
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import (
    DateRange, Dimension, Metric, RunReportRequest
)

st.set_page_config(page_title="Price Checker", layout="wide")

# ============================================================
# LOGOWANIE
# ============================================================
APP_PASSWORD = st.secrets["app"]["app_password"]

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Logowanie do aplikacji")
    st.markdown("Wpisz hasło, aby uzyskać dostęp do porównania danych.")
    with st.form("login_form"):
        password_input = st.text_input("Hasło:", type="password")
        submitted = st.form_submit_button("🔓 Zaloguj")
        if submitted:
            if password_input == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Nieprawidłowe hasło!")
    st.stop()

# ============================================================

st.title("Price Checker – porównanie sklepów")

st.markdown("""
<style>
[data-baseweb="tag"] span {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    max-width: none !important;
}
[data-baseweb="tag"] {
    height: auto !important;
    white-space: normal !important;
}
[data-baseweb="select"] [role="option"] {
    white-space: normal !important;
    word-break: break-word !important;
}
</style>
""", unsafe_allow_html=True)


HTTP_USERNAME = st.secrets["http_auth"]["username"]
HTTP_PASSWORD = st.secrets["http_auth"]["password"]

SHOP_TO_MPK = {
    '50stylepl': 'S501', 'butysportowe': 'S503', 'sizeerpl': 'S500',
    'sizeerde': 'G500', 'sizeercz': 'CZ50', 'sizeersk': 'SK50',
    'sizeerlt': 'LT50', 'sizeerro': 'RO50', 'timberland': 'S502',
    'jdsportspl': 'S512', 'jdsportsro': 'RO55', 'jdsportssk': 'SK52',
    'jdsportshu': 'HU52', 'jdsportslt': 'LT52', 'jdsportsbg': 'BG52',
    'jdsportscz': 'CZ55', 'jdsportsua': 'UA52', 'jdsportshr': 'HR52',
    'jdsportssi': 'SI52', 'jdsportsee': 'EE52',
}

SHOP_DICT   = {name: url for name, url in st.secrets["shop_urls"].items()}
MPK_TO_SHOP = {SHOP_TO_MPK.get(s, s): s for s in SHOP_DICT.keys()}

# MPK -> GA4 property ID mapping
GA4_PROPERTIES = {
    mpk: vals[0]
    for mpk, vals in st.secrets["ga4_properties"].items()
}

# ────────────────────────────────────────────────────────────
# GA4 CLIENT
# ────────────────────────────────────────────────────────────

@st.cache_resource
def get_ga4_client():
    credentials = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/analytics.readonly"]
    )
    return BetaAnalyticsDataClient(credentials=credentials)

try:
    ga4_client = get_ga4_client()
except Exception as e:
    st.warning(f"Nie udało się połączyć z Google Analytics: {e}")
    ga4_client = None

# ────────────────────────────────────────────────────────────
# HELPERS – PRICE APP
# ────────────────────────────────────────────────────────────

def get_mpk_code(shop_name):
    return SHOP_TO_MPK.get(shop_name, shop_name)


def load_csv(url):
    resp = requests.get(url, auth=HTTPBasicAuth(HTTP_USERNAME, HTTP_PASSWORD))
    resp.raise_for_status()
    raw = resp.content
    for enc in ('utf-8-sig', 'utf-8', 'cp1250', 'iso-8859-2', 'latin-1'):
        try:
            text = raw.decode(enc)
            sep  = ';' if ';' in text.splitlines()[0] else ','
            df   = pd.read_csv(StringIO(text), sep=sep, on_bad_lines='skip')
            return df
        except (UnicodeDecodeError, Exception):
            continue
    raise ValueError("Nie można odczytać pliku CSV — nieznany encoding")


def extract_id_from_url(url):
    try:
        slug  = str(url).rstrip('/').split('/')[-1]
        if len(slug) < 3:
            slug = str(url).rstrip('/').split('/')[-2]
        parts = slug.split('-')
        last_digit_idx = -1
        for i in range(len(parts) - 1, -1, -1):
            if re.search(r'\d', parts[i]):
                last_digit_idx = i
                break
        if last_digit_idx == -1:
            return ''
        return '-'.join(parts[last_digit_idx:]).upper()
    except:
        return ''


def extract_name_from_url(url, product_id):
    try:
        slug = str(url).rstrip('/').split('/')[-1]
        if len(slug) < 3:
            slug = str(url).rstrip('/').split('/')[-2]

        pid = str(product_id).lower().strip()
        slug_lower = slug.lower()

        if pid and pid in slug_lower:
            idx = slug_lower.rfind(pid)
            name_part = slug[:idx].rstrip('-')
        else:
            parts = slug.split('-')
            last_digit_idx = -1
            for i in range(len(parts) - 1, -1, -1):
                if re.search(r'\d', parts[i]):
                    last_digit_idx = i
                    break
            if last_digit_idx > 0:
                name_part = '-'.join(parts[:last_digit_idx])
            else:
                name_part = slug

        return name_part.replace('-', ' ').title().strip()
    except:
        return ''


def count_sizes(val):
    try:
        if pd.isna(val) or str(val).strip() == '':
            return 0
        return len(str(val).split('|'))
    except:
        return 0


def pct_diff(a, b):
    if b == 0:
        return None
    return round((a - b) / b * 100, 2)


def color_diff(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ''
    if pd.isna(v) or v == 0:
        return 'color: gray'
    return 'color: red' if v > 0 else 'color: green'


def color_diff_inverted(val):
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ''
    if pd.isna(v) or v == 0:
        return 'color: gray'
    return 'color: green' if v > 0 else 'color: red'


# ────────────────────────────────────────────────────────────
# HELPERS – GA4
# ────────────────────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_ga4_items(property_id: str, start_date: str, end_date: str) -> pd.DataFrame:
    """
    Pobiera itemsViewed i itemRevenue z GA4 dla zakresu dat.
    Wymiar: itemId.
    Zwraca DataFrame z kolumnami: itemId, itemsViewed, itemRevenue.
    """
    if ga4_client is None:
        return pd.DataFrame()
    try:
        request = RunReportRequest(
            property=f"properties/{property_id}",
            dimensions=[Dimension(name="itemId")],
            metrics=[
                Metric(name="itemsViewed"),
                Metric(name="itemRevenue"),
            ],
            date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
        )
        response = ga4_client.run_report(request)
        rows = []
        for row in response.rows:
            rows.append({
                "itemId":       row.dimension_values[0].value,
                "itemsViewed":  float(row.metric_values[0].value),
                "itemRevenue":  float(row.metric_values[1].value),
            })
        return pd.DataFrame(rows)
    except Exception as e:
        st.warning(f"GA4 błąd dla property {property_id}: {e}")
        return pd.DataFrame()


def build_ga4_for_mpk(mpk_code: str) -> pd.DataFrame:
    """
    Dla danego MPK pobiera dane GA4 za 7 i 30 dni.
    Zwraca DataFrame z kolumnami:
      itemId,
      itemsViewed_7d, itemRevenue_7d,
      itemsViewed_30d, itemRevenue_30d,
      itemsViewed_Diff, itemsViewed_Diff_%,
      itemRevenue_Diff, itemRevenue_Diff_%
    """
    property_id = GA4_PROPERTIES.get(mpk_code)
    if not property_id:
        return pd.DataFrame()

    today     = date.today()
    yesterday = today - timedelta(days=1)

    end_str   = yesterday.strftime('%Y-%m-%d')
    start_7d  = (yesterday - timedelta(days=6)).strftime('%Y-%m-%d')
    start_30d = (yesterday - timedelta(days=29)).strftime('%Y-%m-%d')

    df7  = fetch_ga4_items(property_id, start_7d,  end_str)
    df30 = fetch_ga4_items(property_id, start_30d, end_str)

    if df7.empty and df30.empty:
        return pd.DataFrame()

    rename7  = {"itemsViewed": "itemsViewed_7d",  "itemRevenue": "itemRevenue_7d"}
    rename30 = {"itemsViewed": "itemsViewed_30d", "itemRevenue": "itemRevenue_30d"}

    if not df7.empty:
        df7  = df7.rename(columns=rename7)
    if not df30.empty:
        df30 = df30.rename(columns=rename30)

    if df7.empty:
        merged = df30
        for c in ["itemsViewed_7d", "itemRevenue_7d"]:
            merged[c] = 0.0
    elif df30.empty:
        merged = df7
        for c in ["itemsViewed_30d", "itemRevenue_30d"]:
            merged[c] = 0.0
    else:
        merged = pd.merge(df7, df30, on="itemId", how="outer").fillna(0)

    # Oblicz różnicę 7d vs 30d
    for metric in ["itemsViewed", "itemRevenue"]:
        c7  = f"{metric}_7d"
        c30 = f"{metric}_30d"
        merged[f"{metric}_Diff"]   = (merged[c7] - merged[c30]).round(2)
        merged[f"{metric}_Diff_%"] = merged.apply(
            lambda r, c7=c7, c30=c30: pct_diff(r[c7], r[c30]), axis=1
        )

    merged["itemId"] = merged["itemId"].astype(str).str.strip().str.upper()
    return merged


# ────────────────────────────────────────────────────────────
# WYBÓR SKLEPÓW (max 2)
# ────────────────────────────────────────────────────────────

all_mpk_codes = list(MPK_TO_SHOP.keys())

selected_mpk_codes = st.multiselect(
    "Wybierz sklepy do porównania (maksymalnie 2)",
    all_mpk_codes,
    default=[],
    max_selections=2,
)
selected_shops = [MPK_TO_SHOP[m] for m in selected_mpk_codes]

if not selected_shops:
    st.info("👆 Wybierz przynajmniej jeden sklep, aby rozpocząć")
    st.stop()

# ────────────────────────────────────────────────────────────
# ORIENTACJA (tylko przy 2 sklepach)
# ────────────────────────────────────────────────────────────

mpk1 = mpk2 = None

if len(selected_shops) == 2:
    mpk_a = get_mpk_code(selected_shops[0])
    mpk_b = get_mpk_code(selected_shops[1])

    st.markdown("#### Porównanie")
    orientation_label = st.radio(
        "Kierunek porównania:",
        options=[
            f"{mpk_a} → {mpk_b}",
            f"{mpk_b} → {mpk_a}",
        ],
        horizontal=True,
        key="orientation",
    )

    if orientation_label.startswith(mpk_a):
        mpk1, mpk2 = mpk_a, mpk_b
        shop1, shop2 = selected_shops[0], selected_shops[1]
    else:
        mpk1, mpk2 = mpk_b, mpk_a
        shop1, shop2 = selected_shops[1], selected_shops[0]

# ────────────────────────────────────────────────────────────
# WCZYTANIE DANYCH CSV
# ────────────────────────────────────────────────────────────

shop_data = {}
for shop_name in selected_shops:
    mpk_code = get_mpk_code(shop_name)
    with st.spinner(f'Wczytuję dane z {mpk_code}...'):
        df = load_csv(SHOP_DICT[shop_name])

        for col in ['ID', 'Brand', 'Quantity', 'Variants', 'Sizes', 'CategoryName', 'Seasonality']:
            if col not in df.columns:
                df[col] = ''

        df['ID'] = df['ID'].astype(str).str.strip().str.upper()
        mask_empty = df['ID'].isin(['', 'NAN', 'NONE'])
        df.loc[mask_empty, 'ID'] = df.loc[mask_empty, 'URL'].apply(extract_id_from_url)

        df['Variants']    = pd.to_numeric(df['Variants'],  errors='coerce').fillna(0)
        df['Quantity']    = pd.to_numeric(df['Quantity'],   errors='coerce').fillna(0)
        df['SizesCount']  = df['Sizes'].apply(count_sizes)
        df['MPK']         = mpk_code

        df['ProductName'] = df.apply(
            lambda row: extract_name_from_url(row.get('URL', ''), row['ID']),
            axis=1
        )

        shop_data[shop_name] = df

# ────────────────────────────────────────────────────────────
# WCZYTANIE DANYCH GA4
# ────────────────────────────────────────────────────────────

ga4_data = {}
for shop_name in selected_shops:
    mpk_code = get_mpk_code(shop_name)
    with st.spinner(f'Wczytuję dane GA4 dla {mpk_code}...'):
        ga4_df = build_ga4_for_mpk(mpk_code)
        if not ga4_df.empty:
            ga4_data[mpk_code] = ga4_df

# ────────────────────────────────────────────────────────────
# SCALENIE GA4 Z DANYMI CSV
# Łączymy najpierw po Index, potem po ID (fallback)
# ────────────────────────────────────────────────────────────

def merge_with_ga4(df: pd.DataFrame, mpk_code: str) -> pd.DataFrame:
    """Dołącza kolumny GA4 do df cenowego dla danego MPK."""
    if mpk_code not in ga4_data:
        # Dodaj puste kolumny żeby tabela była spójna
        for c in ["itemsViewed_7d", "itemRevenue_7d",
                  "itemsViewed_30d", "itemRevenue_30d",
                  "itemsViewed_Diff", "itemsViewed_Diff_%",
                  "itemRevenue_Diff", "itemRevenue_Diff_%"]:
            df[c] = None
        return df

    ga4 = ga4_data[mpk_code].copy()

    # Próba 1: łącz po Index
    if 'Index' in df.columns:
        df_idx = df.copy()
        df_idx['_join_key'] = df_idx['Index'].astype(str).str.strip().str.upper()
        ga4['_join_key']    = ga4['itemId']
        merged = pd.merge(df_idx, ga4.drop(columns=['itemId']), on='_join_key', how='left')
        matched = merged[ga4.drop(columns=['itemId']).columns[0]].notna().sum() if len(ga4.drop(columns=['itemId']).columns) > 0 else 0
        merged.drop(columns=['_join_key'], inplace=True)

        # Jeśli Index dał ≥1 trafienie — OK
        ga4_metric_cols = [c for c in ga4.columns if c != 'itemId']
        hit_count = merged[ga4_metric_cols[0]].notna().sum() if ga4_metric_cols else 0

        if hit_count > 0:
            return merged

    # Próba 2: łącz po ID
    df_id = df.copy()
    df_id['_join_key'] = df_id['ID'].astype(str).str.strip().str.upper()
    ga4['_join_key']   = ga4['itemId']
    merged = pd.merge(df_id, ga4.drop(columns=['itemId']), on='_join_key', how='left')
    merged.drop(columns=['_join_key'], inplace=True)
    return merged


# Zastosuj GA4 merge dla każdego sklepu
for shop_name in selected_shops:
    mpk_code = get_mpk_code(shop_name)
    shop_data[shop_name] = merge_with_ga4(shop_data[shop_name], mpk_code)

# ────────────────────────────────────────────────────────────
# BUDOWANIE TABELI WYNIKOWEJ
# ────────────────────────────────────────────────────────────

GA4_COLS = [
    "itemsViewed_7d", "itemRevenue_7d",
    "itemsViewed_30d", "itemRevenue_30d",
    "itemsViewed_Diff", "itemsViewed_Diff_%",
    "itemRevenue_Diff", "itemRevenue_Diff_%",
]

INFO_COLS = ['Index', 'ID', 'ProductName', 'Brand', 'CategoryName', 'Seasonality']

if len(selected_shops) == 1:
    sn = selected_shops[0]
    df = shop_data[sn]
    base_cols = INFO_COLS + ['Price', 'SizesCount', 'Variants', 'Quantity', 'MPK']
    ga4_present = [c for c in GA4_COLS if c in df.columns]
    result_final = df[base_cols + ga4_present].copy()

else:  # 2 sklepy
    st.session_state[f'len_{mpk1}'] = len(shop_data[shop1])
    st.session_state[f'len_{mpk2}'] = len(shop_data[shop2])
    df1, df2 = shop_data[shop1], shop_data[shop2]

    merged = pd.merge(df1, df2, on='Index', suffixes=(f'_{mpk1}', f'_{mpk2}'), how='inner')

    if merged.empty:
        st.info("Brak wspólnych po Index — próbuję po ID...")
        merged = pd.merge(
            df1.rename(columns={'Index': f'Index_{mpk1}'}),
            df2.rename(columns={'Index': f'Index_{mpk2}'}),
            on='ID', suffixes=(f'_{mpk1}', f'_{mpk2}'), how='inner'
        )
        if merged.empty:
            st.warning("Brak wspólnych produktów po Index ani ID między wybranymi sklepami.")
            st.stop()
        merged['Index'] = merged[f'Index_{mpk1}'].combine_first(merged[f'Index_{mpk2}'])

    for metric in ['Price', 'Variants', 'Quantity', 'SizesCount']:
        merged[f'{metric}_Diff']     = merged[f'{metric}_{mpk1}'] - merged[f'{metric}_{mpk2}']
        merged[f'{metric}_Diff_Pct'] = merged.apply(
            lambda r, m=metric: pct_diff(r[f'{m}_{mpk1}'], r[f'{m}_{mpk2}']), axis=1)

    def pick(col):
        c1 = f'{col}_{mpk1}'
        if c1 in merged.columns:
            return merged[c1]
        elif col in merged.columns:
            return merged[col]
        return ''

    if f'ID_{mpk1}' in merged.columns and f'ID_{mpk2}' in merged.columns:
        id_val = merged[f'ID_{mpk1}'].replace('', pd.NA).combine_first(merged[f'ID_{mpk2}'])
    elif 'ID' in merged.columns:
        id_val = merged['ID']
    else:
        id_val = pick('ID')

    # Bazowe kolumny cenowe
    result_dict = {
        'Index':              merged['Index'],
        'ID':                 id_val,
        'ProductName':        pick('ProductName'),
        'Brand':              pick('Brand'),
        'CategoryName':       pick('CategoryName'),
        'Seasonality':        pick('Seasonality'),
        f'Price_{mpk1}':      merged[f'Price_{mpk1}'],
        f'Price_{mpk2}':      merged[f'Price_{mpk2}'],
        'Price_Diff':         merged['Price_Diff'].round(2),
        'Price_Diff_%':       merged['Price_Diff_Pct'],
        f'SizesCount_{mpk1}': merged[f'SizesCount_{mpk1}'],
        f'SizesCount_{mpk2}': merged[f'SizesCount_{mpk2}'],
        'SizesCount_Diff':    merged['SizesCount_Diff'],
        'SizesCount_Diff_%':  merged['SizesCount_Diff_Pct'],
        f'Variants_{mpk1}':   merged[f'Variants_{mpk1}'],
        f'Variants_{mpk2}':   merged[f'Variants_{mpk2}'],
        'Variants_Diff':      merged['Variants_Diff'],
        'Variants_Diff_%':    merged['Variants_Diff_Pct'],
        f'Quantity_{mpk1}':   merged[f'Quantity_{mpk1}'],
        f'Quantity_{mpk2}':   merged[f'Quantity_{mpk2}'],
        'Quantity_Diff':      merged['Quantity_Diff'],
        'Quantity_Diff_%':    merged['Quantity_Diff_Pct'],
    }

    # GA4 kolumny — dla trybu 2-sklepowego bierzemy dane z mpk1 (sklep bazowy)
    # GA4 kolumny mają sufiks _{mpk1} po merge, więc je wyciągamy
    for ga4c in GA4_COLS:
        col_with_suffix = f'{ga4c}_{mpk1}'
        if col_with_suffix in merged.columns:
            result_dict[ga4c] = merged[col_with_suffix]
        elif ga4c in merged.columns:
            result_dict[ga4c] = merged[ga4c]
        # jeśli brak — kolumna nie trafia do wyniku (nie ma GA4 dla tego MPK)

    result_final = pd.DataFrame(result_dict)

# ────────────────────────────────────────────────────────────
# FILTRY — aktywacja TYLKO po przycisku „Filtruj"
# ────────────────────────────────────────────────────────────
skip_filter  = ['Index']
text_cols    = [c for c in result_final.columns
                if c not in skip_filter and not pd.api.types.is_numeric_dtype(result_final[c])]
numeric_cols = [c for c in result_final.columns
                if c not in skip_filter and pd.api.types.is_numeric_dtype(result_final[c])]
all_columns  = text_cols + numeric_cols

if 'applied_filters' not in st.session_state:
    st.session_state['applied_filters'] = {}
if 'filter_reset_counter' not in st.session_state:
    st.session_state['filter_reset_counter'] = 0

rc = st.session_state['filter_reset_counter']

st.markdown("---")

active = 0
for cn, fv in st.session_state['applied_filters'].items():
    if cn in text_cols and fv:
        active += 1
    elif cn in numeric_cols and fv:
        cd = result_final[cn].dropna()
        if len(cd) > 0:
            full_range = (float(cd.min()), float(cd.max()))
            if tuple(fv) != full_range:
                active += 1

label = f"🔍 Filtry danych{' — ✅ ' + str(active) + ' aktywnych' if active else ''}"
with st.expander(label, expanded=False):
    with st.form(key=f"filter_form_{rc}"):
        for i in range(0, len(all_columns), 4):
            cols = st.columns(4)
            for idx, cn in enumerate(all_columns[i:i+4]):
                with cols[idx]:
                    with st.expander(f"🔽 {cn}", expanded=False):
                        if cn in text_cols:
                            all_vals = sorted(result_final[cn].dropna().astype(str).unique())
                            current_sel = st.session_state['applied_filters'].get(cn, [])
                            if isinstance(current_sel, set):
                                current_sel = []
                            st.multiselect(
                                "Wartości", options=all_vals,
                                default=[v for v in current_sel if v in all_vals],
                                key=f"form_multi_{cn}_{rc}"
                            )
                        else:
                            cd = result_final[cn].dropna()
                            if len(cd):
                                mn, mx = float(cd.min()), float(cd.max())
                                if mn != mx:
                                    cur = st.session_state['applied_filters'].get(cn, (mn, mx))
                                    st.slider(
                                        "Zakres", min_value=mn, max_value=mx,
                                        value=(max(mn, float(cur[0])), min(mx, float(cur[1]))),
                                        key=f"form_slider_{cn}_{rc}"
                                    )

        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            submitted = st.form_submit_button("🔍 Filtruj", use_container_width=True, type="primary")
        with btn_col2:
            reset = st.form_submit_button("🔄 Resetuj wszystkie filtry", use_container_width=True)

    if submitted:
        new_filters = {}
        for cn in all_columns:
            key_m = f"form_multi_{cn}_{rc}"
            key_s = f"form_slider_{cn}_{rc}"
            if key_m in st.session_state:
                new_filters[cn] = st.session_state[key_m]
            elif key_s in st.session_state:
                new_filters[cn] = st.session_state[key_s]
        st.session_state['applied_filters'] = new_filters
        st.rerun()

    if reset:
        st.session_state['applied_filters'] = {}
        st.session_state['filter_reset_counter'] += 1
        st.rerun()

filtered_df = result_final.copy()
for cn, fv in st.session_state['applied_filters'].items():
    if cn in text_cols and fv:
        filtered_df = filtered_df[filtered_df[cn].astype(str).isin(fv)]
    elif cn in numeric_cols and fv:
        filtered_df = filtered_df[(filtered_df[cn] >= fv[0]) & (filtered_df[cn] <= fv[1])]

# ────────────────────────────────────────────────────────────
# TABELA Z KOLOROWANIEM
# ────────────────────────────────────────────────────────────
if filtered_df is not None and not filtered_df.empty:
    st.markdown("---")
    st.subheader(f"Porównanie: {', '.join(selected_mpk_codes)}")
    st.caption(f"Wyświetlono {len(filtered_df)} z {len(result_final)} produktów")

    diff_cols = [c for c in filtered_df.columns if 'Diff' in c]
    inverted_keywords = ('SizesCount', 'Variants', 'Quantity', 'itemsViewed', 'itemRevenue')
    diff_inverted = [c for c in diff_cols if any(k in c for k in inverted_keywords)]
    diff_normal   = [c for c in diff_cols if c not in diff_inverted]

    format_rules = {}
    for col in filtered_df.columns:
        if pd.api.types.is_numeric_dtype(filtered_df[col]):
            if 'Price' in col and 'Diff' not in col:
                format_rules[col] = "{:.2f}"
            elif 'Price' in col and 'Diff' in col and '%' not in col:
                format_rules[col] = "{:+.2f}"
            elif col.endswith('%') or col.endswith('_%'):
                format_rules[col] = "{:+.1f}%"
            elif 'Diff' in col:
                format_rules[col] = "{:+.0f}"
            elif 'itemRevenue' in col:
                format_rules[col] = "{:.2f}"
            elif 'itemsViewed' in col:
                format_rules[col] = "{:.0f}"

    try:
        styled = filtered_df.style
        if diff_normal:
            styled = styled.map(color_diff, subset=diff_normal)
        if diff_inverted:
            styled = styled.map(color_diff_inverted, subset=diff_inverted)
    except AttributeError:
        styled = filtered_df.style
        if diff_normal:
            styled = styled.applymap(color_diff, subset=diff_normal)
        if diff_inverted:
            styled = styled.applymap(color_diff_inverted, subset=diff_inverted)
    styled = styled.format(format_rules, na_rep='—')

    pinned_config = {
        "Index":       st.column_config.Column(pinned=True),
        "ID":          st.column_config.Column(pinned=True),
        "ProductName": st.column_config.Column(pinned=True),
        "Brand":       st.column_config.Column(pinned=True),
        "CategoryName":st.column_config.Column(pinned=True),
        "Seasonality": st.column_config.Column(pinned=True),
    }

    st.dataframe(
        styled,
        use_container_width=True,
        height=520,
        column_config=pinned_config
    )

    # ────────────────────────────────────────────────────────
    # PODSUMOWANIE (tylko 2 sklepy)
    # ────────────────────────────────────────────────────────
    if len(selected_shops) == 2 and mpk1 and mpk2 and 'Price_Diff' in filtered_df.columns:
        st.markdown("---")
        st.markdown("### 📊 Podsumowanie wspólnych produktów")

        total     = len(filtered_df)
        orig_len1 = st.session_state.get(f'len_{mpk1}', total)
        orig_len2 = st.session_state.get(f'len_{mpk2}', total)
        pct_of_1  = round(total / orig_len1 * 100, 2) if orig_len1 else 0
        pct_of_2  = round(total / orig_len2 * 100, 2) if orig_len2 else 0

        cheaper_mask = filtered_df['Price_Diff'] < 0
        dearer_mask  = filtered_df['Price_Diff'] > 0
        equal_mask   = filtered_df['Price_Diff'] == 0

        cheaper = int(cheaper_mask.sum())
        dearer  = int(dearer_mask.sum())
        equal   = int(equal_mask.sum())

        pct_c = round(cheaper / total * 100, 1) if total else 0
        pct_d = round(dearer  / total * 100, 1) if total else 0
        pct_e = round(equal   / total * 100, 1) if total else 0

        avg_diff_c = filtered_df.loc[cheaper_mask, 'Price_Diff'].mean() if cheaper else 0
        avg_diff_d = filtered_df.loc[dearer_mask,  'Price_Diff'].mean() if dearer  else 0

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Wspólnych produktów", total)
            st.caption(f"({pct_of_1}% {mpk1}, {pct_of_2}% {mpk2})")
        with c2:
            st.metric(f"🟢 {mpk1} tańszy niż {mpk2}", cheaper)
            st.caption(f"{pct_c}% produktów\n\nŚr. taniej o: {avg_diff_c:.2f}")
        with c3:
            st.metric(f"🔴 {mpk1} droższy niż {mpk2}", dearer)
            st.caption(f"{pct_d}% produktów\n\nŚr. drożej o: +{avg_diff_d:.2f}")
        with c4:
            st.metric("⚪ Równe ceny", equal)
            st.caption(f"{pct_e}% produktów")

    # ────────────────────────────────────────────────────────
    # POBIERANIE – XLSX z datą
    # ────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📥 Pobierz dane")

    now_str = datetime.now().strftime('%Y-%m-%d_%H-%M')
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Porównanie')
    buffer.seek(0)

    st.download_button(
        label="📥 Pobierz XLSX",
        data=buffer,
        file_name=f"porownanie_{'_'.join(selected_mpk_codes)}_{now_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.warning("Brak danych po zastosowaniu filtrów. Spróbuj zmienić kryteria filtrowania.")
