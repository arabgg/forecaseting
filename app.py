"""
=====================================================================
DASHBOARD FORECASTING PERSEDIAAN BAHAN BAKU UMKM F&B
Studi Kasus: Mitra XYZ - Kendalsari
Metode: Moving Average (MA) & Single Exponential Smoothing (SES)
Sifat: Serverless (tanpa database), membaca langsung file Excel/CSV
=====================================================================
Cara menjalankan:
    pip install streamlit pandas plotly openpyxl numpy
    streamlit run app.py

File data yang dibaca otomatis (jika ada di folder yang sama):
    - SO_Harian_Kendalsari_Updated.xlsx   (sheet bulan: November 2025, Desember 2025, Januari 2026, dst)
    - SO_Harian_Kendalsari__1_.xlsx       (sheet bulan: Maret 2026, April 2026, ... + List Barang, FlowBarang)

Jika file tidak ditemukan, dashboard otomatis membuat MOCK DATA agar tetap
bisa didemokan tanpa data asli.
=====================================================================
"""

import io
import os
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

# =====================================================================
# 0. KONFIGURASI HALAMAN
# =====================================================================
st.set_page_config(
    page_title="Dashboard Forecasting Persediaan - Mitra XYZ",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

PRIMARY_COLOR = "#1F77B4"      # Aktual
MA_COLOR = "#FF7F0E"           # Moving Average
SES_COLOR = "#2CA02C"          # SES
DANGER_COLOR = "#D62728"
SAFE_COLOR = "#2CA02C"

DEFAULT_FILE_1 = "SO_Harian_Kendalsari_Updated.xlsx"   # Nov 2025 - Jan 2026
DEFAULT_FILE_2 = "SO_Harian_Kendalsari__1_.xlsx"       # Maret 2026 - Agustus 2026 + sheet pendukung

NON_MONTH_SHEETS = {
    "list barang", "flowbarang", "dashboard harian",
    # Catatan: sheet di bawah ini terdeteksi sebagai TEMPLATE/PLACEHOLDER kosong
    # (isi datanya identik 100% antara 'Juli 2026' dan 'Agustus 2026' -> belum diisi data riil).
    # Sesuai data aktual penelitian yang hanya tersedia s.d. Juni 2026, kedua sheet ini dikecualikan.
    "juli 2026", "agustus 2026",
}

BULAN_ID = {
    "januari": 1, "februari": 2, "maret": 3, "april": 4, "mei": 5, "juni": 6,
    "juli": 7, "agustus": 8, "september": 9, "oktober": 10, "november": 11, "desember": 12,
}


def _expected_month_year(sheet_name: str):
    """Mengekstrak (bulan, tahun) yang seharusnya dari nama sheet, misal 'Januari 2026' -> (1, 2026)."""
    parts = str(sheet_name).strip().lower().split()
    bulan, tahun = None, None
    for p in parts:
        if p in BULAN_ID:
            bulan = BULAN_ID[p]
        elif p.isdigit() and len(p) == 4:
            tahun = int(p)
    return bulan, tahun

ITEM_PRIORITAS = [
    "Ayam Paprika", "Beras", "Cabe Rawit Merah", "Cabe Rawit Ijo",
    "Bawang Merah", "Bawang Putih",
]


# =====================================================================
# 1. MOCK DATA GENERATOR (agar script langsung bisa dijalankan tanpa file asli)
# =====================================================================
def generate_mock_data() -> pd.DataFrame:
    """
    Membuat data dummy time series harian pemakaian bahan baku,
    meniru struktur data SO Harian Kendalsari, untuk periode
    November 2025 s.d. Juli 2026 (dengan gap Februari, sesuai data riil).
    """
    rng = np.random.default_rng(42)

    bulan_tersedia = [
        ("2025-11-01", "2025-11-30"),
        ("2025-12-01", "2025-12-31"),
        ("2026-01-01", "2026-01-31"),
        ("2026-03-01", "2026-03-31"),
        ("2026-04-01", "2026-04-30"),
        ("2026-05-01", "2026-05-31"),
        ("2026-06-01", "2026-06-30"),
    ]

    item_config = {
        "Ayam Paprika":      {"base": 25, "noise": 8,  "satuan": "pack", "batas": 15, "harga": 38000},
        "Beras":             {"base": 12, "noise": 4,  "satuan": "kg",   "batas": 10, "harga": 14000},
        "Cabe Rawit Merah":  {"base": 5,  "noise": 2.5,"satuan": "kg",   "batas": 4,  "harga": 55000},
        "Cabe Rawit Ijo":    {"base": 4,  "noise": 2,  "satuan": "kg",   "batas": 3,  "harga": 45000},
        "Bawang Merah":      {"base": 6,  "noise": 2,  "satuan": "kg",   "batas": 5,  "harga": 38000},
        "Bawang Putih":      {"base": 4,  "noise": 1.5,"satuan": "kg",   "batas": 3,  "harga": 36000},
    }

    rows = []
    for item, cfg in item_config.items():
        for start, end in bulan_tersedia:
            dates = pd.date_range(start, end, freq="D")
            # tren musiman ringan + noise + efek weekend (Jumat-Minggu lebih ramai)
            trend = np.linspace(0, rng.uniform(-2, 2), len(dates))
            weekend_boost = np.array([1.3 if d.dayofweek in [4, 5, 6] else 1.0 for d in dates])
            noise = rng.normal(0, cfg["noise"], len(dates))
            pemakaian = (cfg["base"] + trend) * weekend_boost + noise
            pemakaian = np.clip(pemakaian, 0, None).round(1)

            for d, p in zip(dates, pemakaian):
                rows.append({
                    "Tanggal": d,
                    "ITEM": item,
                    "Satuan": cfg["satuan"],
                    "Pemakaian": p,
                    "Batas": cfg["batas"],
                })

    df = pd.DataFrame(rows)
    # Sisa stok hari ini disimulasikan dari pemakaian terakhir
    sisa_map = {item: max(cfg["batas"] * rng.uniform(0.6, 2.2), 1) for item, cfg in item_config.items()}
    df["SisaHariIni"] = df["ITEM"].map(sisa_map).round(1)
    return df


# =====================================================================
# 2. LOADER DATA ASLI (EXCEL / CSV) - PREPROCESSING DENGAN PANDAS
# =====================================================================
def _parse_month_sheet(df_raw: pd.DataFrame, sheet_name: str = "") -> pd.DataFrame:
    """
    Mengubah satu sheet bulanan (format SO Harian) menjadi long-format:
    Tanggal | ITEM | Satuan | Pemakaian | Batas | SisaHariIni
    Struktur kolom tetap (berdasarkan posisi):
      0: NO | 1: ITEM | 2: Satuan | 3: Batas | 4: Sisa Barang/Hari ini | 5..: tanggal-tanggal

    Catatan: beberapa file sumber memiliki anomali penulisan header tanggal
    (contoh: pada sheet 'Januari 2026' ditemukan 1 header bertuliskan '2025-01-05'
    yang seharusnya '2026-01-05'). Fungsi ini mengoreksi anomali tersebut dengan
    menyamakan bulan & tahun header ke bulan/tahun sheet, selama tanggal (hari)
    masih valid. Jika tidak bisa dikoreksi, kolom tersebut dilewati.
    """
    if df_raw.shape[1] < 6:
        return pd.DataFrame()

    cols = list(df_raw.columns)
    item_col, satuan_col, batas_col, sisa_col = cols[1], cols[2], cols[3], cols[4]
    date_cols = cols[5:]

    expected_month, expected_year = _expected_month_year(sheet_name)

    # Pra-proses: validasi & koreksi setiap header tanggal sekali saja
    valid_date_cols = []
    for dc in date_cols:
        tgl = pd.to_datetime(dc, errors="coerce")
        if pd.isna(tgl):
            continue
        if expected_month and expected_year and (tgl.month != expected_month or tgl.year != expected_year):
            try:
                tgl = tgl.replace(year=expected_year, month=expected_month)
            except ValueError:
                continue  # tanggal tidak valid untuk bulan tujuan (misal 31 di bulan 30 hari) -> skip
        valid_date_cols.append((dc, tgl))

    records = []
    for _, row in df_raw.iterrows():
        item = row[item_col]
        if pd.isna(item) or str(item).strip() == "" or str(item).strip().upper() == "PIC (PENANGGUNG JAWAB)":
            continue
        item = str(item).strip()
        satuan = row[satuan_col] if satuan_col in row else None
        batas = pd.to_numeric(row[batas_col], errors="coerce")
        sisa = pd.to_numeric(row[sisa_col], errors="coerce")

        for dc_original, tgl_corrected in valid_date_cols:
            val = pd.to_numeric(row[dc_original], errors="coerce")
            records.append({
                "Tanggal": tgl_corrected,
                "ITEM": item,
                "Satuan": satuan,
                "Pemakaian": val,
                "Batas": batas,
                "SisaHariIni": sisa,
            })

    return pd.DataFrame(records)


@st.cache_data(show_spinner=False)
def load_excel_sources(file_bytes_list) -> pd.DataFrame:
    """
    Membaca satu atau lebih file excel (bytes), mengambil seluruh sheet
    bulanan yang valid (punya kolom ITEM), lalu menggabungkannya menjadi
    satu dataframe time-series panjang (long format).
    """
    all_frames = []
    for fbytes in file_bytes_list:
        try:
            xls = pd.ExcelFile(io.BytesIO(fbytes))
        except Exception:
            continue
        for sheet_name in xls.sheet_names:
            if sheet_name.strip().lower() in NON_MONTH_SHEETS:
                continue
            try:
                df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=1)
            except Exception:
                continue
            if "ITEM" not in [str(c).strip() for c in df_raw.columns]:
                continue
            parsed = _parse_month_sheet(df_raw, sheet_name=sheet_name)
            if not parsed.empty:
                all_frames.append(parsed)

    if not all_frames:
        return pd.DataFrame()

    df = pd.concat(all_frames, ignore_index=True)
    return clean_dataset(df)


def clean_dataset(df: pd.DataFrame) -> pd.DataFrame:
    """
    PREPROCESSING (sesuai ketentuan skripsi):
    - Hilangkan duplikat (item & tanggal sama)
    - Missing value pemakaian -> 0
    - Urutkan berdasarkan tanggal
    - Interpolasi tanggal yang kosong di tengah rentang historis (per item)
    """
    df = df.dropna(subset=["Tanggal", "ITEM"]).copy()
    df = df.drop_duplicates(subset=["Tanggal", "ITEM"], keep="last")
    df["Pemakaian"] = pd.to_numeric(df["Pemakaian"], errors="coerce").fillna(0)
    df = df.sort_values(["ITEM", "Tanggal"]).reset_index(drop=True)

    # Interpolasi tanggal kosong (gap) per item agar deret waktu harian utuh
    filled_frames = []
    for item, g in df.groupby("ITEM"):
        g = g.set_index("Tanggal").sort_index()
        full_idx = pd.date_range(g.index.min(), g.index.max(), freq="D")
        g = g.reindex(full_idx)
        g["ITEM"] = item
        g["Satuan"] = g["Satuan"].ffill().bfill()
        g["Batas"] = g["Batas"].ffill().bfill()
        g["SisaHariIni"] = g["SisaHariIni"].ffill().bfill()
        # Pemakaian yang NaN (hasil reindex/gap tanggal) -> interpolasi linear, lalu sisa NaN -> 0
        g["Pemakaian"] = g["Pemakaian"].interpolate(method="linear").fillna(0)
        g.index.name = "Tanggal"
        filled_frames.append(g.reset_index())

    return pd.concat(filled_frames, ignore_index=True)


def get_item_master_info(df: pd.DataFrame) -> pd.DataFrame:
    """Mengambil info terkini (Batas & Sisa Stok) per item, dari data tanggal paling baru."""
    latest = df.sort_values("Tanggal").groupby("ITEM").tail(1)
    return latest[["ITEM", "Satuan", "Batas", "SisaHariIni"]].reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_default_or_mock():
    """Coba baca file default di folder lokal. Jika tidak ada, pakai mock data."""
    file_bytes = []
    for fname in [DEFAULT_FILE_1, DEFAULT_FILE_2]:
        if os.path.exists(fname):
            with open(fname, "rb") as f:
                file_bytes.append(f.read())

    if file_bytes:
        df = load_excel_sources(file_bytes)
        if not df.empty:
            return df, "FILE_ASLI"

    return generate_mock_data(), "MOCK_DATA"


# =====================================================================
# 3. FUNGSI PERHITUNGAN METODE FORECASTING (MODULAR)
# =====================================================================
def calc_moving_average(series: pd.Series, n: int) -> pd.Series:
    """
    Forecast Moving Average: forecast(t) = rata-rata aktual (t-n) s/d (t-1)
    """
    return series.rolling(window=n).mean().shift(1)


def calc_moving_average_next(series: pd.Series, n: int) -> float:
    """Forecast t+1 (hari esok) menggunakan rata-rata n data aktual terakhir."""
    if len(series) < n:
        return float(series.mean())
    return float(series.iloc[-n:].mean())


def calc_ses(series: pd.Series, alpha: float) -> pd.Series:
    """
    Forecast Single Exponential Smoothing:
    F(1) = A(1)
    F(t) = alpha * A(t-1) + (1-alpha) * F(t-1)
    """
    values = series.values
    forecast = np.empty(len(values))
    forecast[:] = np.nan
    if len(values) == 0:
        return pd.Series(forecast, index=series.index)

    forecast[0] = values[0]  # inisialisasi, bukan forecast riil
    for t in range(1, len(values)):
        forecast[t] = alpha * values[t - 1] + (1 - alpha) * forecast[t - 1]

    result = pd.Series(forecast, index=series.index)
    result.iloc[0] = np.nan  # titik pertama bukan hasil forecast, dikosongkan untuk evaluasi error
    return result


def calc_ses_next(series: pd.Series, alpha: float) -> float:
    """Forecast t+1 (hari esok) berdasarkan rekursi SES penuh hingga data terakhir."""
    values = series.values
    f_prev = values[0]
    for t in range(1, len(values)):
        f_prev = alpha * values[t - 1] + (1 - alpha) * f_prev
    f_next = alpha * values[-1] + (1 - alpha) * f_prev
    return float(f_next)


def calc_mad(actual: pd.Series, forecast: pd.Series) -> float:
    valid = actual.notna() & forecast.notna()
    if valid.sum() == 0:
        return np.nan
    return float((actual[valid] - forecast[valid]).abs().mean())


def calc_mse(actual: pd.Series, forecast: pd.Series) -> float:
    valid = actual.notna() & forecast.notna()
    if valid.sum() == 0:
        return np.nan
    return float(((actual[valid] - forecast[valid]) ** 2).mean())


def calc_mape(actual: pd.Series, forecast: pd.Series) -> float:
    valid = actual.notna() & forecast.notna() & (actual != 0)
    if valid.sum() == 0:
        return np.nan
    return float((((actual[valid] - forecast[valid]).abs()) / actual[valid]).mean() * 100)


def evaluate_method(actual: pd.Series, forecast: pd.Series) -> dict:
    return {
        "MAD": calc_mad(actual, forecast),
        "MSE": calc_mse(actual, forecast),
        "MAPE": calc_mape(actual, forecast),
    }


# =====================================================================
# 4. LOAD DATA
# =====================================================================
df_all, data_source = load_default_or_mock()

# =====================================================================
# 5. SIDEBAR - KONTROL UTAMA
# =====================================================================
with st.sidebar:
    st.image(
        "https://api.dicebear.com/7.x/shapes/svg?seed=MitraXYZ&backgroundColor=1F77B4",
        width=80,
    )
    st.title("📦 Mitra XYZ")
    st.caption("Dashboard Forecasting Persediaan Bahan Baku")
    st.markdown("---")

    st.subheader("📁 Sumber Data")
    if data_source == "MOCK_DATA":
        st.info("Menggunakan **mock data** demo (file asli tidak ditemukan di folder).")
    else:
        st.success("Data dimuat dari file Excel SO Harian.")

    uploaded_files = st.file_uploader(
        "Upload file Excel SO Harian (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Bisa upload lebih dari 1 file (misal data Nov-Jan dan Maret-Agustus).",
    )
    if uploaded_files:
        bytes_list = [f.read() for f in uploaded_files]
        df_uploaded = load_excel_sources(bytes_list)
        if not df_uploaded.empty:
            df_all = df_uploaded
            data_source = "FILE_UPLOAD"
            st.success(f"Berhasil memuat {len(uploaded_files)} file upload.")
        else:
            st.warning("Format file tidak sesuai, tetap menggunakan data sebelumnya.")

    st.markdown("---")
    st.subheader("🧂 Pilih Bahan Baku")

    item_list = sorted(df_all["ITEM"].dropna().unique().tolist())
    default_item = next((i for i in ITEM_PRIORITAS if i in item_list), item_list[0] if item_list else None)
    selected_item = st.selectbox("Jenis Bahan Baku", item_list, index=item_list.index(default_item) if default_item in item_list else 0)

    st.markdown("---")
    st.subheader("⚙️ Parameter Metode")

    ma_n = st.selectbox("Moving Average - Window (n)", options=[3, 5, 7], index=1)
    ses_alpha = st.slider("SES - Nilai Alpha (α)", min_value=0.1, max_value=0.9, value=0.3, step=0.1)

    st.markdown("---")
    st.subheader("📅 Filter Rentang Tanggal")

    item_df_full = df_all[df_all["ITEM"] == selected_item].sort_values("Tanggal")
    min_date = item_df_full["Tanggal"].min().date()
    max_date = item_df_full["Tanggal"].max().date()

    date_range = st.date_input(
        "Rentang Tanggal Historis",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date, end_date = min_date, max_date

    st.markdown("---")
    st.caption("Skripsi: Dashboard Forecasting Persediaan Bahan Baku UMKM F&B")
    st.caption("Metode: Moving Average vs Single Exponential Smoothing")


# =====================================================================
# 6. PERSIAPAN DATA TERFILTER UNTUK ITEM TERPILIH
# =====================================================================
mask = (item_df_full["Tanggal"].dt.date >= start_date) & (item_df_full["Tanggal"].dt.date <= end_date)
item_df = item_df_full.loc[mask].reset_index(drop=True)

if item_df.empty or len(item_df) < max(ma_n, 2):
    st.warning("Data historis pada rentang tanggal ini terlalu sedikit untuk dihitung. Silakan perlebar rentang tanggal.")
    st.stop()

actual_series = item_df.set_index("Tanggal")["Pemakaian"]

ma_forecast_series = calc_moving_average(actual_series, ma_n)
ses_forecast_series = calc_ses(actual_series, ses_alpha)

ma_eval = evaluate_method(actual_series, ma_forecast_series)
ses_eval = evaluate_method(actual_series, ses_forecast_series)

ma_next = calc_moving_average_next(actual_series, ma_n)
ses_next = calc_ses_next(actual_series, ses_alpha)

# Penentuan metode terbaik berdasarkan MAPE terkecil
if ma_eval["MAPE"] is not None and not np.isnan(ma_eval["MAPE"]) and (
    np.isnan(ses_eval["MAPE"]) or ma_eval["MAPE"] <= ses_eval["MAPE"]
):
    best_method_name = f"Moving Average (n={ma_n})"
    best_mape = ma_eval["MAPE"]
    best_next_forecast = ma_next
else:
    best_method_name = f"Single Exponential Smoothing (α={ses_alpha})"
    best_mape = ses_eval["MAPE"]
    best_next_forecast = ses_next

# Info stok terkini
sisa_stok = float(item_df["SisaHariIni"].iloc[-1]) if pd.notna(item_df["SisaHariIni"].iloc[-1]) else 0.0
batas_minimum = float(item_df["Batas"].iloc[-1]) if pd.notna(item_df["Batas"].iloc[-1]) else 0.0
satuan = item_df["Satuan"].iloc[-1] if pd.notna(item_df["Satuan"].iloc[-1]) else ""

status_stok = "🟢 Aman" if sisa_stok > batas_minimum else "🔴 Harus Order!!!"
rekomendasi_restock = max(best_next_forecast - sisa_stok, 0)


# =====================================================================
# 7. MAIN PAGE - HEADER
# =====================================================================
st.title("📊 Dashboard Forecasting Persediaan Bahan Baku")
st.caption(f"UMKM F&B - Mitra XYZ (Cabang Kendalsari) | Bahan Baku: **{selected_item}**")
st.markdown("---")


# =====================================================================
# 8. BAGIAN 1 - ROW KPI CARDS
# =====================================================================
st.subheader("📌 Ringkasan Stok Saat Ini")
kpi1, kpi2, kpi3 = st.columns(3)

with kpi1:
    st.metric(
        label=f"Sisa Stok Hari Ini ({satuan})",
        value=f"{sisa_stok:,.1f}",
    )

with kpi2:
    st.metric(
        label=f"Batas Minimum / Safety Stock ({satuan})",
        value=f"{batas_minimum:,.1f}",
    )

with kpi3:
    if "Harus Order" in status_stok:
        st.error(f"### {status_stok}")
    else:
        st.success(f"### {status_stok}")

st.markdown("---")


# =====================================================================
# 9. BAGIAN 2 - VISUALISASI UTAMA (PLOTLY LINE CHART)
# =====================================================================
st.subheader("📈 Perbandingan Data Aktual vs Hasil Peramalan")

fig = go.Figure()

fig.add_trace(go.Scatter(
    x=item_df["Tanggal"], y=actual_series.values,
    mode="lines+markers", name="Data Aktual (Permintaan)",
    line=dict(color=PRIMARY_COLOR, width=2.5),
    marker=dict(size=4),
))

fig.add_trace(go.Scatter(
    x=item_df["Tanggal"], y=ma_forecast_series.values,
    mode="lines", name=f"Forecast MA (n={ma_n})",
    line=dict(color=MA_COLOR, width=2, dash="dash"),
))

fig.add_trace(go.Scatter(
    x=item_df["Tanggal"], y=ses_forecast_series.values,
    mode="lines", name=f"Forecast SES (α={ses_alpha})",
    line=dict(color=SES_COLOR, width=2, dash="dot"),
))

fig.update_layout(
    height=480,
    hovermode="x unified",
    template="plotly_white",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    margin=dict(l=10, r=10, t=40, b=10),
    xaxis_title="Tanggal",
    yaxis_title=f"Pemakaian ({satuan})",
)

st.plotly_chart(fig, use_container_width=True)
st.markdown("---")


# =====================================================================
# 10. BAGIAN 3 - EVALUASI PERFORMA (TABEL & PERBANDINGAN ERROR)
# =====================================================================
st.subheader("🧮 Evaluasi Akurasi Metode Forecasting")

col_ma, col_ses = st.columns(2)

with col_ma:
    st.markdown(f"#### 🟠 Moving Average (n = {ma_n})")
    st.metric("MAD (Mean Absolute Deviation)", f"{ma_eval['MAD']:.3f}")
    st.metric("MSE (Mean Squared Error)", f"{ma_eval['MSE']:.3f}")
    st.metric("MAPE (Mean Absolute Percentage Error)", f"{ma_eval['MAPE']:.2f}%")

with col_ses:
    st.markdown(f"#### 🟢 Single Exponential Smoothing (α = {ses_alpha})")
    st.metric("MAD (Mean Absolute Deviation)", f"{ses_eval['MAD']:.3f}")
    st.metric("MSE (Mean Squared Error)", f"{ses_eval['MSE']:.3f}")
    st.metric("MAPE (Mean Absolute Percentage Error)", f"{ses_eval['MAPE']:.2f}%")

st.markdown("")
st.success(
    f"✅ **Metode Terbaik untuk {selected_item} adalah {best_method_name}** "
    f"dengan tingkat error MAPE sebesar **{best_mape:.2f}%**."
)

st.markdown("---")


# =====================================================================
# 11. BAGIAN 4 - FORECAST t+1 & REKOMENDASI RESTOCK
# =====================================================================
st.subheader("🔮 Hasil Forecasting Periode Berikutnya & Rekomendasi Restock")

next_date = item_df["Tanggal"].max() + pd.Timedelta(days=1)

forecast_table = pd.DataFrame({
    "Tanggal Forecast (t+1)": [next_date.date()],
    "Forecast MA": [round(ma_next, 2)],
    "Forecast SES": [round(ses_next, 2)],
    "Metode Terbaik": [best_method_name],
    f"Forecast Metode Terbaik ({satuan})": [round(best_next_forecast, 2)],
    f"Sisa Stok Saat Ini ({satuan})": [round(sisa_stok, 2)],
    f"Rekomendasi Jumlah Restock ({satuan})": [round(rekomendasi_restock, 2)],
})

st.dataframe(forecast_table, use_container_width=True, hide_index=True)

if rekomendasi_restock > 0:
    st.warning(
        f"⚠️ Disarankan melakukan **restock sebesar {rekomendasi_restock:,.2f} {satuan}** "
        f"untuk **{selected_item}** guna memenuhi proyeksi permintaan tanggal {next_date.date()}."
    )
else:
    st.info(f"ℹ️ Stok **{selected_item}** saat ini masih mencukupi proyeksi permintaan esok hari.")

st.markdown("---")
with st.expander("📋 Lihat Data Historis Lengkap (setelah preprocessing)"):
    st.dataframe(
        item_df[["Tanggal", "ITEM", "Satuan", "Pemakaian", "Batas", "SisaHariIni"]],
        use_container_width=True,
        hide_index=True,
    )

st.caption(
    "Catatan metodologis: Forecast MA(t) = rata-rata aktual (t-n) s/d (t-1). "
    "Forecast SES(t) = α·Aktual(t-1) + (1-α)·Forecast(t-1), dengan F(1) = A(1). "
    "Evaluasi MAD, MSE, MAPE dihitung pada titik-titik yang memiliki pasangan forecast valid."
)
