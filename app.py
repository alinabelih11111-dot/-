import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
from datetime import date
import io
import re

st.set_page_config(page_title="NikiRent | Панель руководителя", page_icon="🏢", layout="wide")

st.markdown("""
<style>
    .stApp { background: #f5f7fb; }
    [data-testid="stMetric"] { background: white; padding: 18px 20px; border-radius: 14px; border: 1px solid #e8edf5; }
    h1, h2, h3 { color: #15243a; }
    div[data-testid="stDataFrame"] { background: white; border-radius: 12px; }
</style>
""", unsafe_allow_html=True)

st.title("🏢 NikiRent — Панель руководителя")
st.caption("Аналитика обращений гостей и контроль нерешённых бытовых вопросов")

REQUIRED = {
    "Дата поступления вопроса": ["дата поступления вопроса", "дата", "дата обращения"],
    "квартира/дом": ["квартира/дом", "квартира", "объект", "адрес"],
    "Коротко суть вопроса": ["коротко суть вопроса", "суть вопроса", "вопрос", "проблема", "описание"],
    "Категория": ["категория"],
    "Ответственный кому передали вопрос": ["ответственный кому передали вопрос", "ответственный", "сотрудник"],
    "Статус": ["статус"],
}
ALIASES = {a: canonical for canonical, vals in REQUIRED.items() for a in vals}

def clean_cell(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def normalize_columns(df):
    # Remove completely empty columns and normalize header spelling.
    df = df.dropna(axis=1, how="all").copy()
    rename = {}
    for col in df.columns:
        key = re.sub(r"\s+", " ", str(col).strip().lower())
        if key in ALIASES:
            rename[col] = ALIASES[key]
    df = df.rename(columns=rename)
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("Не найдены обязательные столбцы: " + ", ".join(missing))
    return df[list(REQUIRED.keys())].copy()

def read_excel(file):
    book = pd.ExcelFile(file)
    chunks = []
    for sheet in book.sheet_names:
        raw = pd.read_excel(file, sheet_name=sheet, header=0)
        # Month divider rows have no date; real records are selected by valid dates below.
        raw = normalize_columns(raw)
        chunks.append(raw)
    return pd.concat(chunks, ignore_index=True) if chunks else pd.DataFrame(columns=list(REQUIRED))

def read_csv(file):
    file.seek(0)
    try:
        raw = pd.read_csv(file, sep=None, engine="python", encoding="utf-8-sig")
    except UnicodeDecodeError:
        file.seek(0)
        raw = pd.read_csv(file, sep=None, engine="python", encoding="cp1251")
    return normalize_columns(raw)

def prepare(df):
    df = df.copy()
    df["Дата поступления вопроса"] = pd.to_datetime(df["Дата поступления вопроса"], errors="coerce", dayfirst=True)
    for c in ["квартира/дом", "Коротко суть вопроса", "Категория", "Ответственный кому передали вопрос", "Статус"]:
        df[c] = df[c].map(clean_cell)
    df = df[df["Дата поступления вопроса"].notna()].copy()
    df = df[df["Категория"].ne("")].copy()
    df["Месяц"] = df["Дата поступления вопроса"].dt.to_period("M").astype(str)
    return df.sort_values("Дата поступления вопроса")

def load_default():
    path = Path(__file__).parent / "data" / "nikirent_source.xlsx"
    if path.exists():
        return prepare(read_excel(path))
    return pd.DataFrame(columns=list(REQUIRED) + ["Месяц"])

st.sidebar.header("📂 Данные")
uploaded = st.sidebar.file_uploader("Загрузить обновлённый Excel или CSV", type=["xlsx", "xls", "csv"])
try:
    if uploaded is not None:
        if uploaded.name.lower().endswith(".csv"):
            source_df = prepare(read_csv(uploaded))
        else:
            source_df = prepare(read_excel(uploaded))
        st.sidebar.success(f"Загружено обращений: {len(source_df)}")
        st.sidebar.caption("Загруженный файл используется для текущего просмотра. Исходный файл проекта не изменяется.")
    else:
        source_df = load_default()
        st.sidebar.info("Используются данные из комплекта проекта.")
except Exception as e:
    st.error(f"Не удалось прочитать файл: {e}")
    st.stop()

if source_df.empty:
    st.warning("В файле пока нет строк с корректной датой и категорией.")
    st.stop()

today = pd.Timestamp(date.today())
months = sorted(source_df["Месяц"].dropna().unique().tolist())
default_month = today.strftime("%Y-%m") if today.strftime("%Y-%m") in months else months[-1]
month = st.selectbox("Месяц статистики", months, index=months.index(default_month), format_func=lambda x: pd.Period(x).strftime("%B %Y"))

monthly = source_df[source_df["Месяц"] == month].copy()
if month == today.strftime("%Y-%m"):
    monthly = monthly[monthly["Дата поступления вопроса"] <= today]

total = len(monthly)
category_counts = monthly["Категория"].value_counts().rename_axis("Категория").reset_index(name="Количество")
category_counts["Доля, %"] = (category_counts["Количество"] / total * 100).round(1) if total else 0

def is_household(value):
    return "быт.вопрос" in str(value).lower().replace(" ", "")

def is_resolved(value):
    return str(value).strip().lower() == "решен"

household = source_df[source_df["Категория"].map(is_household)].copy()
unresolved = household[~household["Статус"].map(is_resolved)].sort_values("Дата поступления вопроса")
unresolved_count = len(unresolved)

st.subheader("Ключевые показатели")
c1, c2, c3 = st.columns(3)
c1.metric("Обращения за выбранный месяц", total)
c2.metric("Категорий обращений", monthly["Категория"].nunique())
c3.metric("Нерешённые бытовые вопросы", unresolved_count)

st.divider()
left, right = st.columns([1, 1])
with left:
    st.subheader("Обращения по категориям")
    if not category_counts.empty:
        st.dataframe(category_counts, use_container_width=True, hide_index=True)
    else:
        st.info("За выбранный месяц обращений нет.")
with right:
    st.subheader("Распределение категорий")
    if not category_counts.empty:
        fig = px.pie(category_counts, names="Категория", values="Количество", hole=0.48)
        fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Нет данных для диаграммы.")

st.subheader("Количество обращений по категориям")
if not category_counts.empty:
    fig_bar = px.bar(category_counts.sort_values("Количество", ascending=True), x="Количество", y="Категория", orientation="h", text="Количество")
    fig_bar.update_layout(xaxis_title="Количество обращений", yaxis_title="", margin=dict(t=10, b=10))
    st.plotly_chart(fig_bar, use_container_width=True)

st.divider()
st.subheader("🔧 Нерешённые бытовые вопросы")
st.caption("Все месяцы. В список попадают все статусы, кроме точного статуса «решен», включая пустые и отложенные.")
if unresolved.empty:
    st.success("Нерешённых бытовых вопросов нет.")
else:
    display = unresolved[[
        "Дата поступления вопроса", "квартира/дом", "Коротко суть вопроса",
        "Ответственный кому передали вопрос", "Статус"
    ]].copy()
    display.columns = ["Дата поступления", "Квартира", "Суть проблемы", "Ответственный", "Статус"]
    display["Дата поступления"] = display["Дата поступления"].dt.strftime("%d.%m.%Y")
    display["Статус"] = display["Статус"].replace("", "Не указан")
    st.dataframe(display, use_container_width=True, hide_index=True)

st.divider()
st.subheader("⬇️ Выгрузка данных")
csv_bytes = source_df.drop(columns=["Месяц"]).to_csv(index=False).encode("utf-8-sig")
st.download_button("Скачать обработанные обращения (CSV)", data=csv_bytes, file_name="nikirent_processed.csv", mime="text/csv")

with st.expander("Методика расчёта"):
    st.write("Обращение учитывается, если заполнены корректная дата поступления и категория. Строки-разделители месяцев и пустые строки не учитываются. Для текущего месяца берутся записи с 1-го числа по сегодняшнюю дату. Нерешённые бытовые вопросы ищутся по всей истории данных; решённым считается только статус, равный «решен» без учёта регистра и пробелов.")
    st.write(f"Всего распознано обращений в загруженном файле: **{len(source_df)}**.")
