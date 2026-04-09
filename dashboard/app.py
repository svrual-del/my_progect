"""
Дашборд эффективности контент-менеджеров.
Данные из Google Sheets (KaspiReporter).
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import json
import re

# ============================================
# НАСТРОЙКИ СТРАНИЦЫ
# ============================================
st.set_page_config(
    page_title="Эффективность контент-менеджеров",
    page_icon="📊",
    layout="wide",
)

# ============================================
# ПОДКЛЮЧЕНИЕ К GOOGLE SHEETS
# ============================================
SPREADSHEET_ID = "16NoTXUjutOw_anh_oSuufYEEfEu6FiHGm9OFnrSkdN8"

# Ожидаемые столбцы в каждом месячном листе
EXPECTED_COLUMNS = [
    "Кабинет", "Артикул", "Название товара", "Дата добавления",
    "Менеджер", "Отметка менеджера", "Дата исчезновения", "Дней до решения"
]


@st.cache_resource
def get_google_client():
    """Подключение к Google Sheets через сервисный аккаунт"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


@st.cache_data(ttl=300)  # кэш 5 минут
def load_all_data():
    """Загрузка данных со всех месячных листов"""
    client = get_google_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    all_frames = []
    month_pattern = re.compile(r"^\d{4}-\d{2}$")

    for ws in spreadsheet.worksheets():
        if not month_pattern.match(ws.title):
            continue

        rows = ws.get_all_values()
        if len(rows) < 2:
            continue

        headers = rows[0]
        data = rows[1:]

        df = pd.DataFrame(data, columns=headers)
        # Унификация: "Ответственный" -> "Менеджер"
        if "Ответственный" in df.columns and "Менеджер" not in df.columns:
            df.rename(columns={"Ответственный": "Менеджер"}, inplace=True)
        df["Месяц"] = ws.title
        all_frames.append(df)

    if not all_frames:
        return pd.DataFrame()

    df = pd.concat(all_frames, ignore_index=True)

    # Парсинг дат (формат dd.mm.yyyy)
    for col in ["Дата добавления", "Отметка менеджера", "Дата исчезновения"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format="%d.%m.%Y", errors="coerce")

    # Парсинг числового столбца
    if "Дней до решения" in df.columns:
        df["Дней до решения"] = pd.to_numeric(
            df["Дней до решения"].replace("ОБМАН", pd.NA), errors="coerce"
        )

    # Флаг обмана (из оригинального значения до замены)
    df["Обман"] = False
    if "Дней до решения" in df.columns:
        # Перечитаем оригинальные значения — проще пометить по условию:
        # Отметка менеджера заполнена, а Дата исчезновения пустая
        df["Обман"] = df["Отметка менеджера"].notna() & df["Дата исчезновения"].isna()

    # Статус задачи
    df["Статус"] = "В работе"
    df.loc[df["Дата исчезновения"].notna(), "Статус"] = "Закрыто"
    df.loc[df["Обман"], "Статус"] = "Обман"

    # Категория скорости (только для закрытых с числом дней)
    df["Скорость"] = pd.NA
    mask_closed = df["Дата исчезновения"].notna() & df["Дней до решения"].notna()
    df.loc[mask_closed & (df["Дней до решения"] <= 3), "Скорость"] = "1-3 дня (норма)"
    df.loc[mask_closed & (df["Дней до решения"] > 3), "Скорость"] = ">3 дней (долго)"

    return df


# ============================================
# ЗАГРУЗКА ДАННЫХ
# ============================================
try:
    df = load_all_data()
except Exception as e:
    st.error(f"Ошибка подключения к Google Sheets: {e}")
    st.stop()

if df.empty:
    st.warning("Нет данных в таблице. Убедитесь, что есть месячные листы (формат: 2026-03).")
    st.stop()

# ============================================
# САЙДБАР — ФИЛЬТРЫ
# ============================================
st.sidebar.title("Фильтры")

# Выбор месяцев
months = sorted(df["Месяц"].unique(), reverse=True)
selected_months = st.sidebar.multiselect("Месяц", months, default=months[:1])

# Выбор кабинетов
merchants = sorted(df["Кабинет"].dropna().unique())
selected_merchants = st.sidebar.multiselect("Кабинет", merchants, default=merchants)

# Выбор менеджеров
managers = sorted(df["Менеджер"].dropna().unique())
selected_managers = st.sidebar.multiselect("Менеджер", managers, default=managers)

# Применяем фильтры
filtered = df[
    df["Месяц"].isin(selected_months)
    & df["Кабинет"].isin(selected_merchants)
    & df["Менеджер"].isin(selected_managers)
]

# Кнопка обновления
if st.sidebar.button("Обновить данные"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.caption(f"Последнее обновление: {datetime.now().strftime('%H:%M:%S')}")

# ============================================
# ЗАГОЛОВОК
# ============================================
st.title("Эффективность контент-менеджеров")

# ============================================
# ОБЩИЕ МЕТРИКИ (КАРТОЧКИ)
# ============================================
total = len(filtered)
closed = len(filtered[filtered["Статус"] == "Закрыто"])
in_progress = len(filtered[filtered["Статус"] == "В работе"])
cheats = len(filtered[filtered["Статус"] == "Обман"])
pct = round(closed / total * 100, 1) if total > 0 else 0
avg_days = filtered.loc[filtered["Дата исчезновения"].notna(), "Дней до решения"].mean()

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Всего задач", total)
col2.metric("Закрыто", closed)
col3.metric("В работе", in_progress)
col4.metric("Обман", cheats)
col5.metric("Среднее дней", round(avg_days, 1) if pd.notna(avg_days) else "—")

st.divider()

# ============================================
# ТАБЛИЦА ПО МЕНЕДЖЕРАМ
# ============================================
st.subheader("Показатели по менеджерам")


def build_manager_stats(data):
    """Сводная статистика по менеджерам"""
    stats = []
    for manager in sorted(data["Менеджер"].dropna().unique()):
        m = data[data["Менеджер"] == manager]
        total_m = len(m)
        closed_m = len(m[m["Статус"] == "Закрыто"])
        in_progress_m = len(m[m["Статус"] == "В работе"])
        cheats_m = len(m[m["Статус"] == "Обман"])
        pct_m = round(closed_m / total_m * 100, 1) if total_m > 0 else 0

        days_closed = m.loc[m["Дата исчезновения"].notna(), "Дней до решения"]
        avg_m = round(days_closed.mean(), 1) if len(days_closed) > 0 and days_closed.notna().any() else None

        fast_m = len(m[m["Скорость"] == "1-3 дня (норма)"])
        slow_m = len(m[m["Скорость"] == ">3 дней (долго)"])

        stats.append({
            "Менеджер": manager,
            "Назначено": total_m,
            "Закрыто": closed_m,
            "В работе": in_progress_m,
            "Обман": cheats_m,
            "% выполнения": pct_m,
            "Ср. дней": avg_m if avg_m is not None else "—",
            "Норма (1-3)": fast_m,
            "Долго (>3)": slow_m,
        })
    return pd.DataFrame(stats)


stats_df = build_manager_stats(filtered)

if not stats_df.empty:
    st.dataframe(
        stats_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "% выполнения": st.column_config.ProgressColumn(
                "% выполнения", min_value=0, max_value=100, format="%.1f%%"
            ),
        },
    )

st.divider()

# ============================================
# ГРАФИКИ
# ============================================
chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    st.subheader("Назначено vs Закрыто")
    if not stats_df.empty:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="Назначено", x=stats_df["Менеджер"], y=stats_df["Назначено"],
            marker_color="#636EFA"
        ))
        fig.add_trace(go.Bar(
            name="Закрыто", x=stats_df["Менеджер"], y=stats_df["Закрыто"],
            marker_color="#00CC96"
        ))
        fig.update_layout(barmode="group", height=400, margin=dict(t=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

with chart_col2:
    st.subheader("Среднее дней до решения")
    if not stats_df.empty:
        avg_data = stats_df[stats_df["Ср. дней"] != "—"].copy()
        if not avg_data.empty:
            avg_data["Ср. дней"] = pd.to_numeric(avg_data["Ср. дней"])
            colors = ["#00CC96" if v <= 3 else "#EF553B" for v in avg_data["Ср. дней"]]
            fig2 = go.Figure(go.Bar(
                x=avg_data["Менеджер"], y=avg_data["Ср. дней"],
                marker_color=colors,
            ))
            fig2.add_hline(y=3, line_dash="dash", line_color="gray",
                           annotation_text="Норма (3 дня)")
            fig2.update_layout(height=400, margin=dict(t=20, b=20))
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Нет данных о закрытых задачах")

st.divider()

# ============================================
# РАСПРЕДЕЛЕНИЕ СКОРОСТИ
# ============================================
chart_col3, chart_col4 = st.columns(2)

with chart_col3:
    st.subheader("Скорость обработки")
    speed_data = filtered[filtered["Скорость"].notna()]
    if not speed_data.empty:
        fig3 = px.histogram(
            speed_data, x="Менеджер", color="Скорость",
            color_discrete_map={
                "1-3 дня (норма)": "#00CC96",
                ">3 дней (долго)": "#EF553B",
            },
            barmode="group",
            height=400,
        )
        fig3.update_layout(margin=dict(t=20, b=20))
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Нет данных о скорости")

with chart_col4:
    st.subheader("Статусы задач")
    status_counts = filtered["Статус"].value_counts()
    if not status_counts.empty:
        fig4 = px.pie(
            values=status_counts.values,
            names=status_counts.index,
            color=status_counts.index,
            color_discrete_map={
                "Закрыто": "#00CC96",
                "В работе": "#636EFA",
                "Обман": "#EF553B",
            },
            height=400,
        )
        fig4.update_layout(margin=dict(t=20, b=20))
        st.plotly_chart(fig4, use_container_width=True)

st.divider()

# ============================================
# ДИНАМИКА ПО ДНЯМ
# ============================================
st.subheader("Динамика добавления задач по дням")

daily = filtered[filtered["Дата добавления"].notna()].copy()
if not daily.empty:
    daily_counts = (
        daily.groupby([daily["Дата добавления"].dt.date, "Менеджер"])
        .size()
        .reset_index(name="Количество")
    )
    daily_counts.rename(columns={"Дата добавления": "Дата"}, inplace=True)

    fig5 = px.line(
        daily_counts, x="Дата", y="Количество", color="Менеджер",
        height=400, markers=True,
    )
    fig5.update_layout(margin=dict(t=20, b=20))
    st.plotly_chart(fig5, use_container_width=True)
else:
    st.info("Нет данных для графика динамики")

# ============================================
# ДЕТАЛЬНАЯ ТАБЛИЦА (РАЗВОРАЧИВАЕМАЯ)
# ============================================
with st.expander("Детальные данные"):
    display_cols = [
        "Месяц", "Кабинет", "Артикул", "Название товара",
        "Дата добавления", "Менеджер", "Отметка менеджера",
        "Дата исчезновения", "Дней до решения", "Статус"
    ]
    existing_cols = [c for c in display_cols if c in filtered.columns]
    st.dataframe(
        filtered[existing_cols].sort_values("Дата добавления", ascending=False),
        use_container_width=True,
        hide_index=True,
        height=400,
    )
