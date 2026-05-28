import streamlit as st
import pandas as pd
import altair as alt
from pathlib import Path
from datetime import date, timedelta
import calendar
import math
import json
import time
import uuid
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

st.set_page_config(
    page_title="WorkTime Sync AI Assistant",
    page_icon="🕒",
    layout="wide"
)

DATA_FILE = "Сочи2016.xlsx"

ROLE_OPTIONS = [
    "Сотрудник",
    "Администратор",
    "HR",
    "Проектный менеджер",
    "Аналитик",
    "Руководитель"
]

RAINBOW = ["#e40303", "#ff8c00", "#ffed00", "#008026", "#004dff", "#750787"]

def load_sheet(file_path, sheet_name):
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

@st.cache_data
def load_data(file_path):
    return {
        "employees": load_sheet(file_path, "employees"),
        "schedules": load_sheet(file_path, "schedules"),
        "absences": load_sheet(file_path, "absences"),
        "calendar_events": load_sheet(file_path, "calendar_events"),
        "workload_logs": load_sheet(file_path, "workload_logs"),
        "analytics_metrics": load_sheet(file_path, "analytics_metrics"),
        "recommendations": load_sheet(file_path, "recommendations"),
        "bi_dashboard": load_sheet(file_path, "bi_dashboard"),
        "bi_heatmap": load_sheet(file_path, "bi_heatmap"),
        "bi_workload": load_sheet(file_path, "bi_workload")
    }

def normalize_columns(df):
    df = df.copy()
    df.columns = [str(col).strip().lower() for col in df.columns]
    return df

def normalize_id_columns(df):
    df = normalize_columns(df)

    if "employee_id" not in df.columns and "id" in df.columns:
        if df["id"].astype(str).str.upper().str.startswith("EMP").any():
            df = df.rename(columns={"id": "employee_id"})

    return df

def safe_numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)

def safe_mean(df, column):
    if df.empty or column not in df.columns:
        return 0
    return round(safe_numeric(df[column]).mean(), 2)

def safe_sum(df, column):
    if df.empty or column not in df.columns:
        return 0
    return round(safe_numeric(df[column]).sum(), 2)

def parse_excel_datetime_series(series):
    raw = pd.Series(series)
    numeric = pd.to_numeric(raw, errors="coerce")
    parsed = pd.to_datetime(raw, errors="coerce")
    excel_mask = numeric.between(20000, 60000)
    if excel_mask.any():
        parsed.loc[excel_mask] = pd.to_datetime(numeric.loc[excel_mask], unit="D", origin="1899-12-30", errors="coerce")
    return parsed

def escape_html(value):
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def get_employee_name(row, employees):
    employees = normalize_id_columns(employees)

    if "employee_id" not in row.index or employees.empty:
        return "Сотрудник"

    employee_id = str(row["employee_id"])

    if "employee_id" in employees.columns and "full_name" in employees.columns:
        match = employees[employees["employee_id"].astype(str) == employee_id]
        if not match.empty:
            return str(match.iloc[0]["full_name"])

    return employee_id

def add_employee_names(df, employees):
    df = normalize_id_columns(df)
    employees = normalize_id_columns(employees)

    if df.empty:
        return df

    if "employee_id" in df.columns and "employee_id" in employees.columns:
        cols = ["employee_id"]
        for col in ["full_name", "department", "role"]:
            if col in employees.columns:
                cols.append(col)

        result = df.merge(employees[cols], on="employee_id", how="left")

        if "full_name" not in result.columns:
            result["full_name"] = result["employee_id"].astype(str)
        else:
            result["full_name"] = result["full_name"].fillna(result["employee_id"].astype(str))

        return result

    return df

def classify_risk(value):
    try:
        value = float(value)
    except Exception:
        return "не определен"

    if value >= 0.75:
        return "критический"
    if value >= 0.55:
        return "высокий"
    if value >= 0.35:
        return "средний"
    return "низкий"

def score_color(value):
    try:
        value = float(value)
    except Exception:
        value = 0

    if value >= 80:
        return "#1f8f4d"
    if value >= 60:
        return "#8abf26"
    if value >= 40:
        return "#f2b705"
    if value >= 20:
        return "#f27c38"
    return "#d93636"

def risk_color(value):
    try:
        value = float(value)
    except Exception:
        value = 0

    if value >= 0.75:
        return "#d93636"
    if value >= 0.55:
        return "#f27c38"
    if value >= 0.35:
        return "#f2b705"
    return "#1f8f4d"

def normalize_percent_value(value):
    if pd.isna(value):
        return 0

    text = str(value).strip().replace("%", "").replace(",", ".")

    try:
        number = float(text)
    except Exception:
        return 0

    if 0 <= number <= 1:
        number *= 100

    return max(0, min(100, number))

def render_color_table(df, columns=None, title=None, score_column=None, risk_column=None):
    if df.empty:
        st.info("Нет данных для отображения.")
        return

    view = df.copy()

    if columns:
        view = view[[col for col in columns if col in view.columns]]

    if title:
        st.subheader(title)

    html = """
    <style>
    .wts-table {
        width: 100%;
        border-collapse: collapse;
        font-family: Arial, sans-serif;
        font-size: 14px;
        border-radius: 12px;
        overflow: hidden;
        background: white;
    }
    .wts-table th {
        background: #111827;
        color: white;
        padding: 10px;
        text-align: left;
        border: 1px solid #374151;
    }
    .wts-table td {
        padding: 9px;
        border: 1px solid #e5e7eb;
        color: #111827;
        background: white;
    }
    .wts-table tr:nth-child(even) td {
        background: #f9fafb;
    }
    .wts-pill {
        display: inline-block;
        min-width: 58px;
        padding: 4px 8px;
        border-radius: 999px;
        color: white;
        text-align: center;
        font-weight: 700;
    }
    </style>
    <table class="wts-table">
    <thead><tr>
    """

    for col in view.columns:
        html += f"<th>{escape_html(col)}</th>"

    html += "</tr></thead><tbody>"

    for _, row in view.iterrows():
        html += "<tr>"
        for col in view.columns:
            value = row[col]

            if score_column and col == score_column:
                try:
                    numeric = round(float(value), 1)
                except Exception:
                    numeric = value
                color = score_color(numeric)
                html += f'<td><span class="wts-pill" style="background:{color}">{escape_html(numeric)}</span></td>'

            elif risk_column and col == risk_column:
                try:
                    numeric = round(float(value), 2)
                except Exception:
                    numeric = value
                color = risk_color(numeric)
                html += f'<td><span class="wts-pill" style="background:{color}">{escape_html(numeric)}</span></td>'

            else:
                html += f"<td>{escape_html(value)}</td>"

        html += "</tr>"

    html += "</tbody></table>"
    st.markdown(html, unsafe_allow_html=True)

def render_rainbow_bar_chart(df, label_col, value_col, title, max_value=None):
    if df.empty or label_col not in df.columns or value_col not in df.columns:
        st.info("Нет данных для графика.")
        return

    st.subheader(title)

    chart = df[[label_col, value_col]].copy()
    chart[value_col] = safe_numeric(chart[value_col])
    chart = chart.dropna()

    if chart.empty:
        st.info("Нет данных для графика.")
        return

    chart[label_col] = chart[label_col].astype(str)
    chart = chart.sort_values(value_col, ascending=False)

    st.bar_chart(chart.set_index(label_col)[value_col])

def clean_heatmap(raw):
    raw = raw.copy()

    if raw.empty:
        return pd.DataFrame()

    df = raw.copy()

    first_col = str(df.columns[0]).strip().lower()

    if first_col in ["time_slot", "time", "время"]:
        clean = df.copy()
    else:
        header_idx = None

        for idx, row in df.iterrows():
            values = [str(x).strip().lower() for x in row.tolist()]
            if "время" in values or "time_slot" in values or "time" in values:
                header_idx = idx
                break

        if header_idx is None:
            return pd.DataFrame()

        headers = df.loc[header_idx].tolist()

        clean_cols = []
        for i, col in enumerate(headers):
            name = str(col).strip()
            if name == "nan" or name == "":
                name = f"extra_{i}"
            clean_cols.append(name)

        clean = df.loc[header_idx + 1:].copy()
        clean.columns = clean_cols

    cols = list(clean.columns)

    time_col = None
    for col in cols:
        if str(col).strip().lower() in ["время", "time", "time_slot"]:
            time_col = col
            break

    if time_col is None:
        time_col = cols[0]

    keep_cols = [time_col]

    for col in clean.columns:
        col_text = str(col).strip().lower()
        if col == time_col:
            continue
        if col_text in ["пн", "вт", "ср", "чт", "пт", "mon", "tue", "wed", "thu", "fri", "monday", "tuesday", "wednesday", "thursday", "friday"]:
            keep_cols.append(col)

    if len(keep_cols) == 1:
        for col in clean.columns[1:6]:
            keep_cols.append(col)

    clean = clean[keep_cols].copy()
    clean = clean.dropna(how="all")

    clean = clean[clean[time_col].notna()]
    clean = clean[~clean[time_col].astype(str).str.lower().isin(["nan", "время"])]

    clean = clean.rename(columns={time_col: "Время"})

    rename_days = {
        "mon": "Пн",
        "monday": "Пн",
        "tue": "Вт",
        "tuesday": "Вт",
        "wed": "Ср",
        "wednesday": "Ср",
        "thu": "Чт",
        "thursday": "Чт",
        "fri": "Пт",
        "friday": "Пт"
    }

    clean = clean.rename(columns={col: rename_days.get(str(col).strip().lower(), col) for col in clean.columns})

    return clean

def render_heatmap_html(raw_heatmap):
    heatmap = clean_heatmap(raw_heatmap)

    if heatmap.empty:
        st.info("Лист bi_heatmap не найден или пустой.")
        return

    day_cols = [col for col in heatmap.columns if col != "Время"]

    html = """
    <style>
    .heatmap-table {
        width: 100%;
        border-collapse: collapse;
        font-family: Arial, sans-serif;
        font-size: 14px;
        border-radius: 12px;
        overflow: hidden;
    }
    .heatmap-table th {
        background: #111827;
        color: white;
        padding: 10px;
        text-align: center;
        border: 1px solid #374151;
    }
    .heatmap-table td {
        padding: 10px;
        text-align: center;
        border: 1px solid rgba(255,255,255,0.45);
        color: #111827;
        font-weight: 700;
    }
    .heatmap-time {
        background: white !important;
        font-weight: 800;
    }
    </style>
    <table class="heatmap-table">
    <thead><tr><th>Время</th>
    """

    for col in day_cols:
        html += f"<th>{escape_html(col)}</th>"

    html += "</tr></thead><tbody>"

    for _, row in heatmap.iterrows():
        html += f'<tr><td class="heatmap-time">{escape_html(row["Время"])}</td>'

        for col in day_cols:
            value = normalize_percent_value(row[col])
            hue = int(value * 1.35)
            color = f"hsl({hue}, 92%, 78%)"
            html += f'<td style="background:{color}">{round(value, 0)}%</td>'

        html += "</tr>"

    html += "</tbody></table>"
    st.markdown(html, unsafe_allow_html=True)

    numeric = heatmap.copy()
    for col in day_cols:
        numeric[col] = numeric[col].apply(normalize_percent_value)

    day_availability = numeric[day_cols].mean(axis=0).reset_index()
    day_availability.columns = ["День", "Средняя доступность"]

    render_rainbow_bar_chart(day_availability, "День", "Средняя доступность", "Средняя доступность по дням", 100)

def compute_general_score(data):
    employees = normalize_id_columns(data["employees"])
    metrics = normalize_id_columns(data["analytics_metrics"])
    workload_logs = normalize_id_columns(data["workload_logs"])
    calendar_events = normalize_id_columns(data["calendar_events"])

    if employees.empty or "employee_id" not in employees.columns:
        return pd.DataFrame()

    result = employees[["employee_id"]].copy()

    result["full_name"] = employees["full_name"] if "full_name" in employees.columns else employees["employee_id"]

    if "department" in employees.columns:
        result["department"] = employees["department"]
    else:
        result["department"] = "Общий отдел"

    if "role" in employees.columns:
        result["role"] = employees["role"]

    result["tasks_done"] = 0
    result["project_hours"] = 0.0
    result["activity_points"] = 0
    result["overtime_hours"] = 0.0

    if not workload_logs.empty and "employee_id" in workload_logs.columns:
        logs = workload_logs.copy()

        if "spent_hours" in logs.columns:
            logs["spent_hours"] = safe_numeric(logs["spent_hours"])
        else:
            logs["spent_hours"] = 0

        tasks_done = logs.groupby("employee_id").size()
        project_hours = logs.groupby("employee_id")["spent_hours"].sum()

        result["tasks_done"] = result["employee_id"].map(tasks_done).fillna(0).astype(int)
        result["project_hours"] = result["employee_id"].map(project_hours).fillna(0).round(1)

        text_cols = [col for col in ["task_name", "task_template", "task_priority", "source"] if col in logs.columns]

        if text_cols:
            text = logs[text_cols].astype(str).agg(" ".join, axis=1).str.lower()
            activity_mask = text.str.contains("доп|замен|extra|replace|support|помощ|сроч|urgent|critical|high", regex=True)
            activity_points = logs[activity_mask].groupby("employee_id").size()
            result["activity_points"] = result["employee_id"].map(activity_points).fillna(0).astype(int)

    if not metrics.empty and "employee_id" in metrics.columns:
        metric_cols = [col for col in ["employee_id", "risk_score", "load_ratio", "workload_score", "conflict_count", "days_since_update", "outside_hours_ratio", "busy_hours"] if col in metrics.columns]
        result = result.merge(metrics[metric_cols], on="employee_id", how="left")

        if "conflict_count" in result.columns:
            result["overtime_hours"] += safe_numeric(result["conflict_count"]) * 0.5

        if "outside_hours_ratio" in result.columns and "busy_hours" in result.columns:
            result["overtime_hours"] += safe_numeric(result["outside_hours_ratio"]) * safe_numeric(result["busy_hours"])

    if not calendar_events.empty and "employee_id" in calendar_events.columns:
        events = calendar_events.copy()
        text_cols = [col for col in ["event_type", "comment", "source"] if col in events.columns]

        if text_cols:
            text = events[text_cols].astype(str).agg(" ".join, axis=1).str.lower()
            activity_mask = text.str.contains("доп|замен|extra|replace|support|помощ|сроч|urgent", regex=True)
            event_activity = events[activity_mask].groupby("employee_id").size()
            result["activity_points"] = result["activity_points"] + result["employee_id"].map(event_activity).fillna(0).astype(int)

    def normalize(series):
        series = safe_numeric(series)
        max_value = series.max()
        if max_value <= 0:
            return series * 0
        return series / max_value

    tasks_part = normalize(result["tasks_done"]) * 35
    hours_part = normalize(result["project_hours"]) * 25
    activity_part = normalize(result["activity_points"]) * 20
    overtime_penalty = normalize(result["overtime_hours"]) * 20

    result["general_score"] = (tasks_part + hours_part + activity_part - overtime_penalty).round(1)
    result["general_score"] = result["general_score"].clip(lower=0, upper=100)
    result["rank"] = result["general_score"].rank(method="dense", ascending=False).astype(int)

    result = result.sort_values(["rank", "general_score"], ascending=[True, False])

    return result

def build_team_summary(metrics, employees):
    metrics = normalize_id_columns(metrics)

    if metrics.empty:
        return "Недостаточно данных для анализа команды."

    total = len(metrics)
    risk_col = "risk_score" if "risk_score" in metrics.columns else None
    load_col = "workload_score" if "workload_score" in metrics.columns else "load_ratio" if "load_ratio" in metrics.columns else None

    avg_risk = safe_mean(metrics, risk_col) if risk_col else 0
    avg_load = safe_mean(metrics, load_col) if load_col else 0

    high_risk = int((safe_numeric(metrics[risk_col]) >= 0.55).sum()) if risk_col else 0
    overloaded = int((safe_numeric(metrics[load_col]) >= 0.8).sum()) if load_col else 0
    outdated = int((safe_numeric(metrics["days_since_update"]) > 60).sum()) if "days_since_update" in metrics.columns else 0
    conflicts_total = int(safe_sum(metrics, "conflict_count")) if "conflict_count" in metrics.columns else 0

    text = (
        f"В команде проанализировано сотрудников: {total}.\n\n"
        f"Средний risk score: {avg_risk}.\n\n"
        f"Средняя загрузка: {avg_load}.\n\n"
        f"Сотрудников с высокой загрузкой: {overloaded}.\n\n"
        f"Сотрудников с высоким риском: {high_risk}.\n\n"
        f"Сотрудников с устаревшим графиком: {outdated}.\n\n"
        f"Общее количество конфликтов: {conflicts_total}.\n\n"
    )

    solutions = []

    if outdated > 0:
        solutions.append("HR: отправить запросы на подтверждение графика сотрудникам, у которых данные не обновлялись более 60 дней.")
    if overloaded > 0:
        solutions.append("Руководитель: перераспределить часть задач с перегруженных сотрудников на менее загруженных.")
    if conflicts_total > 0:
        solutions.append("Проектный менеджер: пересмотреть регулярные встречи и перенести события, которые выходят за рабочее время.")
    if high_risk > 0:
        solutions.append("Аналитик: отдельно проверить сотрудников с высоким risk score и определить главную причину риска.")

    if solutions:
        text += "Возможное решение:\n\n" + "\n\n".join([f"• {item}" for item in solutions])
    else:
        text += "Серьезных проблем по команде не выявлено.\n\nВозможное решение:\n\n• Оставить текущий график, но настроить регулярное подтверждение актуальности данных раз в 30 дней."

    return text

def build_employee_recommendation(row, employees):
    name = get_employee_name(row, employees)

    risk = row.get("risk_score", 0)
    workload = row.get("workload_score", row.get("load_ratio", 0))
    outside_ratio = row.get("outside_worktime_ratio", row.get("outside_hours_ratio", 0))
    conflicts = row.get("conflict_count", 0)
    days = row.get("days_since_update", 0)

    analytics = []
    actions = []

    analytics.append(f"По сотруднику {name} рассчитан уровень риска: {classify_risk(risk)}.")

    try:
        if float(days) > 60:
            analytics.append(f"График не обновлялся {int(days)} дней.")
            actions.append("Отправить сотруднику запрос на подтверждение рабочего графика и поставить дедлайн обновления на ближайшие 1–2 рабочих дня.")
    except Exception:
        pass

    try:
        if float(workload) >= 0.8:
            analytics.append("У сотрудника высокая загрузка.")
            actions.append("Не назначать новые задачи без согласования с руководителем и перераспределить часть задач.")
    except Exception:
        pass

    try:
        if float(outside_ratio) >= 0.25:
            analytics.append("Значительная часть событий проходит вне рабочего времени.")
            actions.append("Проверить регулярные встречи после рабочего дня и перенести повторяющиеся созвоны в доступное командное окно.")
    except Exception:
        pass

    try:
        if float(conflicts) >= 3:
            analytics.append(f"Обнаружено {int(conflicts)} конфликтов между графиком, календарем или исключениями.")
            actions.append("Сверить календарь сотрудника с отсутствиями, отпуском и рабочими часами.")
    except Exception:
        pass

    if not actions:
        actions.append("Критичных проблем не обнаружено. Рекомендуется оставить текущий график и повторно подтвердить его в конце отчетного периода.")

    return (
        "Аналитика:\n"
        + "\n".join([f"• {item}" for item in analytics])
        + "\n\nВозможное решение:\n"
        + "\n".join([f"• {item}" for item in actions])
    )

def get_top_problem_employees(metrics, employees, limit=10):
    metrics = add_employee_names(metrics, employees)

    if metrics.empty or "risk_score" not in metrics.columns:
        return pd.DataFrame()

    metrics["risk_score"] = safe_numeric(metrics["risk_score"])
    top = metrics.sort_values("risk_score", ascending=False).head(limit).copy()

    visible = ["full_name"]

    for col in ["department", "risk_score", "load_ratio", "workload_score", "outside_hours_ratio", "conflict_count", "days_since_update", "risk_status", "behavior_segment"]:
        if col in top.columns:
            visible.append(col)

    return top[visible]

def get_role_analytics(role, data):
    score_df = compute_general_score(data)

    if score_df.empty:
        return "Недостаточно данных для ролевой аналитики."

    top = score_df.head(3)
    leader_names = ", ".join(top["full_name"].astype(str).tolist())

    if role == "Сотрудник":
        return "Сотрудник видит личный скор, выполненные задачи, часы на проекте, переработки и рекомендации по улучшению результата."

    if role == "HR":
        outdated_count = int((safe_numeric(score_df["days_since_update"]) > 60).sum()) if "days_since_update" in score_df.columns else 0
        return f"HR видит актуальность профилей, отсутствия и часовые пояса. Сейчас график требует проверки у {outdated_count} сотрудников."

    if role == "Проектный менеджер":
        return f"Проектный менеджер видит задачи, проектные часы и переработки. Текущие лидеры по общему скору: {leader_names}."

    if role == "Аналитик":
        avg_score = round(safe_numeric(score_df["general_score"]).mean(), 1)
        return f"Аналитик видит общий скоринг, департаменты и аномалии. Средний общий скор команды: {avg_score}."

    if role == "Руководитель":
        return f"Руководитель видит итоговый лидерборд и управленческие сигналы. Текущие лидеры: {leader_names}."

    if role == "Администратор":
        return "Администратор видит качество данных, роли, источники и корректность связей между таблицами."

    return "Ролевая аналитика сформирована."

# =====================================================================
# GigaChat integration
# Подключение реальной LLM с обезличиванием персональных данных и
# fallback на детерминированную логику при недоступности API.
# =====================================================================

GIGACHAT_OAUTH_URL = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
GIGACHAT_CHAT_URL = "https://gigachat.devices.sberbank.ru/api/v1/chat/completions"

GIGACHAT_SYSTEM_PROMPT = (
    "Ты эксперт-ассистент приложения WorkTime Sync — системы актуализации рабочих графиков "
    "и анализа загрузки команды.\n\n"
    "КОНТЕКСТ ПРИЛОЖЕНИЯ\n"
    "Данные собираются из таблиц employees, schedules, absences, calendar_events, "
    "workload_logs, analytics_metrics, recommendations и BI-листов (bi_dashboard, "
    "bi_heatmap, bi_workload). Сначала всё агрегируется в общем контуре, затем каждой "
    "роли показывается свой срез аналитики.\n\n"
    "РОЛИ И ИХ ЭКРАНЫ\n"
    "• Сотрудник — личный кабинет: свой график, трудозатраты, личный скор, персональные рекомендации.\n"
    "• HR — кадровый контур: список устаревших графиков, форматы работы, отсутствия, часовые пояса.\n"
    "• Проектный менеджер — проектный контур: лидерборд, детализация трудозатрат, задачи, загрузка команды.\n"
    "• Аналитик — BI-панель: дашборды, тепловая карта доступности, разбивка по департаментам, аномалии.\n"
    "• Руководитель — управленческая панель: BI, лидерборд, проблемные сотрудники, итоговые рекомендации.\n"
    "• Администратор — техническая панель: качество данных, связи таблиц, пользователи, источники.\n\n"
    "МЕТРИКИ И ПОРОГИ\n"
    "• risk_score (0..1): <0.35 низкий, 0.35–0.55 средний, 0.55–0.75 высокий, ≥0.75 критический.\n"
    "• workload_score / load_ratio (0..1): ≥0.8 — перегрузка, требуется перераспределение задач.\n"
    "• days_since_update (дни): >60 — график устарел, нужен запрос на подтверждение от сотрудника.\n"
    "• outside_hours_ratio (0..1): ≥0.25 — значимая часть событий вне рабочего времени, проверить регулярные встречи.\n"
    "• conflict_count (шт): ≥3 — много конфликтов между графиком, календарём и отсутствиями.\n"
    "• general_score (0..100): итоговый скор сотрудника. "
    "Формула: 35% выполненных задач + 25% проектных часов + 20% доп. активностей − 20% переработок.\n"
    "• overtime_hours (часы): переработки, считаются из конфликтов и событий вне рабочего времени.\n\n"
    "ОБЕЗЛИЧИВАНИЕ\n"
    "Сотрудники в блоке DATA обозначены как \"Сотрудник 1\", \"Сотрудник 2\" и так далее. "
    "Это анонимизация, реальные ФИО подставит приложение после твоего ответа.\n\n"
    "ПРАВИЛА ОТВЕТА\n"
    "1. Используй ТОЛЬКО факты из блока DATA. Не выдумывай цифры, имена и причины.\n"
    "2. Сохраняй обозначения сотрудников ровно как в DATA. Никогда не придумывай ФИО.\n"
    "3. Когда называешь метрику, опирайся на пороги выше "
    "(например: «risk_score 0.62 — высокий уровень риска»).\n"
    "4. Можешь ссылаться на конкретные экраны системы из списка ролей выше "
    "(например: «откройте лидерборд», «проверьте тепловую карту доступности»).\n"
    "5. Адаптируй фокус ответа под роль пользователя.\n"
    "6. Отвечай на русском языке.\n"
    "7. Структурируй ответ строго по шаблону:\n\n"
    "Аналитика:\n"
    "• наблюдение 1\n"
    "• наблюдение 2\n\n"
    "Возможное решение:\n"
    "• конкретное действие 1\n"
    "• конкретное действие 2\n\n"
    "8. Если данных недостаточно, честно скажи об этом, не выдумывай.\n"
    "9. Будь конкретен: предлагай действия, а не общие советы. 4–8 пунктов суммарно."
)


def _gigachat_secret(key, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def _gigachat_configured():
    return bool(_gigachat_secret("GIGACHAT_CREDENTIALS"))


def _gigachat_get_token():
    now = time.time()
    cached = st.session_state.get("_gigachat_token")
    if cached and cached.get("expires_at", 0) > now + 60:
        return cached["access_token"]

    credentials = _gigachat_secret("GIGACHAT_CREDENTIALS")
    scope = _gigachat_secret("GIGACHAT_SCOPE", "GIGACHAT_API_PERS")
    verify_ssl = bool(_gigachat_secret("GIGACHAT_VERIFY_SSL", False))

    response = requests.post(
        GIGACHAT_OAUTH_URL,
        headers={
            "Authorization": f"Basic {credentials}",
            "RqUID": str(uuid.uuid4()),
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        },
        data={"scope": scope},
        verify=verify_ssl,
        timeout=30,
    )
    response.raise_for_status()
    payload = response.json()
    raw_expires = payload.get("expires_at")
    if raw_expires and raw_expires > 10_000_000_000:
        expires_at = raw_expires / 1000
    elif raw_expires:
        expires_at = float(raw_expires)
    else:
        expires_at = now + 1700

    st.session_state["_gigachat_token"] = {
        "access_token": payload["access_token"],
        "expires_at": expires_at,
    }
    return payload["access_token"]


def _gigachat_call(system_prompt, user_message):
    token = _gigachat_get_token()
    model = _gigachat_secret("GIGACHAT_MODEL", "GigaChat")
    verify_ssl = bool(_gigachat_secret("GIGACHAT_VERIFY_SSL", False))

    response = requests.post(
        GIGACHAT_CHAT_URL,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message},
            ],
            "temperature": 0.3,
            "top_p": 0.9,
            "max_tokens": 1024,
        },
        verify=verify_ssl,
        timeout=60,
    )
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"].strip()


def _register_pseudonym(real_name, pseudo_map, reverse_map):
    real_name = str(real_name).strip()
    if not real_name or real_name.lower() == "nan":
        return None
    if real_name in reverse_map:
        return reverse_map[real_name]
    pseudo = f"Сотрудник {len(pseudo_map) + 1}"
    pseudo_map[pseudo] = real_name
    reverse_map[real_name] = pseudo
    return pseudo


def _safe_value(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return round(float(value), 3)
    return str(value)


def collect_facts_for_role(role, data, selected_employee_id=None):
    """Собирает обезличенные факты и карту псевдоним -> реальное имя."""
    employees = normalize_id_columns(data["employees"])
    metrics_named = add_employee_names(data["analytics_metrics"], employees)
    metrics = normalize_id_columns(data["analytics_metrics"])
    score_df = compute_general_score(data)

    pseudo_map = {}
    reverse_map = {}

    facts = {"role": role, "team": {}, "problem_employees": [], "leaders": []}

    if not metrics.empty:
        risk_col = "risk_score" if "risk_score" in metrics.columns else None
        load_col = (
            "workload_score" if "workload_score" in metrics.columns
            else "load_ratio" if "load_ratio" in metrics.columns
            else None
        )
        facts["team"] = {
            "total": len(metrics),
            "avg_risk": safe_mean(metrics, risk_col) if risk_col else None,
            "avg_load": safe_mean(metrics, load_col) if load_col else None,
            "overloaded": int((safe_numeric(metrics[load_col]) >= 0.8).sum()) if load_col else 0,
            "high_risk": int((safe_numeric(metrics[risk_col]) >= 0.55).sum()) if risk_col else 0,
            "outdated_schedules": (
                int((safe_numeric(metrics["days_since_update"]) > 60).sum())
                if "days_since_update" in metrics.columns else 0
            ),
            "conflicts_total": (
                int(safe_sum(metrics, "conflict_count"))
                if "conflict_count" in metrics.columns else 0
            ),
        }

    top_problems = get_top_problem_employees(metrics_named, employees, 5)
    if not top_problems.empty:
        for _, row in top_problems.iterrows():
            pseudo = _register_pseudonym(row.get("full_name", row.get("employee_id", "")), pseudo_map, reverse_map)
            if pseudo is None:
                continue
            facts["problem_employees"].append({
                "id": pseudo,
                "department": _safe_value(row.get("department")),
                "risk_score": _safe_value(row.get("risk_score")),
                "workload": _safe_value(row.get("workload_score", row.get("load_ratio"))),
                "conflict_count": _safe_value(row.get("conflict_count")),
                "days_since_update": _safe_value(row.get("days_since_update")),
                "outside_hours_ratio": _safe_value(row.get("outside_hours_ratio")),
            })

    if not score_df.empty:
        for _, row in score_df.head(5).iterrows():
            pseudo = _register_pseudonym(row.get("full_name", row.get("employee_id", "")), pseudo_map, reverse_map)
            if pseudo is None:
                continue
            facts["leaders"].append({
                "id": pseudo,
                "department": _safe_value(row.get("department")),
                "general_score": _safe_value(row.get("general_score")),
                "tasks_done": _safe_value(row.get("tasks_done")),
                "project_hours": _safe_value(row.get("project_hours")),
                "overtime_hours": _safe_value(row.get("overtime_hours")),
            })

    if role == "Сотрудник" and selected_employee_id is not None:
        emp_metrics = filter_employee_rows(data["analytics_metrics"], selected_employee_id)
        if not emp_metrics.empty:
            row = emp_metrics.iloc[0]
            emp_row = employees[employees["employee_id"].astype(str) == str(selected_employee_id)]
            real_name = (
                str(emp_row.iloc[0]["full_name"])
                if not emp_row.empty and "full_name" in emp_row.columns
                else str(selected_employee_id)
            )
            pseudo = _register_pseudonym(real_name, pseudo_map, reverse_map)
            facts["personal"] = {
                "id": pseudo,
                "department": _safe_value(emp_row.iloc[0].get("department")) if not emp_row.empty else None,
                "risk_score": _safe_value(row.get("risk_score")),
                "workload": _safe_value(row.get("workload_score", row.get("load_ratio"))),
                "days_since_update": _safe_value(row.get("days_since_update")),
                "conflict_count": _safe_value(row.get("conflict_count")),
                "outside_hours_ratio": _safe_value(row.get("outside_hours_ratio")),
            }

    return facts, pseudo_map


def deanonymize(text, pseudo_map):
    if not text or not pseudo_map:
        return text
    for pseudo in sorted(pseudo_map.keys(), key=len, reverse=True):
        text = text.replace(pseudo, pseudo_map[pseudo])
    return text


def _build_user_prompt(role, question, facts):
    return (
        f"Роль пользователя: {role}\n"
        f"Описание роли: {build_role_description(role)}\n\n"
        f"Вопрос пользователя: {question}\n\n"
        f"DATA (обезличенные факты по команде):\n"
        f"{json.dumps(facts, ensure_ascii=False, indent=2)}"
    )


def _fallback_answer(role, question, data, selected_employee_id=None):
    if role == "Сотрудник" and selected_employee_id is not None:
        employees = normalize_id_columns(data["employees"])
        employee_metrics = filter_employee_rows(data["analytics_metrics"], selected_employee_id)
        if not employee_metrics.empty:
            return build_employee_recommendation(employee_metrics.iloc[0], employees)
        return "По выбранному сотруднику не найдены метрики."
    return answer_for_role(role, question, data)


def llm_answer(role, question, data, selected_employee_id=None):
    """Главная точка входа: GigaChat с обезличиванием и fallback на встроенную логику."""
    if not _gigachat_configured():
        return _fallback_answer(role, question, data, selected_employee_id), "fallback"

    try:
        facts, pseudo_map = collect_facts_for_role(role, data, selected_employee_id)
        user_prompt = _build_user_prompt(role, question, facts)
        raw_answer = _gigachat_call(GIGACHAT_SYSTEM_PROMPT, user_prompt)
        return deanonymize(raw_answer, pseudo_map), "gigachat"
    except requests.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "?"
        st.warning(f"GigaChat ответил кодом {status}. Используем встроенную логику.")
    except requests.RequestException as exc:
        st.warning(f"Не удалось обратиться к GigaChat ({exc.__class__.__name__}). Используем встроенную логику.")
    except Exception as exc:
        st.warning(f"Ошибка при вызове GigaChat: {exc}. Используем встроенную логику.")

    return _fallback_answer(role, question, data, selected_employee_id), "fallback"


def answer_for_role(role, question, data):
    employees = data["employees"]
    metrics = normalize_id_columns(data["analytics_metrics"])
    question_lower = question.lower()

    if "кто" in question_lower and ("перегруж" in question_lower or "риск" in question_lower):
        top = get_top_problem_employees(metrics, employees)

        if top.empty:
            return "Не удалось определить сотрудников с высоким риском: в данных нет нужных метрик."

        rows = []
        for _, row in top.head(5).iterrows():
            risk = row.get("risk_score", "—")
            load = row.get("workload_score", row.get("load_ratio", "—"))
            rows.append(f"• {row['full_name']} — risk score: {risk}, загрузка: {load}")

        return (
            "Сотрудники, на которых стоит обратить внимание:\n\n"
            + "\n".join(rows)
            + "\n\nВозможное решение:\n\n"
            + "• По сотрудникам с высоким risk score проверить причину риска.\n\n"
            + "• Если причина в перегрузке — перераспределить задачи.\n\n"
            + "• Если причина в устаревшем графике — отправить запрос на обновление.\n\n"
            + "• Если причина в конфликтах календаря — перенести регулярные встречи в доступное окно."
        )

    if "график" in question_lower and ("обнов" in question_lower or "устар" in question_lower):
        if metrics.empty or "days_since_update" not in metrics.columns:
            return "В данных нет показателя days_since_update для оценки устаревших графиков."

        outdated = metrics[safe_numeric(metrics["days_since_update"]) > 60].copy()

        if outdated.empty:
            return "Сотрудников с графиком старше 60 дней не найдено."

        names = [get_employee_name(row, employees) for _, row in outdated.head(10).iterrows()]

        return (
            "График нужно проверить у следующих сотрудников:\n\n"
            + "\n".join([f"• {name}" for name in names])
            + "\n\nВозможное решение:\n\n"
            + "• Отправить этим сотрудникам запрос на подтверждение графика.\n\n"
            + "• HR должен сверить их графики с отпусками, больничными и форматом работы.\n\n"
            + "• После подтверждения обновить дату актуализации в системе."
        )

    if "лидер" in question_lower or "скор" in question_lower or "эффектив" in question_lower:
        score = compute_general_score(data)
        if score.empty:
            return "Недостаточно данных для расчета общего скора."

        leaders = score.head(5)
        lines = [f"• {row['full_name']} — общий скор: {row['general_score']}" for _, row in leaders.iterrows()]

        return (
            "Лидеры по общему скору:\n\n"
            + "\n".join(lines)
            + "\n\nВозможное решение:\n\n"
            + "• Использовать лидерборд как элемент командной мотивации.\n\n"
            + "• Поощрять сотрудников с высоким вкладом.\n\n"
            + "• При этом отдельно контролировать переработки, чтобы соревнование не приводило к перегрузке."
        )

    if role == "Сотрудник":
        if metrics.empty:
            return "Нет данных по сотруднику."
        return build_employee_recommendation(metrics.iloc[0], employees)

    return build_team_summary(metrics, employees)

def build_role_description(role):
    descriptions = {
        "Сотрудник": "Видит личный график, трудозатраты, рекомендации и запросы на обновление данных.",
        "Администратор": "Настраивает пользователей, роли, источники данных и права доступа.",
        "HR": "Контролирует актуальность графиков, форматы работы, отпуска, больничные и часовые пояса.",
        "Проектный менеджер": "Связывает рабочее время с задачами, спринтами, дедлайнами и загрузкой команды.",
        "Аналитик": "Анализирует BI-дашборды, метрики, перегрузки, конфликты и закономерности.",
        "Руководитель": "Получает итоговую управленческую картину и принимает решения по нагрузке и графикам."
    }
    return descriptions.get(role, "")

def render_bi_dashboards(data):
    employees = normalize_id_columns(data["employees"])
    metrics = add_employee_names(data["analytics_metrics"], employees)

    st.header("BI Dashboard")
    st.caption("Сначала данные агрегируются по всем сотрудникам, затем выводятся в аналитике под конкретные роли.")

    if metrics.empty:
        st.warning("Лист analytics_metrics не найден или пустой.")
        return

    load_col = "workload_score" if "workload_score" in metrics.columns else "load_ratio" if "load_ratio" in metrics.columns else None

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.metric("Сотрудников", len(employees) if not employees.empty else len(metrics))
    with c2:
        st.metric("Средний риск", safe_mean(metrics, "risk_score"))
    with c3:
        st.metric("Средняя загрузка", safe_mean(metrics, load_col) if load_col else "—")
    with c4:
        st.metric("Конфликтов", int(safe_sum(metrics, "conflict_count")) if "conflict_count" in metrics.columns else 0)
    with c5:
        outdated = int((safe_numeric(metrics["days_since_update"]) > 60).sum()) if "days_since_update" in metrics.columns else 0
        st.metric("График устарел", outdated)

    st.divider()

    left, right = st.columns(2)

    with left:
        chart = metrics.copy()
        if "risk_score" in chart.columns:
            chart["risk_score"] = safe_numeric(chart["risk_score"])
            chart = chart.sort_values("risk_score", ascending=False).head(12)
            render_rainbow_bar_chart(chart, "full_name", "risk_score", "Risk score по сотрудникам", 1)

    with right:
        if load_col:
            chart = metrics.copy()
            chart[load_col] = safe_numeric(chart[load_col])
            chart = chart.sort_values(load_col, ascending=False).head(12)
            render_rainbow_bar_chart(chart, "full_name", load_col, "Загрузка по сотрудникам", 1)

    st.divider()

    left, right = st.columns(2)

    with left:
        if "conflict_count" in metrics.columns:
            chart = metrics.copy()
            chart["conflict_count"] = safe_numeric(chart["conflict_count"])
            chart = chart.sort_values("conflict_count", ascending=False).head(12)
            render_rainbow_bar_chart(chart, "full_name", "conflict_count", "Конфликты встреч")

    with right:
        if "days_since_update" in metrics.columns:
            chart = metrics.copy()
            chart["days_since_update"] = safe_numeric(chart["days_since_update"])
            chart = chart.sort_values("days_since_update", ascending=False).head(12)
            render_rainbow_bar_chart(chart, "full_name", "days_since_update", "Дни с последнего обновления графика")

    st.divider()

    st.subheader("Тепловая карта доступности команды")
    render_heatmap_html(data["bi_heatmap"])

    st.divider()

    render_color_table(
        get_top_problem_employees(metrics, employees, 10),
        title="Таблица проблемных сотрудников",
        risk_column="risk_score"
    )

def render_leaderboard_dashboard(data, role):
    score_df = compute_general_score(data)

    st.header("Лидерборд эффективности")
    st.caption("Общий скор агрегирует выполненные задачи, проектные часы, переработки и дополнительные активности.")

    if score_df.empty:
        st.warning("Недостаточно данных для построения лидерборда. Проверьте, что в employees есть id или employee_id.")
        return

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.metric("Средний общий скор", round(safe_numeric(score_df["general_score"]).mean(), 1))
    with c2:
        st.metric("Выполнено задач", int(safe_numeric(score_df["tasks_done"]).sum()))
    with c3:
        st.metric("Часы на проекты", round(safe_numeric(score_df["project_hours"]).sum(), 1))
    with c4:
        st.metric("Переработки", round(safe_numeric(score_df["overtime_hours"]).sum(), 1))

    st.info(get_role_analytics(role, data))

    st.subheader("Формула общего скора")
    st.code("general_score = 35% * сделанные таски + 25% * проектные часы + 20% * активности - 20% * переработки", language="text")

    render_color_table(
        score_df.head(25),
        columns=["rank", "full_name", "department", "general_score", "tasks_done", "project_hours", "overtime_hours", "activity_points", "risk_score"],
        title="Лидерборд сотрудников",
        score_column="general_score",
        risk_column="risk_score"
    )

    st.divider()

    render_rainbow_bar_chart(score_df.head(15), "full_name", "general_score", "Общий скор сотрудников", 100)

    st.divider()

    st.subheader("Разбивка по департаментам")

    if "department" in score_df.columns:
        dept = score_df.groupby("department", as_index=False).agg(
            avg_score=("general_score", "mean"),
            tasks_done=("tasks_done", "sum"),
            project_hours=("project_hours", "sum"),
            overtime_hours=("overtime_hours", "sum"),
            activity_points=("activity_points", "sum")
        )

        dept["avg_score"] = dept["avg_score"].round(1)
        dept = dept.sort_values("avg_score", ascending=False)

        render_color_table(
            dept,
            columns=["department", "avg_score", "tasks_done", "project_hours", "overtime_hours", "activity_points"],
            title="Таблица по департаментам",
            score_column="avg_score"
        )

        render_rainbow_bar_chart(dept, "department", "avg_score", "Скор по департаментам", 100)

    st.divider()

    role_rows = pd.DataFrame([
        {"role": "Сотрудник", "analytics": "личный скор, задачи, часы, рекомендации", "hidden": "чувствительные HR-данные других сотрудников"},
        {"role": "HR", "analytics": "актуальность графиков, отсутствия, часовые пояса, устаревшие профили", "hidden": "детальная проектная декомпозиция без необходимости"},
        {"role": "Проектный менеджер", "analytics": "задачи, проектные часы, загрузка, переработки, доступность команды", "hidden": "часть кадровых данных"},
        {"role": "Аналитик", "analytics": "агрегированные метрики, департаменты, аномалии, эффективность", "hidden": "личные чувствительные данные"},
        {"role": "Руководитель", "analytics": "итоговый лидерборд, риски, рекомендации, управленческие сигналы", "hidden": "технические настройки системы"},
        {"role": "Администратор", "analytics": "качество данных, роли, источники, связи таблиц", "hidden": "управленческие оценки без доступа"}
    ])

    render_color_table(role_rows, title="Ролевая аналитика после агрегации")


def filter_employee_rows(df, employee_id):
    df = normalize_id_columns(df)

    if df.empty or "employee_id" not in df.columns:
        return df

    return df[df["employee_id"].astype(str) == str(employee_id)].copy()


PRIORITY_LABELS = {
    "critical": "критический",
    "high": "высокий",
    "medium": "средний",
    "low": "низкий",
}

PRIORITY_COLORS = {
    "critical": "#EF4444",
    "high": "#E47419",
    "medium": "#FACC15",
    "low": "#29B171",
}

PRIORITY_LABEL_COLORS = {
    "критический": "#EF4444",
    "высокий": "#E47419",
    "средний": "#FACC15",
    "низкий": "#29B171",
}

ABSENCE_LABELS = {
    "vacation": "отпуск",
    "sick_leave": "больничный",
    "business_trip": "командировка",
    "personal_hours": "личные часы",
}

ABSENCE_COLORS = {
    "отпуск": "#EF4444",
    "больничный": "#29B171",
    "командировка": "#3882F6",
    "личные часы": "#E47419",
}


def priority_label(value):
    return PRIORITY_LABELS.get(str(value).strip().lower(), str(value) if str(value).strip() else "не указан")


def priority_color(value):
    return PRIORITY_COLORS.get(str(value).strip().lower(), "#94A3B8")


def absence_label(value):
    return ABSENCE_LABELS.get(str(value).strip().lower(), str(value) if str(value).strip() else "не указано")


def risk_level_label(value):
    try:
        value = float(value)
    except Exception:
        return "нет данных"
    if value >= 0.75:
        return "критический"
    if value >= 0.55:
        return "высокий"
    if value >= 0.35:
        return "средний"
    return "низкий"


def render_section_title(icon, title, subtitle=None):
    subtitle_html = f'<p>{escape_html(subtitle)}</p>' if subtitle else ""
    st.markdown(
        f'<div class="streamlit-panel-title"><span class="heading-icon accent">{escape_html(icon)}</span>'
        f'<div><h2>{escape_html(title)}</h2>{subtitle_html}</div></div>',
        unsafe_allow_html=True,
    )


def render_priority_table(df, columns, title=None):
    if title:
        st.subheader(title)
    if df.empty:
        st.info("Данных для отображения нет.")
        return
    view = df[[col for col in columns if col in df.columns]].copy()
    if "приоритет" in view.columns:
        def row_style(row):
            color = priority_color(row.get("_priority_key", row.get("приоритет", ""))) if "_priority_key" in row.index else {
                "критический": "#EF4444",
                "высокий": "#E47419",
                "средний": "#FACC15",
                "низкий": "#29B171",
            }.get(str(row.get("приоритет", "")).lower(), "#94A3B8")
            return [f"border-left: 5px solid {color};" if col == "приоритет" else "" for col in row.index]
        styler = view.style.apply(row_style, axis=1)
        st.dataframe(styler, use_container_width=True, hide_index=True)
    else:
        st.dataframe(view, use_container_width=True, hide_index=True)


def render_colored_bar_chart(df, x, y, color_col=None, title=None, horizontal=False, color_scale=None):
    if df.empty or x not in df.columns or y not in df.columns:
        st.info("Недостаточно данных для графика.")
        return
    chart = alt.Chart(df).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
        tooltip=list(df.columns)
    )
    if horizontal:
        chart = chart.encode(
            x=alt.X(f"{y}:Q", title=None),
            y=alt.Y(f"{x}:N", title=None, sort="-x"),
        )
    else:
        chart = chart.encode(
            x=alt.X(f"{x}:N", title=None, sort="-y"),
            y=alt.Y(f"{y}:Q", title=None),
        )
    if color_col and color_col in df.columns:
        if color_scale:
            chart = chart.encode(color=alt.Color(f"{color_col}:N", scale=alt.Scale(domain=list(color_scale.keys()), range=list(color_scale.values())), legend=alt.Legend(title=None)))
        else:
            chart = chart.encode(color=alt.Color(f"{color_col}:N", legend=alt.Legend(title=None)))
    else:
        chart = chart.encode(color=alt.value("#7BD7FF"))
    if title:
        chart = chart.properties(title=title)
    st.altair_chart(chart, use_container_width=True)


def render_donut_chart(df, theta, color, color_scale=None, title=None):
    if df.empty or theta not in df.columns or color not in df.columns:
        st.info("Недостаточно данных для диаграммы.")
        return
    if color_scale:
        color_encoding = alt.Color(
            f"{color}:N",
            scale=alt.Scale(domain=list(color_scale.keys()), range=list(color_scale.values())),
            legend=alt.Legend(title=None),
        )
    else:
        color_encoding = alt.Color(f"{color}:N", legend=alt.Legend(title=None))
    chart = alt.Chart(df).mark_arc(innerRadius=55, outerRadius=95).encode(
        theta=alt.Theta(f"{theta}:Q"),
        color=color_encoding,
        tooltip=list(df.columns),
    )
    if title:
        chart = chart.properties(title=title)
    st.altair_chart(chart, use_container_width=True)


def prepare_employee_workload(data, employee_id):
    workload = filter_employee_rows(data["workload_logs"], employee_id)
    if workload.empty:
        return workload
    workload = normalize_columns(workload)
    result = pd.DataFrame()
    result["дата"] = pd.to_datetime(workload.get("log_date"), errors="coerce").dt.strftime("%d.%m.%Y")
    result["задача"] = workload.get("task_name", "—")
    result["тип задачи"] = workload.get("task_template", "—")
    result["часы"] = safe_numeric(workload.get("spent_hours", pd.Series(dtype=float)))
    priority_key = workload.get("task_priority", pd.Series(dtype=str)).astype(str).str.lower()
    result["_priority_key"] = priority_key
    result["приоритет"] = priority_key.map(PRIORITY_LABELS).fillna(priority_key)
    result["источник"] = workload.get("source", "—")
    return result.sort_values("дата", ascending=False)


def prepare_project_tasks(data):
    workload = add_employee_names(data["workload_logs"], data["employees"])
    metrics = add_employee_names(data["analytics_metrics"], data["employees"])
    if workload.empty:
        return pd.DataFrame()
    workload = normalize_columns(workload)
    workload["log_date"] = pd.to_datetime(workload.get("log_date"), errors="coerce")
    priority = workload.get("task_priority", pd.Series(dtype=str)).astype(str).str.lower()
    due_days = priority.map({"critical": 1, "high": 3, "medium": 7, "low": 14}).fillna(7)
    workload["дедлайн"] = workload["log_date"] + pd.to_timedelta(due_days, unit="D")
    today = pd.Timestamp(date.today())
    workload["дней до дедлайна"] = (workload["дедлайн"] - today).dt.days
    workload["статус"] = workload["дней до дедлайна"].apply(
        lambda x: "просрочено" if pd.notna(x) and x < 0 else "сегодня" if x == 0 else "скоро дедлайн" if pd.notna(x) and x <= 2 else "в норме"
    )
    workload["_priority_key"] = priority
    workload["приоритет"] = priority.map(PRIORITY_LABELS).fillna(priority)
    workload["дата"] = workload["log_date"].dt.strftime("%d.%m.%Y")
    workload["дедлайн"] = workload["дедлайн"].dt.strftime("%d.%m.%Y")
    workload["трудозатраты"] = safe_numeric(workload.get("spent_hours", pd.Series(dtype=float)))
    metrics_small = metrics[["employee_id", "load_ratio", "risk_score"]] if not metrics.empty and "employee_id" in metrics.columns else pd.DataFrame()
    if not metrics_small.empty:
        workload = workload.merge(metrics_small, on="employee_id", how="left", suffixes=("", "_metric"))
    workload["загрузка"] = safe_numeric(workload["load_ratio"]) if "load_ratio" in workload.columns else 0
    workload["риск"] = safe_numeric(workload["risk_score"]) if "risk_score" in workload.columns else 0
    workload["рекомендация"] = workload.apply(
        lambda row: f"У сотрудника {row.get('full_name', row.get('employee_id', '—'))} высокая загрузка и задача с дедлайном через {row.get('дней до дедлайна', '—')} дн. Рекомендуется перераспределить часть задач или подключить другого исполнителя."
        if row.get("загрузка", 0) >= 0.7 or row.get("статус") in ["просрочено", "сегодня", "скоро дедлайн"]
        else "Контроль в плановом режиме.",
        axis=1,
    )
    return workload


def build_data_quality_table(data):
    employees = normalize_id_columns(data["employees"])
    schedules = normalize_id_columns(data["schedules"])
    metrics = normalize_id_columns(data["analytics_metrics"])
    events = normalize_id_columns(data["calendar_events"])

    rows = []
    schedule_ids = set(schedules["employee_id"].astype(str)) if "employee_id" in schedules.columns else set()
    duplicate_ids = set(employees.loc[employees["employee_id"].astype(str).duplicated(keep=False), "employee_id"].astype(str)) if "employee_id" in employees.columns else set()

    for _, row in employees.iterrows():
        emp_id = str(row.get("employee_id", row.get("id", "")))
        checks = []
        if not str(row.get("email", "")).strip() or str(row.get("email", "")).lower() == "nan":
            checks.append(("ошибка", "пустая почта"))
        if emp_id not in schedule_ids:
            checks.append(("ошибка", "отсутствующий график"))
        if emp_id in duplicate_ids:
            checks.append(("ошибка", "дублирующийся ID"))

        metric = metrics[metrics["employee_id"].astype(str) == emp_id] if "employee_id" in metrics.columns else pd.DataFrame()
        days_since = None
        risk = None
        conflict = None
        if not metric.empty:
            m = metric.iloc[0]
            days_since = _safe_value(m.get("days_since_update"))
            risk = _safe_value(m.get("risk_score"))
            conflict = _safe_value(m.get("conflict_count"))
            if safe_numeric(pd.Series([m.get("days_since_update")])).iloc[0] > 60:
                checks.append(("предупреждение", "устаревший график"))
            if safe_numeric(pd.Series([m.get("risk_score")])).iloc[0] >= 0.55:
                checks.append(("ошибка", "высокий риск"))
            if safe_numeric(pd.Series([m.get("conflict_count")])).iloc[0] > 0 or safe_numeric(pd.Series([m.get("hr_calendar_mismatch")])).iloc[0] > 0:
                checks.append(("предупреждение", "конфликт календаря и графика"))

        if not checks:
            checks.append(("норма", "критичных проблем нет"))

        worst = "ошибка" if any(level == "ошибка" for level, _ in checks) else "предупреждение" if any(level == "предупреждение" for level, _ in checks) else "норма"
        rows.append({
            "сотрудник": row.get("full_name", emp_id),
            "отдел": row.get("department", "—"),
            "статус": worst,
            "что проверить": "; ".join(reason for _, reason in checks),
            "дней без обновления": days_since,
            "риск": risk,
            "конфликты": conflict,
        })
    return pd.DataFrame(rows)


def style_status_table(df):
    def style_row(row):
        color = {"ошибка": "#EF4444", "предупреждение": "#E47419", "норма": "#29B171"}.get(str(row.get("статус", "")).lower(), "#94A3B8")
        return [f"border-left: 5px solid {color}; background-color: {color}22;" if col == "статус" else "" for col in row.index]
    return df.style.apply(style_row, axis=1)

def get_employee_options(employees):
    employees = normalize_id_columns(employees)

    if employees.empty or "employee_id" not in employees.columns:
        return []

    options = []

    for _, row in employees.iterrows():
        name = row.get("full_name", row.get("employee_id"))
        emp_id = row.get("employee_id")
        options.append((str(emp_id), f"{name} ({emp_id})"))

    return options

def render_employee_view(data, selected_employee_id):
    employees = normalize_id_columns(data["employees"])
    metrics = filter_employee_rows(data["analytics_metrics"], selected_employee_id)
    schedules = filter_employee_rows(data["schedules"], selected_employee_id)
    workload = filter_employee_rows(data["workload_logs"], selected_employee_id)
    events = filter_employee_rows(data["calendar_events"], selected_employee_id)
    recommendations = filter_employee_rows(data["recommendations"], selected_employee_id)
    score_df = compute_general_score(data)
    my_score = score_df[score_df["employee_id"].astype(str) == str(selected_employee_id)].copy() if not score_df.empty else pd.DataFrame()

    employee_row = employees[employees["employee_id"].astype(str) == str(selected_employee_id)]

    st.header("Личный кабинет сотрудника")
    st.caption("Сотрудник видит только свои данные: график, трудозатраты, личный скор и рекомендации.")

    if not employee_row.empty:
        emp = employee_row.iloc[0]
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Сотрудник", emp.get("full_name", selected_employee_id))
        with c2:
            st.metric("Департамент", emp.get("department", "—"))
        with c3:
            st.metric("Формат", emp.get("work_format", "—"))
        with c4:
            st.metric("Часовой пояс", emp.get("timezone", "—"))

    if not my_score.empty:
        row = my_score.iloc[0]
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Мой общий скор", row.get("general_score", "—"))
        with c2:
            st.metric("Выполнено задач", row.get("tasks_done", "—"))
        with c3:
            st.metric("Часы на проект", row.get("project_hours", "—"))
        with c4:
            st.metric("Переработки", row.get("overtime_hours", "—"))

        render_color_table(
            my_score,
            columns=["rank", "full_name", "department", "general_score", "tasks_done", "project_hours", "overtime_hours", "activity_points", "risk_score"],
            title="Моя аналитика",
            score_column="general_score",
            risk_column="risk_score"
        )

    st.divider()

    left, right = st.columns(2)

    with left:
        st.subheader("Мой график")
        if schedules.empty:
            st.info("Данные графика не найдены.")
        else:
            st.dataframe(schedules, use_container_width=True)

    with right:
        st.subheader("Мои рекомендации")
        if not metrics.empty:
            st.info(build_employee_recommendation(metrics.iloc[0], employees))
        elif not recommendations.empty:
            st.dataframe(recommendations, use_container_width=True)
        else:
            st.info("Рекомендации не найдены.")

    st.divider()

    tab_a, tab_b = st.tabs(["Мои трудозатраты", "Мои события"])

    with tab_a:
        if workload.empty:
            st.info("Трудозатраты не найдены.")
        else:
            st.dataframe(workload, use_container_width=True)

    with tab_b:
        if events.empty:
            st.info("События календаря не найдены.")
        else:
            st.dataframe(events, use_container_width=True)

def render_admin_view(data):
    st.header("Панель администратора")
    st.caption("Администратор быстро видит качество данных, проблемные записи и причины ошибок.")

    employees = normalize_id_columns(data["employees"])
    schedules = normalize_id_columns(data["schedules"])
    metrics = normalize_id_columns(data["analytics_metrics"])
    workload = normalize_id_columns(data["workload_logs"])
    events = normalize_id_columns(data["calendar_events"])
    quality = build_data_quality_table(data)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.metric("Пользователей", len(employees))
    with c2:
        st.metric("Графиков", len(schedules))
    with c3:
        st.metric("Событий", len(events))
    with c4:
        st.metric("Записей трудозатрат", len(workload))

    st.divider()
    render_section_title("⚠", "Ошибки качества данных", "Красный — ошибка, оранжевый — предупреждение, зеленый — норма.")

    f1, f2, f3 = st.columns([1, 1, 1])
    with f1:
        dep_options = ["Все отделы"] + sorted(quality["отдел"].dropna().astype(str).unique().tolist()) if not quality.empty and "отдел" in quality.columns else ["Все отделы"]
        dep_filter = st.selectbox("Отдел", dep_options, key="admin_quality_department")
    with f2:
        status_filter = st.selectbox("Статус ошибки", ["Все статусы", "ошибка", "предупреждение", "норма"], key="admin_quality_status")
    with f3:
        only_problem = st.toggle("Показать только проблемные записи", value=True, key="admin_quality_only_problem")

    quality_view = quality.copy()
    if dep_filter != "Все отделы":
        quality_view = quality_view[quality_view["отдел"].astype(str) == dep_filter]
    if status_filter != "Все статусы":
        quality_view = quality_view[quality_view["статус"].astype(str) == status_filter]
    if only_problem:
        quality_view = quality_view[quality_view["статус"].astype(str) != "норма"]

    q1, q2, q3 = st.columns(3)
    with q1:
        st.metric("Ошибки", int((quality["статус"] == "ошибка").sum()) if not quality.empty else 0)
    with q2:
        st.metric("Предупреждения", int((quality["статус"] == "предупреждение").sum()) if not quality.empty else 0)
    with q3:
        st.metric("Норма", int((quality["статус"] == "норма").sum()) if not quality.empty else 0)

    if quality_view.empty:
        st.info("По выбранным фильтрам проблем не найдено.")
    else:
        st.dataframe(style_status_table(quality_view.head(120)), use_container_width=True, hide_index=True)

    st.divider()
    render_section_title("▥", "Сводка по таблицам")
    quality_rows = []
    for name, df in [
        ("Сотрудники", employees),
        ("Графики", schedules),
        ("События", events),
        ("Трудозатраты", workload),
        ("Метрики", metrics)
    ]:
        quality_rows.append({
            "таблица": name,
            "записей": len(df),
            "есть employee_id": "employee_id" in df.columns,
            "пустых значений": int(df.isna().sum().sum()) if not df.empty else 0
        })
    st.dataframe(pd.DataFrame(quality_rows), use_container_width=True, hide_index=True)

    tab1, tab2, tab3, tab4 = st.tabs(["Пользователи", "Графики", "События", "Метрики"])

    with tab1:
        st.dataframe(employees, use_container_width=True)
    with tab2:
        st.dataframe(schedules, use_container_width=True)
    with tab3:
        st.dataframe(events, use_container_width=True)
    with tab4:
        st.dataframe(metrics, use_container_width=True)

def render_hr_view(data):
    st.header("HR-аналитика")
    st.caption("HR быстро видит, кому надо обновить график, где есть отсутствия и кадровые исключения.")

    employees = normalize_id_columns(data["employees"])
    metrics = add_employee_names(data["analytics_metrics"], employees)
    schedules = add_employee_names(data["schedules"], employees)
    absences = add_employee_names(data["absences"], employees)

    outdated = pd.DataFrame()

    if not metrics.empty and "days_since_update" in metrics.columns:
        outdated = metrics[safe_numeric(metrics["days_since_update"]) > 60].copy()

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.metric("Сотрудников", len(employees))
    with c2:
        st.metric("Устаревшие графики", len(outdated))
    with c3:
        st.metric("Отсутствия", len(absences))
    with c4:
        risk_count = int((safe_numeric(metrics["risk_score"]) >= 0.55).sum()) if "risk_score" in metrics.columns else 0
        st.metric("Высокий риск", risk_count)

    st.info(get_role_analytics("HR", data))

    st.divider()
    render_section_title("▦", "Графики к подтверждению", "Слева список сотрудников, справа быстрый график по отделам.")
    sort_option = st.selectbox("Сортировка HR-таблицы", ["по количеству дней с последнего обновления", "по риску", "по отделу"], key="hr_sort_outdated")
    outdated_view = outdated.copy()
    if not outdated_view.empty:
        if sort_option == "по риску" and "risk_score" in outdated_view.columns:
            outdated_view = outdated_view.sort_values("risk_score", ascending=False)
        elif sort_option == "по отделу" and "department" in outdated_view.columns:
            outdated_view = outdated_view.sort_values("department")
        elif "days_since_update" in outdated_view.columns:
            outdated_view = outdated_view.sort_values("days_since_update", ascending=False)

    left, right = st.columns([1.2, 1])
    with left:
        render_color_table(
            outdated_view.head(40),
            columns=["full_name", "department", "days_since_update", "risk_score", "risk_status", "behavior_segment"],
            title="Кому нужно подтвердить график",
            risk_column="risk_score"
        )
    with right:
        if not outdated_view.empty and "department" in outdated_view.columns:
            dept_outdated = outdated_view.groupby("department", as_index=False).size().rename(columns={"size": "сотрудники"})
            render_colored_bar_chart(dept_outdated, "department", "сотрудники", title="Устаревшие графики по отделам", horizontal=True)
        else:
            st.info("Нет устаревших графиков.")

    st.divider()
    render_section_title("◷", "Отсутствия и причины", "Цвет показывает причину отсутствия.")
    if absences.empty:
        st.info("Отсутствия не найдены.")
    else:
        abs_view = normalize_columns(absences).copy()
        abs_view["причина"] = abs_view["absence_type"].map(absence_label) if "absence_type" in abs_view.columns else "не указано"
        abs_sort = st.selectbox("Сортировка отсутствий", ["по причине отсутствия", "по отделу"], key="hr_absence_sort")
        if abs_sort == "по отделу" and "department" in abs_view.columns:
            abs_view = abs_view.sort_values("department")
        else:
            abs_view = abs_view.sort_values("причина")
        left, right = st.columns([1.15, 1])
        with left:
            table = abs_view.rename(columns={
                "full_name": "сотрудник",
                "department": "отдел",
                "start_date": "начало",
                "end_date": "окончание",
                "comment": "комментарий",
            })
            st.dataframe(table[[col for col in ["сотрудник", "отдел", "причина", "начало", "окончание", "комментарий"] if col in table.columns]].head(80), use_container_width=True, hide_index=True)
        with right:
            reason = abs_view.groupby("причина", as_index=False).size().rename(columns={"size": "записей"})
            render_colored_bar_chart(reason, "причина", "записей", "причина", "Причины отсутствий", color_scale=ABSENCE_COLORS)

    st.divider()
    render_section_title("△", "Высокий риск и часовые пояса")
    left, right = st.columns(2)
    with left:
        high_risk = metrics[safe_numeric(metrics["risk_score"]) >= 0.55].copy() if not metrics.empty and "risk_score" in metrics.columns else pd.DataFrame()
        render_color_table(
            high_risk.sort_values("risk_score", ascending=False).head(30) if not high_risk.empty else high_risk,
            columns=["full_name", "department", "risk_score", "days_since_update", "behavior_segment"],
            title="Сотрудники с высоким риском",
            risk_column="risk_score",
        )
    with right:
        if not employees.empty and "timezone" in employees.columns:
            tz = employees.groupby("timezone", as_index=False).size().rename(columns={"size": "сотрудники"})
            render_colored_bar_chart(tz, "timezone", "сотрудники", title="Часовые пояса", horizontal=True)
        else:
            st.info("Часовые пояса не найдены.")

def render_pm_view(data):
    st.header("Проектная аналитика")
    st.caption("Проектный менеджер видит связь задач, дедлайнов, трудозатрат и загрузки команды.")

    tasks = prepare_project_tasks(data)
    if tasks.empty:
        st.info("Задачи и трудозатраты не найдены.")
        return

    overdue = tasks[tasks["статус"] == "просрочено"].copy()
    hurry = tasks[tasks["статус"].isin(["просрочено", "сегодня", "скоро дедлайн"])].copy()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Задач", len(tasks))
    with c2:
        st.metric("Просрочено", len(overdue))
    with c3:
        st.metric("Скоро дедлайн", len(hurry))
    with c4:
        st.metric("Трудозатраты", round(safe_sum(tasks, "трудозатраты"), 1))

    render_section_title("⚠", "Кого поторопить", "Задачи с близким или просроченным дедлайном и рекомендацией.")
    hurry_view = hurry.sort_values(["дней до дедлайна", "загрузка"], ascending=[True, False]).head(20)
    if hurry_view.empty:
        st.success("Критичных дедлайнов сейчас нет.")
    else:
        table = hurry_view.rename(columns={
            "full_name": "сотрудник",
            "task_name": "задача",
        })
        render_priority_table(
            table,
            ["сотрудник", "задача", "дедлайн", "дней до дедлайна", "статус", "приоритет", "трудозатраты", "загрузка", "рекомендация"],
        )

    st.divider()
    render_section_title("▥", "Проектные графики")
    g1, g2 = st.columns(2)
    with g1:
        dept_deadline = tasks.groupby("department", as_index=False)["дней до дедлайна"].mean().rename(columns={"дней до дедлайна": "среднее дней до дедлайна"})
        render_colored_bar_chart(dept_deadline, "department", "среднее дней до дедлайна", title="Сколько дней осталось до дедлайна по отделам", horizontal=True)
    with g2:
        overdue_dept = overdue.groupby("department", as_index=False).size().rename(columns={"size": "просрочено"})
        render_colored_bar_chart(overdue_dept, "department", "просрочено", title="Просроченные задачи по отделам", horizontal=True)

    g3, g4 = st.columns(2)
    with g3:
        priority = tasks.groupby(["приоритет", "_priority_key"], as_index=False).size().rename(columns={"size": "задачи"})
        render_colored_bar_chart(priority, "приоритет", "задачи", "приоритет", "Задачи по приоритетам", color_scale=PRIORITY_LABEL_COLORS)
    with g4:
        by_type = tasks.groupby("task_template", as_index=False)["трудозатраты"].sum().sort_values("трудозатраты", ascending=False)
        render_colored_bar_chart(by_type, "task_template", "трудозатраты", title="Трудозатраты по типам задач", horizontal=True)

    g5, _ = st.columns([1, 1])
    with g5:
        load = tasks.groupby(["full_name", "department"], as_index=False)["загрузка"].mean().sort_values("загрузка", ascending=False).head(15)
        render_colored_bar_chart(load, "full_name", "загрузка", "department", "Загрузка сотрудников по проектам", horizontal=True)

    st.divider()
    render_leaderboard_dashboard(data, "Проектный менеджер")

    st.divider()
    render_section_title("☑", "Список задач по сотрудникам")
    task_table = tasks.rename(columns={"full_name": "сотрудник", "department": "отдел", "task_name": "задача", "task_template": "тип задачи"})
    render_priority_table(
        task_table.sort_values(["дней до дедлайна", "загрузка"], ascending=[True, False]).head(100),
        ["сотрудник", "отдел", "задача", "тип задачи", "дедлайн", "дней до дедлайна", "статус", "приоритет", "трудозатраты", "загрузка"],
    )

def render_analyst_view(data):
    st.header("Аналитическая панель")
    st.caption("Аналитик собирает общую картину по департаментам, эффективности, трудоемкости, ошибкам, HR и проектным срокам.")

    employees = normalize_id_columns(data["employees"])
    metrics = add_employee_names(data["analytics_metrics"], employees)
    workload = add_employee_names(data["workload_logs"], employees)
    absences = add_employee_names(data["absences"], employees)
    tasks = prepare_project_tasks(data)
    score = compute_general_score(data)
    quality = build_data_quality_table(data)

    render_section_title("▥", "Аналитика по департаментам", "Где выше риск, загрузка и просрочки.")
    c1, c2 = st.columns(2)
    with c1:
        if not metrics.empty and "department" in metrics.columns:
            dept = metrics.groupby("department", as_index=False).agg({"risk_score": "mean", "load_ratio": "mean"}).round(2)
            render_colored_bar_chart(dept, "department", "risk_score", title="Средний риск по департаментам", horizontal=True)
            st.caption("Смотрите отделы с максимальным риском: там нужны проверка графиков и причин перегруза.")
    with c2:
        if not metrics.empty and "department" in metrics.columns:
            render_colored_bar_chart(dept, "department", "load_ratio", title="Средняя загрузка по департаментам", horizontal=True)
            st.caption("Высокая загрузка вместе с высоким риском — первый кандидат на управленческое действие.")

    st.divider()
    render_section_title("✓", "Аналитика эффективности", "Кто показывает результат и где он распределен по отделам.")
    c1, c2 = st.columns(2)
    with c1:
        if not score.empty:
            top = score.sort_values("general_score", ascending=False).head(15)
            render_colored_bar_chart(top, "full_name", "general_score", "department", "Топ сотрудников по результативности", horizontal=True)
            st.caption("Высокая результативность полезна только если не сопровождается постоянной переработкой.")
    with c2:
        if not score.empty and "department" in score.columns:
            dept_score = score.groupby("department", as_index=False)["general_score"].mean().round(1)
            render_colored_bar_chart(dept_score, "department", "general_score", title="Средняя оценка по отделам", horizontal=True)
            st.caption("Отделы ниже среднего стоит проверять по трудоемкости и качеству данных.")

    st.divider()
    render_section_title("◷", "Аналитика трудоемкости", "Какие типы задач забирают больше всего часов.")
    c1, c2 = st.columns(2)
    with c1:
        if not workload.empty and "task_template" in workload.columns:
            effort = workload.groupby("task_template", as_index=False)["spent_hours"].sum().rename(columns={"spent_hours": "часы"})
            render_colored_bar_chart(effort, "task_template", "часы", title="Трудоемкость по типам задач", horizontal=True)
            st.caption("Самые трудоемкие типы задач помогают увидеть, где нужна автоматизация или перераспределение.")
    with c2:
        if not tasks.empty:
            priority = tasks.groupby(["приоритет", "_priority_key"], as_index=False).size().rename(columns={"size": "задачи"})
            render_donut_chart(priority, "задачи", "приоритет", PRIORITY_LABEL_COLORS, "Задачи по приоритетам")
            st.caption("Круговая диаграмма показывает, насколько команда живет в срочности.")

    st.divider()
    render_section_title("⚠", "Аналитика ошибок данных", "Качество данных доходит до управленческих решений.")
    c1, c2 = st.columns(2)
    with c1:
        status = quality.groupby("статус", as_index=False).size().rename(columns={"size": "записей"})
        render_colored_bar_chart(status, "статус", "записей", "статус", "Статусы качества данных", color_scale={"ошибка": "#EF4444", "предупреждение": "#E47419", "норма": "#29B171"})
        st.caption("Ошибки качества данных надо закрывать до анализа эффективности.")
    with c2:
        reason_rows = []
        for value in quality["что проверить"].dropna().astype(str):
            for item in value.split("; "):
                reason_rows.append(item)
        reasons = pd.Series(reason_rows).value_counts().head(8).reset_index()
        reasons.columns = ["причина", "записей"]
        render_colored_bar_chart(reasons, "причина", "записей", title="Причины ошибок данных", horizontal=True)
        st.caption("Повторяющиеся причины показывают, где нужен системный фикс, а не ручная правка.")

    st.divider()
    render_section_title("HR", "HR-аналитика и отсутствия", "Причины отсутствий и кадровые исключения.")
    c1, c2 = st.columns(2)
    with c1:
        if not absences.empty and "absence_type" in absences.columns:
            abs_view = normalize_columns(absences).copy()
            abs_view["причина"] = abs_view["absence_type"].map(absence_label)
            reason = abs_view.groupby("причина", as_index=False).size().rename(columns={"size": "записей"})
            render_colored_bar_chart(reason, "причина", "записей", "причина", "Причины отсутствий", color_scale=ABSENCE_COLORS)
            st.caption("Всплеск отсутствий по одной причине может менять доступность команды.")
    with c2:
        if not employees.empty and "timezone" in employees.columns:
            tz = employees.groupby("timezone", as_index=False).size().rename(columns={"size": "сотрудники"})
            render_colored_bar_chart(tz, "timezone", "сотрудники", title="Часовые пояса", horizontal=True)
            st.caption("Разные часовые пояса влияют на встречи и календарные конфликты.")

    st.divider()
    render_section_title("△", "Проектные сроки, риски и перегрузки", "Где дедлайны и нагрузка могут сорвать работу.")
    c1, c2 = st.columns(2)
    with c1:
        if not tasks.empty:
            status_tasks = tasks.groupby("статус", as_index=False).size().rename(columns={"size": "задачи"})
            render_colored_bar_chart(status_tasks, "статус", "задачи", title="Статусы задач по дедлайнам")
            st.caption("Просроченные и сегодняшние задачи требуют немедленной реакции.")
    with c2:
        if not metrics.empty:
            risky = metrics.sort_values("risk_score", ascending=False).head(15)
            render_colored_bar_chart(risky, "full_name", "risk_score", "department", "Риски и перегрузки", horizontal=True)
            st.caption("Высокий риск полезно сверять с устаревшими графиками и встречами вне времени.")

    st.divider()
    render_section_title("BI", "Существующие BI-графики")
    render_bi_dashboards(data)
    st.divider()
    render_leaderboard_dashboard(data, "Аналитик")

def render_manager_view(data):
    st.header("Панель руководителя")
    st.caption("Руководитель сразу видит проблему, важность и возможное действие.")

    employees = normalize_id_columns(data["employees"])
    metrics = add_employee_names(data["analytics_metrics"], employees)
    score = compute_general_score(data)
    tasks = prepare_project_tasks(data)
    quality = build_data_quality_table(data)
    recommendations = add_employee_names(data["recommendations"], employees)

    tabs = st.tabs(["Обзор", "Риски", "Эффективность", "Проекты", "Доступность", "Рекомендации"])

    with tabs[0]:
        overview = get_overview_stats(data)
        metric_cards_html([
            ("Высокий риск", int((safe_numeric(metrics["risk_score"]) >= 0.55).sum()) if not metrics.empty and "risk_score" in metrics.columns else 0, "сотрудников", "var(--wts-error)"),
            ("Просроченные задачи", len(tasks[tasks["статус"] == "просрочено"]) if not tasks.empty else 0, "по дедлайнам", "var(--wts-risk)"),
            ("Ошибки данных", int((quality["статус"] == "ошибка").sum()) if not quality.empty else 0, "нужно исправить", "var(--wts-error)"),
            ("Результативность", overview["avg_score"], "средняя оценка", "var(--wts-success)"),
            ("Загрузка", overview["avg_load"], "средняя", "var(--wts-accent)"),
            ("Средний риск", overview["avg_risk"], "risk_score", "var(--wts-risk)"),
        ])
        left, right = st.columns([1.2, 1])
        with left:
            attention = get_top_problem_employees(data["analytics_metrics"], data["employees"], 12)
            render_color_table(attention, columns=["full_name", "department", "risk_score", "conflict_count", "days_since_update", "behavior_segment"], title="Кто требует внимания", risk_column="risk_score")
        with right:
            if not tasks.empty:
                status_tasks = tasks.groupby("статус", as_index=False).size().rename(columns={"size": "задачи"})
                render_donut_chart(status_tasks, "задачи", "статус", title="Задачи по статусам")

    with tabs[1]:
        render_section_title("△", "Риски")
        left, right = st.columns(2)
        with left:
            risky = metrics.sort_values("risk_score", ascending=False).head(15) if not metrics.empty and "risk_score" in metrics.columns else pd.DataFrame()
            render_colored_bar_chart(risky, "full_name", "risk_score", "department", "Топ сотрудников по риску", horizontal=True)
        with right:
            if not metrics.empty and "risk_score" in metrics.columns:
                bins = pd.cut(safe_numeric(metrics["risk_score"]), bins=[0, 0.35, 0.55, 0.75, 1], labels=["низкий", "средний", "высокий", "критический"], include_lowest=True)
                risk_dist = bins.value_counts().rename_axis("уровень").reset_index(name="сотрудники")
                render_donut_chart(risk_dist, "сотрудники", "уровень", {"низкий": "#29B171", "средний": "#FACC15", "высокий": "#E47419", "критический": "#EF4444"}, "Распределение рисков")
        reason_rows = []
        if not metrics.empty:
            for _, row in metrics.iterrows():
                if safe_numeric(pd.Series([row.get("days_since_update")])).iloc[0] > 60:
                    reason_rows.append("устаревший график")
                if safe_numeric(pd.Series([row.get("load_ratio")])).iloc[0] >= 0.7:
                    reason_rows.append("перегрузка")
                if safe_numeric(pd.Series([row.get("conflict_count")])).iloc[0] > 0:
                    reason_rows.append("конфликты")
                if safe_numeric(pd.Series([row.get("outside_hours_ratio")])).iloc[0] > 0.3:
                    reason_rows.append("встречи вне времени")
        reasons = pd.Series(reason_rows).value_counts().reset_index()
        reasons.columns = ["причина", "случаев"] if not reasons.empty else ["причина", "случаев"]
        render_colored_bar_chart(reasons, "причина", "случаев", title="Причины риска", horizontal=True)

    with tabs[2]:
        render_section_title("✓", "Эффективность")
        render_leaderboard_dashboard(data, "Руководитель")
        left, right = st.columns(2)
        with left:
            if not score.empty:
                top_score = score.sort_values("general_score", ascending=False).head(15)
                render_colored_bar_chart(top_score, "full_name", "general_score", "department", "Топ по результативности", horizontal=True)
        with right:
            if not score.empty and "department" in score.columns:
                dept_score = score.groupby("department", as_index=False)["general_score"].mean().round(1)
                render_colored_bar_chart(dept_score, "department", "general_score", title="Средняя оценка по отделам", horizontal=True)
        if not score.empty:
            over = score[(safe_numeric(score["general_score"]) >= 60) & (safe_numeric(score["overtime_hours"]) > 2)].copy()
            render_color_table(over.head(20), columns=["full_name", "department", "general_score", "overtime_hours", "risk_score"], title="Высокая результативность + переработки", score_column="general_score", risk_column="risk_score")

    with tabs[3]:
        render_section_title("☑", "Проекты")
        if tasks.empty:
            st.info("Проектные задачи не найдены.")
        else:
            overdue = tasks[tasks["статус"] == "просрочено"].copy()
            left, right = st.columns(2)
            with left:
                overdue_dept = overdue.groupby("department", as_index=False).size().rename(columns={"size": "просрочено"})
                render_colored_bar_chart(overdue_dept, "department", "просрочено", title="Просрочки по отделам", horizontal=True)
            with right:
                deadline = tasks.sort_values("дней до дедлайна").head(15)
                render_colored_bar_chart(deadline, "full_name", "дней до дедлайна", "статус", "Ближайшие дедлайны", horizontal=True)
            help_needed = tasks[tasks["статус"].isin(["просрочено", "сегодня", "скоро дедлайн"])].sort_values("дней до дедлайна").head(20)
            table = help_needed.rename(columns={"full_name": "сотрудник", "task_name": "задача", "department": "отдел"})
            render_priority_table(table, ["сотрудник", "отдел", "задача", "дедлайн", "дней до дедлайна", "статус", "приоритет", "загрузка", "рекомендация"], "Кому нужна помощь")

    with tabs[4]:
        render_section_title("▦", "Доступность")
        render_bi_dashboards(data)
        events = prepare_calendar_events(data, "Руководитель", None)
        if not events.empty:
            event_counts = events.groupby("date", as_index=False).size().rename(columns={"size": "событий"})
            event_counts["дата"] = pd.to_datetime(event_counts["date"]).dt.strftime("%d.%m")
            busiest = event_counts.sort_values("событий", ascending=False).head(12)
            render_colored_bar_chart(busiest, "дата", "событий", title="Проблемные дни по событиям")
            quiet = event_counts.sort_values("событий").head(3)
            if not quiet.empty:
                windows = ", ".join(quiet["дата"].astype(str).tolist())
                st.info(f"Рекомендуемые окна для встреч: дни с меньшим числом событий — {windows}.")

    with tabs[5]:
        render_section_title("✦", "Рекомендации")
        if recommendations.empty:
            st.info("Рекомендации не найдены.")
        else:
            recs = normalize_columns(recommendations).rename(columns={
                "full_name": "сотрудник",
                "department": "отдел",
                "recommendation_text": "рекомендация",
                "priority": "приоритет",
                "status": "статус",
                "created_at": "дата",
            })
            st.dataframe(recs[[col for col in ["сотрудник", "отдел", "рекомендация", "приоритет", "статус", "дата"] if col in recs.columns]].head(80), use_container_width=True, hide_index=True)


# =====================================================================
# Stable multipage UI restore
# Основа: старая рабочая версия с GigaChat, BI, лидербордом и ролями.
# Дополнительно: страницы, начальная страница профиля, календарь отдельной
# страницей и переключатель тем по приложенному UI.
# =====================================================================

PAGE_OPTIONS = [
    "Обзор",
    "Профиль",
    "Календарь",
    "Аналитика",
    "Лидерборд",
    "Сотрудники",
    "Графики",
    "Задачи",
    "Риски и ошибки",
    "Рекомендации",
    "Настройки",
]

NAV_ICONS = {
    "Обзор": "◉",
    "Профиль": "⌂",
    "Календарь": "▦",
    "Аналитика": "▥",
    "Лидерборд": "★",
    "Сотрудники": "♙",
    "Графики": "◷",
    "Задачи": "☑",
    "Риски и ошибки": "⚠",
    "Рекомендации": "✦",
    "Настройки": "⚙",
}

ROLE_OPTIONS_STABLE = [
    "Сотрудник",
    "Администратор",
    "HR",
    "Проектный менеджер",
    "Аналитик",
    "Руководитель",
]

DARK_THEME = {
    "bg": "#090D14", "button_bg": "#08111C", "card": "#111827", "border": "#243044",
    "accent": "#7C3AED", "info": "#3882F6", "success": "#29B171", "warning": "#FACC15",
    "risk": "#E47419", "error": "#EF4444", "chart": "#A8E5FF", "text": "#F8FAFC", "muted": "#94A3B8",
}

LIGHT_THEME = {
    "bg": "#FFFFFF", "button_bg": "#D4DCED", "card": "#EBF2FF", "border": "#C4CDDD",
    "accent": "#7C3AED", "info": "#3882F6", "success": "#29B171", "warning": "#FACC15",
    "risk": "#E47419", "error": "#EF4444", "chart": "#7BD7FF", "text": "#111827", "muted": "#4B5563",
}


def strip_css_block(css, selector):
    start = css.find(selector)
    if start < 0:
        return css
    brace = css.find("{", start)
    if brace < 0:
        return css
    depth = 0
    for index in range(brace, len(css)):
        if css[index] == "{":
            depth += 1
        elif css[index] == "}":
            depth -= 1
            if depth == 0:
                return css[:start] + css[index + 1:]
    return css


def load_static_ui_css():
    css_path = Path("ui/styles.css")
    if not css_path.exists():
        return ""
    css = css_path.read_text(encoding="utf-8")
    css = strip_css_block(css, ":root")
    css = strip_css_block(css, ".theme-toggle-input:checked ~ .app-shell")
    return css


def css_vars(theme_name):
    p = DARK_THEME if theme_name == "Темная" else LIGHT_THEME
    static_css = load_static_ui_css()
    return f"""
    <style>
    :root {{
        --wts-bg: {p['bg']}; --wts-button-bg: {p['button_bg']}; --wts-card: {p['card']};
        --wts-border: {p['border']}; --wts-accent: {p['accent']}; --wts-info: {p['info']};
        --wts-success: {p['success']}; --wts-warning: {p['warning']}; --wts-risk: {p['risk']};
        --wts-error: {p['error']}; --wts-chart: {p['chart']}; --wts-text: {p['text']}; --wts-muted: {p['muted']};
        --bg: {p['bg']}; --button-bg: {p['button_bg']}; --card: {p['card']}; --border: {p['border']};
        --accent: {p['accent']}; --info: {p['info']}; --success: {p['success']}; --warning: {p['warning']};
        --risk: {p['risk']}; --error: {p['error']}; --chart: {p['chart']}; --text: {p['text']}; --muted: {p['muted']};
        --shadow: {"0 22px 55px rgba(0, 0, 0, 0.34)" if theme_name == "Темная" else "0 18px 50px rgba(56, 66, 90, 0.12)"};
        --radius: 8px;
    }}
    {static_css}
    .stApp {{ background: var(--wts-bg); color: var(--wts-text); font-family: Roboto, Arial, sans-serif; }}
    .block-container {{ padding-top: 1.4rem; padding-bottom: 2.4rem; max-width: 1680px; }}
    [data-testid="stSidebar"] {{ background: color-mix(in srgb, var(--wts-button-bg) 36%, var(--wts-bg)); border-right: 1px solid var(--wts-border); }}
    [data-testid="stSidebar"] * {{ color: var(--wts-text); }}
    [data-testid="stMetric"] {{ background: var(--wts-card); border: 1px solid var(--wts-border); border-radius: 14px; padding: 14px 16px; }}
    .wts-brand {{ display:flex; align-items:center; gap:12px; margin:8px 0 22px 0; font-size:18px; font-weight:900; color:var(--wts-accent); }}
    .wts-logo {{ width:30px; height:30px; border-radius:10px; background:linear-gradient(135deg, var(--wts-accent), #9F7AEA); display:flex; align-items:center; justify-content:center; color:white; font-weight:900; }}
    .wts-hero {{ background: var(--wts-card); border: 1px solid var(--wts-border); border-radius: 18px; padding: 24px; color: var(--wts-text); margin-bottom: 18px; }}
    .wts-hero-title {{ font-size: 32px; font-weight: 900; line-height:1.15; margin-bottom: 8px; }}
    .wts-hero-subtitle {{ color: var(--wts-muted); font-size: 14px; }}
    .wts-card-grid {{ display:grid; grid-template-columns: repeat(6, minmax(120px, 1fr)); gap:14px; margin-bottom:18px; }}
    .wts-card {{ background: var(--wts-card); border:1px solid var(--wts-border); border-radius:14px; padding:16px; min-height:100px; color:var(--wts-text); }}
    .wts-card-title {{ color: var(--wts-muted); font-size:13px; margin-bottom:8px; }}
    .wts-card-value {{ font-size:30px; font-weight:900; line-height:1; }}
    .wts-card-note {{ color: var(--wts-muted); font-size:12px; margin-top:7px; }}
    .wts-profile-card {{ background: var(--wts-card); border:1px solid var(--wts-border); border-radius:18px; padding:20px; min-height:260px; color:var(--wts-text); }}
    .wts-avatar {{ width:150px; height:150px; border-radius:12px; background:#D9D9D9; margin-bottom:18px; }}
    .wts-status {{ display:inline-block; padding:8px 18px; border-radius:8px; background:rgba(41,177,113,.15); color:var(--wts-success); border:1px solid rgba(41,177,113,.35); font-weight:800; }}
    .wts-update-row {{ display:grid; grid-template-columns: 80px 1fr 120px; gap:12px; align-items:center; padding:10px 0; border-bottom:1px solid var(--wts-border); color:var(--wts-text); }}
    .wts-badge {{ padding:5px 10px; border-radius:8px; font-size:12px; font-weight:800; text-align:center; }}
    .badge-ok {{ background:rgba(41,177,113,.16); color:var(--wts-success); }}
    .badge-err {{ background:rgba(239,68,68,.16); color:var(--wts-error); }}
    .badge-info {{ background:rgba(56,130,246,.16); color:var(--wts-info); }}
    .calendar-box {{ background:#F3F4F6; border:1px solid #E5E7EB; border-radius:18px; padding:22px; color:#111827; }}
    .event-card {{ background: var(--wts-card); border:1px solid var(--wts-border); border-radius:14px; padding:14px; color:var(--wts-text); margin-bottom:10px; }}
    .event-meta {{ color: var(--wts-muted); font-size:13px; }}
    div[data-testid="stButton"] > button {{ border-radius: 10px; font-weight:700; }}
    .streamlit-profile-grid {{ display:grid; grid-template-columns:minmax(0, 1fr) minmax(0, 2fr); gap:18px; margin-bottom:18px; }}
    .streamlit-profile-card {{ display:grid; grid-template-columns:150px minmax(0, 1fr); gap:24px; min-height:260px; }}
    .streamlit-profile-card .profile-photo {{ width:150px; height:150px; }}
    .streamlit-profile-card h2 {{ margin:10px 0 34px; font-size:28px; line-height:1.12; color:var(--text); }}
    .streamlit-profile-card p {{ margin:0 0 12px; color:var(--text); }}
    .streamlit-panel-title {{ display:flex; align-items:center; gap:12px; margin-bottom:20px; padding-bottom:16px; border-bottom:1px solid var(--border); }}
    .streamlit-panel-title h2 {{ margin:0; font-size:24px; color:var(--text); }}
    .streamlit-ai-panel {{ margin-bottom:18px; }}
    .overview-hero {{ display:grid; grid-template-columns:minmax(0, 1fr) 320px; gap:18px; align-items:stretch; margin-bottom:18px; }}
    .overview-hero-main {{ background:var(--card); border:1px solid var(--border); border-radius:var(--radius); padding:24px; box-shadow:var(--shadow); }}
    .overview-hero-main h1 {{ margin:0; font-size:32px; line-height:1.15; color:var(--text); }}
    .overview-hero-main p {{ margin:8px 0 0; color:var(--muted); font-size:14px; }}
    .overview-role-card {{ background:var(--card); border:1px solid var(--border); border-radius:var(--radius); padding:22px; box-shadow:var(--shadow); }}
    .overview-role-card h2 {{ margin:0 0 8px; color:var(--text); font-size:22px; }}
    .overview-role-card p {{ margin:0; color:var(--muted); font-size:13px; }}
    .overview-workspace-grid {{ display:grid; grid-template-columns:repeat(3, minmax(0, 1fr)); gap:18px; margin-top:18px; }}
    .overview-workspace-card {{ background:var(--card); border:1px solid var(--border); border-radius:var(--radius); padding:20px; box-shadow:var(--shadow); min-height:150px; }}
    .overview-workspace-card h3 {{ margin:0 0 10px; color:var(--text); font-size:18px; }}
    .overview-workspace-card p {{ margin:0; color:var(--muted); font-size:13px; line-height:1.45; }}
    .stTextArea textarea, .stSelectbox [data-baseweb="select"] {{ border-color:var(--border); background: color-mix(in srgb, var(--button-bg) 60%, var(--bg)); color:var(--text); }}
    @media (max-width: 1180px) {{
        .streamlit-profile-grid {{ grid-template-columns:1fr; }}
        .overview-hero, .overview-workspace-grid {{ grid-template-columns:1fr; }}
        .metrics-grid, .wts-card-grid {{ grid-template-columns: repeat(3, minmax(180px, 1fr)); }}
    }}
    @media (max-width: 760px) {{
        .streamlit-profile-card, .metrics-grid, .wts-card-grid {{ grid-template-columns:1fr; }}
        .update-row, .wts-update-row {{ grid-template-columns:54px minmax(0, 1fr); }}
        .update-row .badge, .wts-update-row .wts-badge {{ grid-column:2; justify-self:start; }}
    }}
    </style>
    """


def inject_theme(theme_name):
    st.markdown(css_vars(theme_name), unsafe_allow_html=True)


def render_nav_brand():
    st.sidebar.markdown(
        """
        <div class="wts-brand">
            <div class="wts-logo">◌</div>
            <div>WorkTime Sync</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_cards_html(items):
    html = '<section class="metrics-grid" aria-label="Основные показатели">'
    icons = ["♙", "♙", "☑", "!", "!", "◷", "✓", "△"]
    classes = ["info", "success", "warning", "risk", "error", "accent", "success", "risk"]
    for index, (title, value, note, color) in enumerate(items):
        icon = icons[index] if index < len(icons) else "•"
        icon_class = classes[index] if index < len(classes) else "info"
        html += (
            f'<div class="metric-card">'
            f'<span class="metric-icon {icon_class}">{escape_html(icon)}</span>'
            f'<div><p>{escape_html(title)}</p>'
            f'<strong style="color:{color};">{escape_html(value)}</strong>'
            f'<small>{escape_html(note)}</small></div>'
            f'</div>'
        )
    html += '</section>'
    st.markdown(html, unsafe_allow_html=True)


def get_employee_options_stable(employees):
    employees = normalize_id_columns(employees)
    if employees.empty or "employee_id" not in employees.columns:
        return []
    options = []
    for _, row in employees.iterrows():
        emp_id = str(row.get("employee_id"))
        name = row.get("full_name", emp_id)
        options.append((emp_id, f"{name} ({emp_id})"))
    return options


def get_selected_employee(data, key="selected_employee_main"):
    options = get_employee_options_stable(data["employees"])
    if not options:
        return None
    label_to_id = {label: emp_id for emp_id, label in options}
    label = st.sidebar.selectbox("Профиль", list(label_to_id.keys()), key=key)
    return label_to_id[label]


def get_current_employee_row(data, employee_id):
    employees = normalize_id_columns(data["employees"])
    if employees.empty or "employee_id" not in employees.columns:
        return pd.DataFrame()
    return employees[employees["employee_id"].astype(str) == str(employee_id)].copy()


def get_profile_metric(data, employee_id):
    metrics = normalize_id_columns(data["analytics_metrics"])
    if metrics.empty or "employee_id" not in metrics.columns:
        return pd.DataFrame()
    return metrics[metrics["employee_id"].astype(str) == str(employee_id)].copy()


def get_profile_score(data, employee_id):
    score = compute_general_score(data)
    if score.empty or "employee_id" not in score.columns:
        return pd.DataFrame()
    return score[score["employee_id"].astype(str) == str(employee_id)].copy()


def get_overview_stats(data):
    employees = normalize_id_columns(data["employees"])
    metrics = normalize_id_columns(data["analytics_metrics"])
    workload = normalize_id_columns(data["workload_logs"])
    score = compute_general_score(data)

    total_employees = len(employees)
    if not employees.empty and "status" in employees.columns:
        active_employees = int((employees["status"].astype(str).str.lower() == "active").sum())
    else:
        active_employees = total_employees

    tasks_in_work = len(workload)
    overdue_tasks = int(safe_sum(metrics, "conflict_count")) if "conflict_count" in metrics.columns else 0
    data_errors = int((safe_numeric(metrics["risk_score"]) >= 0.75).sum()) if "risk_score" in metrics.columns else 0
    avg_load = safe_mean(metrics, "load_ratio") if "load_ratio" in metrics.columns else safe_mean(metrics, "workload_score")
    avg_score = round(safe_numeric(score["general_score"]).mean(), 1) if not score.empty and "general_score" in score.columns else 0
    avg_risk = safe_mean(metrics, "risk_score")

    updated_at = None
    for df, col in [(metrics, "calculated_at"), (data["schedules"], "updated_at"), (workload, "log_date")]:
        if not df.empty and col in normalize_columns(df).columns:
            normalized = normalize_columns(df)
            values = parse_excel_datetime_series(normalized[col]).dropna()
            if not values.empty:
                candidate = values.max()
                if updated_at is None or candidate > updated_at:
                    updated_at = candidate

    if updated_at is None:
        updated_label = date.today().strftime("%d.%m.%Y")
    else:
        updated_label = updated_at.strftime("%d.%m.%Y, %H:%M")

    return {
        "total_employees": total_employees,
        "active_employees": active_employees,
        "tasks_in_work": tasks_in_work,
        "overdue_tasks": overdue_tasks,
        "data_errors": data_errors,
        "avg_load": avg_load,
        "avg_score": avg_score,
        "avg_risk": avg_risk,
        "updated_label": updated_label,
    }


def infer_auto_role(data, employee_id):
    row_df = get_current_employee_row(data, employee_id)
    if row_df.empty:
        return "Сотрудник"
    row = row_df.iloc[0]
    text = " ".join([str(row.get("role", "")), str(row.get("department", "")), str(row.get("email", ""))]).lower()
    if "admin" in text or "админ" in text:
        return "Администратор"
    if "hr" in text or "кадр" in text:
        return "HR"
    if "project" in text or "manager" in text or "pm" in text or "проект" in text:
        return "Проектный менеджер"
    if "analyst" in text or "analytics" in text or "аналит" in text:
        return "Аналитик"
    if "lead" in text or "head" in text or "director" in text or "руковод" in text:
        return "Руководитель"
    return "Сотрудник"


def prepare_calendar_events(data, role="Сотрудник", employee_id=None):
    employees = normalize_id_columns(data["employees"])
    events = add_employee_names(data["calendar_events"], employees)
    absences = add_employee_names(data["absences"], employees)
    metrics = add_employee_names(data["analytics_metrics"], employees)
    rows = []

    if not events.empty:
        events = normalize_id_columns(events)
        if employee_id is not None and "employee_id" in events.columns:
            events = events[events["employee_id"].astype(str) == str(employee_id)]
        if "start_time" in events.columns:
            events["calendar_date"] = pd.to_datetime(events["start_time"], errors="coerce").dt.date
        else:
            events["calendar_date"] = pd.NaT
        for _, row in events.iterrows():
            start = pd.to_datetime(row.get("start_time"), errors="coerce")
            end = pd.to_datetime(row.get("end_time"), errors="coerce")
            person = row.get("full_name", row.get("employee_id", "—"))
            event_type = row.get("event_type", "событие")
            rows.append({
                "date": row.get("calendar_date"),
                "time": start.strftime("%H:%M") if not pd.isna(start) else "—",
                "end_time": end.strftime("%H:%M") if not pd.isna(end) else "—",
                "type": "Событие",
                "status": "event",
                "label": str(event_type) if role == "Сотрудник" else f"{person} — {event_type}",
                "employee": person,
                "comment": row.get("comment", "—"),
            })

    if not absences.empty:
        absences = normalize_id_columns(absences)
        if employee_id is not None and "employee_id" in absences.columns:
            absences = absences[absences["employee_id"].astype(str) == str(employee_id)]
        if "start_date" in absences.columns:
            absences["start_date"] = pd.to_datetime(absences["start_date"], errors="coerce").dt.date
        if "end_date" in absences.columns:
            absences["end_date"] = pd.to_datetime(absences["end_date"], errors="coerce").dt.date
        else:
            absences["end_date"] = absences.get("start_date")
        for _, row in absences.iterrows():
            start = row.get("start_date")
            end = row.get("end_date")
            if pd.isna(start):
                continue
            if pd.isna(end):
                end = start
            current = start
            while current <= end:
                person = row.get("full_name", row.get("employee_id", "—"))
                absence_type = row.get("absence_type", "отсутствие")
                rows.append({
                    "date": current,
                    "time": "весь день",
                    "end_time": "—",
                    "type": "Отсутствие",
                    "status": "absence",
                    "label": str(absence_type) if role == "Сотрудник" else f"{person} — {absence_type}",
                    "employee": person,
                    "comment": row.get("comment", "—"),
                })
                current += timedelta(days=1)

    if role in ["HR", "Руководитель", "Аналитик"] and not metrics.empty and "days_since_update" in metrics.columns:
        risky = metrics[safe_numeric(metrics["days_since_update"]) > 60].copy()
        if "calculated_at" in risky.columns:
            risky["calendar_date"] = parse_excel_datetime_series(risky["calculated_at"]).dt.date
        else:
            risky["calendar_date"] = date.today()
        for _, row in risky.head(80).iterrows():
            person = row.get("full_name", row.get("employee_id", "—"))
            rows.append({
                "date": row.get("calendar_date"),
                "time": "—",
                "end_time": "—",
                "type": "Риск",
                "status": "risk",
                "label": f"{person} — обновить график",
                "employee": person,
                "comment": f"График не обновлялся {row.get('days_since_update', '—')} дней",
            })

    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["date", "time", "end_time", "type", "status", "label", "employee", "comment"])
    return df.dropna(subset=["date"])


def event_icon(status):
    if status == "absence":
        return "🟠"
    if status == "risk":
        return "🔴"
    return "🔵"


def render_calendar_stable(data, role="Сотрудник", employee_id=None, key_prefix="cal_stable"):
    events = prepare_calendar_events(data, role, employee_id)
    if f"{key_prefix}_selected" not in st.session_state:
        if not events.empty and len(events["date"].dropna()) > 0:
            st.session_state[f"{key_prefix}_selected"] = sorted(events["date"].dropna().unique())[0]
        else:
            st.session_state[f"{key_prefix}_selected"] = date.today()

    selected = st.session_state[f"{key_prefix}_selected"]
    month_names = {1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель", 5: "Май", 6: "Июнь", 7: "Июль", 8: "Август", 9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"}
    weekdays_ru = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]

    left, right = st.columns([1.2, 1])
    with left:
        chosen = st.date_input("Дата", value=selected, key=f"{key_prefix}_date")
        c1, c2 = st.columns([2, 1])
        with c1:
            month = st.selectbox("Месяц", list(month_names.keys()), index=chosen.month-1, format_func=lambda x: month_names[x], key=f"{key_prefix}_month")
        with c2:
            min_year = min(2020, int(chosen.year), int(selected.year))
            max_year = max(2035, int(chosen.year), int(selected.year))
            year = st.number_input("Год", min_value=min_year, max_value=max_year, value=int(chosen.year), step=1, key=f"{key_prefix}_year")
        if chosen != selected:
            selected = chosen
            st.session_state[f"{key_prefix}_selected"] = chosen
        if selected.month != month or selected.year != int(year):
            selected = date(int(year), int(month), 1)
            st.session_state[f"{key_prefix}_selected"] = selected

        st.markdown(f"""
        <div class="calendar-box">
            <div style="font-size:24px;font-weight:500;margin-bottom:18px;">{weekdays_ru[selected.weekday()]}, {selected.day:02d}.{selected.month:02d}.{selected.year}</div>
            <div style="background:#E5E7EB;border-radius:10px;padding:12px 14px;font-size:26px;font-weight:800;margin-bottom:16px;">{month_names[int(month)]} {int(year)}</div>
        </div>
        """, unsafe_allow_html=True)
        week_cols = st.columns(7)
        for i, w in enumerate(["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]):
            week_cols[i].markdown(f"<div style='text-align:center;font-weight:900;color:#111827;background:#F8FAFC;border-radius:10px;padding:8px;'>{w}</div>", unsafe_allow_html=True)
        cal = calendar.Calendar(firstweekday=0).monthdatescalendar(int(year), int(month))
        counts = events.groupby("date").size().to_dict() if not events.empty else {}
        statuses = events.groupby("date")["status"].apply(list).to_dict() if not events.empty else {}
        for wi, week in enumerate(cal):
            cols = st.columns(7)
            for di, day in enumerate(week):
                in_month = day.month == int(month)
                count = counts.get(day, 0)
                icons = "".join(event_icon(x) for x in statuses.get(day, [])[:3])
                label = f"{day.day}"
                if count:
                    label = f"{day.day}\n{icons} {count}"
                if not in_month:
                    label = f"({day.day})"
                with cols[di]:
                    if st.button(label, key=f"{key_prefix}_{year}_{month}_{wi}_{di}_{day.isoformat()}", type="primary" if day == st.session_state[f"{key_prefix}_selected"] else "secondary", disabled=not in_month, use_container_width=True):
                        st.session_state[f"{key_prefix}_selected"] = day
                        st.rerun()
        st.caption("🔵 событие · 🟠 отсутствие · 🔴 риск / обновить график")

    with right:
        chosen_day = st.session_state[f"{key_prefix}_selected"]
        st.subheader(f"Мероприятия на {chosen_day.strftime('%d.%m.%Y')}")
        day_rows = events[events["date"] == chosen_day].copy() if not events.empty else pd.DataFrame()
        if day_rows.empty:
            st.info("На выбранный день мероприятий нет.")
        else:
            day_rows = day_rows.sort_values(["time", "type", "label"])
            for _, row in day_rows.iterrows():
                st.markdown(f"""
                <div class="event-card">
                    <div style="font-size:16px;font-weight:900;margin-bottom:6px;">{event_icon(row.get('status'))} {escape_html(row.get('label', 'Мероприятие'))}</div>
                    <div class="event-meta">Время: {escape_html(row.get('time', '—'))}–{escape_html(row.get('end_time', '—'))}</div>
                    <div class="event-meta">Тип: {escape_html(row.get('type', '—'))}</div>
                    <div class="event-meta">Комментарий: {escape_html(row.get('comment', '—'))}</div>
                </div>
                """, unsafe_allow_html=True)


def workspace_page_for_role(role):
    if role == "Сотрудник":
        return "Профиль"
    return "Аналитика"


def render_overview_page(data, selected_employee_id, role):
    stats = get_overview_stats(data)
    employees = normalize_id_columns(data["employees"])
    profile = get_current_employee_row(data, selected_employee_id)
    if not profile.empty:
        name = str(profile.iloc[0].get("full_name", selected_employee_id))
    else:
        name = str(selected_employee_id)

    first_name = name.split()[0] if name else "пользователь"

    st.markdown(
        f'<section class="overview-hero">'
        f'<div class="overview-hero-main"><h1>Добрый день, {escape_html(first_name)}!</h1>'
        f'<p>Обзор WorkTime Sync · последнее обновление: {escape_html(stats["updated_label"])}</p></div>'
        f'<div class="overview-role-card"><h2>Рабочая область</h2>'
        f'<p>{escape_html(build_role_description(role))}</p></div>'
        f'</section>',
        unsafe_allow_html=True,
    )

    metric_cards_html([
        ("Всего сотрудников", stats["total_employees"], "в базе", "var(--wts-info)"),
        ("Активные сотрудники", stats["active_employees"], "доступны", "var(--wts-success)"),
        ("Задачи в работе", stats["tasks_in_work"], "записей трудозатрат", "var(--wts-warning)"),
        ("Просроченные задачи", stats["overdue_tasks"], "конфликты и нарушения", "var(--wts-risk)"),
        ("Ошибки в данных", stats["data_errors"], "к проверке", "var(--wts-error)"),
        ("Средняя загрузка", stats["avg_load"], "по команде", "var(--wts-accent)"),
        ("Результативность", stats["avg_score"], "средний балл", "var(--wts-success)"),
        ("Средний риск", stats["avg_risk"], "risk_score", "var(--wts-risk)"),
    ])

    st.markdown('<div class="panel" style="margin-bottom:18px;">', unsafe_allow_html=True)
    st.markdown(
        '<div class="streamlit-panel-title"><span class="heading-icon accent">◉</span>'
        '<h2>Выбор роли</h2></div>',
        unsafe_allow_html=True,
    )
    selected_role = st.selectbox(
        "Роль пользователя",
        ROLE_OPTIONS_STABLE,
        index=ROLE_OPTIONS_STABLE.index(role) if role in ROLE_OPTIONS_STABLE else 0,
        key="overview_role",
    )
    if selected_role != st.session_state.get("active_role"):
        st.session_state["active_role"] = selected_role
        st.session_state["pending_role"] = selected_role
        role = selected_role

    target_page = workspace_page_for_role(role)
    if st.button(f"Перейти в рабочую область: {target_page}", type="primary", use_container_width=True, key="overview_enter_workspace"):
        st.session_state["pending_page"] = target_page
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="panel" style="margin-bottom:18px;">', unsafe_allow_html=True)
    render_section_title("✦", "GigaChat-ассистент", "Можно выбрать готовый вопрос или написать свой.")
    quick_questions = [
        "Что сейчас самое важное в команде?",
        "Кто требует внимания в первую очередь?",
        "Какие риски надо закрыть сегодня?",
        "Где есть проблемы с графиками?",
        "Какие действия выполнить в первую очередь?",
        "Что происходит с моим графиком?",
    ]
    selected_question = st.selectbox("Готовые вопросы", quick_questions, key="overview_ai_quick_question")
    if st.session_state.get("overview_ai_last_quick_question") != selected_question:
        st.session_state["overview_ai_custom_question"] = selected_question
        st.session_state["overview_ai_last_quick_question"] = selected_question
    custom_question = st.text_area(
        "Свой вопрос",
        height=110,
        key="overview_ai_custom_question",
    )
    if st.button("Спросить GigaChat", type="primary", use_container_width=True, key="overview_ai_ask"):
        question = custom_question.strip() or selected_question
        with st.spinner("Ассистент анализирует данные..."):
            answer, source = llm_answer(role, question, data, selected_employee_id)
        st.caption("Источник: GigaChat" if source == "gigachat" else "Источник: встроенная логика")
        st.info(answer)
    st.markdown("</div>", unsafe_allow_html=True)

    role_focus = {
        "Сотрудник": [
            ("Мой профиль", "Персональный график, календарь, рекомендации и AI-подсказки по выбранному сотруднику."),
            ("Мои события", "Встречи, отсутствия и риски обновления графика без доступа к лишним командным данным."),
            ("Рекомендации", "Практические действия по текущему графику и нагрузке."),
        ],
        "Администратор": [
            ("Контроль данных", "Сотрудники, графики, события и метрики для поиска ошибок синхронизации."),
            ("Ошибки и риски", "Высокий risk_score, конфликты и устаревшие графики."),
            ("Настройки", "Подключения, секреты GigaChat и параметры источников."),
        ],
        "HR": [
            ("Команда", "Сотрудники, отсутствия, графики и доступность по подразделениям."),
            ("Риски выгорания", "Загрузка, переработки и сотрудники с высоким риском."),
            ("Рекомендации", "Приоритетные HR-действия на основе текущих метрик."),
        ],
        "Проектный менеджер": [
            ("Задачи", "Трудозатраты, шаблоны задач и покрытие плановых часов."),
            ("Календарь", "Командные события, отсутствия и точки перегруза."),
            ("Лидерборд", "Производительность и вклад сотрудников."),
        ],
        "Аналитик": [
            ("BI-графики", "Дашборды, тепловые карты и workload-представления из Excel."),
            ("Метрики", "risk_score, load_ratio, task_coverage и другие показатели."),
            ("Лидерборд", "Сравнение сотрудников по итоговому баллу."),
        ],
        "Руководитель": [
            ("Сводка команды", "Ключевые показатели, риски и результативность."),
            ("Лидерборд", "Лучшие сотрудники и зоны управленческого внимания."),
            ("Рекомендации", "Конкретные действия для снижения рисков и перегрузки."),
        ],
    }
    cards = role_focus.get(role, role_focus["Сотрудник"])
    cards_html = "".join(
        f'<div class="overview-workspace-card"><h3>{escape_html(title)}</h3><p>{escape_html(text)}</p></div>'
        for title, text in cards
    )
    st.markdown(f'<section class="overview-workspace-grid">{cards_html}</section>', unsafe_allow_html=True)

    if not employees.empty:
        st.caption(f"Текущий профиль: {name} · роль: {role}")


def render_start_profile_page(data, selected_employee_id, role):
    employees = normalize_id_columns(data["employees"])
    metrics = normalize_id_columns(data["analytics_metrics"])
    score = compute_general_score(data)
    profile = get_current_employee_row(data, selected_employee_id)
    my_score = score[score["employee_id"].astype(str) == str(selected_employee_id)].copy() if not score.empty and "employee_id" in score.columns else pd.DataFrame()
    personal_metrics = filter_employee_rows(data["analytics_metrics"], selected_employee_id)
    personal_workload = prepare_employee_workload(data, selected_employee_id)
    personal_recommendations = filter_employee_rows(data["recommendations"], selected_employee_id)

    name = "Пользователь"; dep = "—"; position = "—"; status = "Активен"
    if not profile.empty:
        row = profile.iloc[0]
        name = row.get("full_name", selected_employee_id)
        dep = row.get("department", "—")
        position = row.get("role", "—")
        status = row.get("status", "Активен")

    if not my_score.empty:
        score_row = my_score.iloc[0]
        personal_cards = [
            ("Моя оценка", f"{score_row.get('general_score', 0)} балла", "моя результативность", "var(--wts-success)"),
            ("Выполненные задачи", score_row.get("tasks_done", 0), "за период", "var(--wts-info)"),
            ("Трудозатраты", f"{score_row.get('project_hours', 0)} ч", "человеко-часы", "var(--wts-chart)"),
            ("Переработки", f"{score_row.get('overtime_hours', 0)} ч", "сверх плана", "var(--wts-risk)"),
            ("Риск", risk_level_label(score_row.get("risk_score", 0)), f"{score_row.get('risk_score', 0)} risk", "var(--wts-error)"),
            ("Актуальность графика", f"{score_row.get('days_since_update', 0)} дн.", "с последнего обновления", "var(--wts-warning)"),
        ]
    else:
        personal_cards = [
            ("Моя оценка", "—", "нет данных", "var(--wts-success)"),
            ("Выполненные задачи", len(personal_workload), "записей", "var(--wts-info)"),
            ("Трудозатраты", f"{safe_sum(personal_workload, 'часы')} ч", "человеко-часы", "var(--wts-chart)"),
            ("Переработки", "—", "нет данных", "var(--wts-risk)"),
            ("Риск", "—", "нет данных", "var(--wts-error)"),
            ("Актуальность графика", "—", "нет данных", "var(--wts-warning)"),
        ]

    st.markdown(f"""
    <div class="wts-hero">
        <div class="wts-hero-title">Добрый день, {escape_html(str(name).split()[0])}! 👋</div>
        <div class="wts-hero-subtitle">Профиль и обзор системы · выбранная роль: {escape_html(role)}</div>
    </div>
    """, unsafe_allow_html=True)
    metric_cards_html(personal_cards)

    updates = [
        ("10:30", "Данные синхронизированы из системы учёта", "Успешно", "success"),
        ("10:25", "Найдено 7 ошибок в данных", "Ошибка", "error"),
        ("10:20", "Обновлены графики трудозатрат", "Успешно", "success"),
        ("10:15", "Добавлен новый сотрудник", "Информация", "info"),
    ]
    update_rows = ""
    for t, txt, badge, kind in updates:
        update_rows += (
            f'<div class="update-row"><time>{escape_html(t)}</time>'
            f'<span>{escape_html(txt)}</span>'
            f'<span class="badge badge-{kind}">{escape_html(badge)}</span></div>'
        )

    st.markdown(
        f'<section class="streamlit-profile-grid">'
        f'<div class="panel streamlit-profile-card">'
        f'<div class="profile-photo" aria-hidden="true"></div>'
        f'<div class="profile-content"><h2>{escape_html(name)}</h2>'
        f'<p>{escape_html(position)}</p><p>{escape_html(dep)}</p>'
        f'<p>Статус: {escape_html(status)}</p><span class="badge badge-success">Активен</span></div>'
        f'</div><div class="panel"><div class="streamlit-panel-title">'
        f'<span class="heading-icon accent">↻</span><h2>Последние обновления</h2>'
        f'</div>{update_rows}</div></section>',
        unsafe_allow_html=True,
    )

    st.divider()
    render_section_title("☑", "Мои задачи и трудозатраты", "Только важные столбцы без технических названий.")
    if personal_workload.empty:
        st.info("Трудозатраты и задачи не найдены.")
    else:
        left_chart, right_chart = st.columns(2)
        with left_chart:
            by_type = personal_workload.groupby("тип задачи", as_index=False)["часы"].sum().sort_values("часы", ascending=False)
            render_colored_bar_chart(by_type, "тип задачи", "часы", title="Трудозатраты по типам задач", horizontal=True)
        with right_chart:
            by_priority = personal_workload.groupby(["приоритет", "_priority_key"], as_index=False).size().rename(columns={"size": "задачи"})
            render_colored_bar_chart(by_priority, "приоритет", "задачи", "приоритет", "Задачи по приоритетам", color_scale=PRIORITY_LABEL_COLORS)

        visible_workload = personal_workload.head(30).copy()
        render_priority_table(
            visible_workload,
            ["дата", "задача", "тип задачи", "часы", "приоритет", "источник"],
            "Таблица задач и трудозатрат",
        )

    if not personal_recommendations.empty:
        st.divider()
        render_section_title("✦", "Мои рекомендации")
        rec_view = normalize_columns(personal_recommendations).copy()
        rec_view = rec_view.rename(columns={"recommendation_text": "рекомендация", "priority": "приоритет", "status": "статус", "created_at": "дата"})
        st.dataframe(rec_view[[col for col in ["дата", "рекомендация", "приоритет", "статус"] if col in rec_view.columns]], use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("""
    <div class="streamlit-panel-title">
        <span class="heading-icon accent">✦</span>
        <h2>AI-ассистент</h2>
    </div>
    """, unsafe_allow_html=True)
    q = st.text_area("Ваш вопрос", value="Что происходит с моим графиком?", height=120, key="profile_ai_q")
    if st.button("Получить рекомендацию", use_container_width=True, key="profile_ai_btn"):
        with st.spinner("Ассистент анализирует данные..."):
            answer, source = llm_answer(role, q, data, selected_employee_id)
        st.caption("Источник: GigaChat" if source == "gigachat" else "Источник: встроенная логика")
        st.info(answer)


def render_calendar_page(data, selected_employee_id, role):
    st.header("Календарь")
    mode = st.radio("Режим", ["Мой календарь", "Командный календарь"], horizontal=True, key="calendar_mode")
    if mode == "Мой календарь":
        render_calendar_stable(data, "Сотрудник", selected_employee_id, key_prefix="calendar_page_personal")
    else:
        render_calendar_stable(data, role, None, key_prefix="calendar_page_team")


def render_people_page(data):
    st.header("Сотрудники")
    df = normalize_id_columns(data["employees"])
    if df.empty:
        st.info("Данных нет.")
        return
    if "department" in df.columns:
        dep = st.selectbox("Отдел", ["Все отделы"] + sorted(df["department"].dropna().astype(str).unique().tolist()), key="people_dep")
        if dep != "Все отделы":
            df = df[df["department"].astype(str) == dep]
    st.dataframe(df, use_container_width=True)


def render_schedules_page(data):
    st.header("Графики")
    st.dataframe(add_employee_names(data["schedules"], data["employees"]), use_container_width=True)


def render_tasks_page(data):
    st.header("Задачи")
    df = add_employee_names(data["workload_logs"], data["employees"])
    if not df.empty and "spent_hours" in df.columns and "task_template" in df.columns:
        df["spent_hours"] = safe_numeric(df["spent_hours"])
        group = df.groupby("task_template", as_index=False)["spent_hours"].sum().sort_values("spent_hours", ascending=False)
        st.bar_chart(group.set_index("task_template")["spent_hours"])
    st.dataframe(df, use_container_width=True)


def render_risks_errors_page(data):
    st.header("Риски и ошибки")
    metrics = add_employee_names(data["analytics_metrics"], data["employees"])
    if not metrics.empty and "risk_score" in metrics.columns:
        metrics["risk_score"] = safe_numeric(metrics["risk_score"])
        top = metrics.sort_values("risk_score", ascending=False).head(30)
        st.bar_chart(top.set_index("full_name")["risk_score"])
        render_color_table(top.head(20), columns=["full_name", "department", "risk_score", "risk_status", "behavior_segment", "days_since_update"], title="Сотрудники с высоким риском", risk_column="risk_score")


def render_recommendations_page(data, selected_employee_id, role):
    st.header("Рекомендации")
    left, right = st.columns([1.15, 1])
    with left:
        recs = add_employee_names(data["recommendations"], data["employees"])
        if recs.empty:
            st.info("Рекомендации не найдены.")
        else:
            st.dataframe(recs, use_container_width=True)
    with right:
        st.subheader("AI-ассистент")
        q = st.text_area("Вопрос", value="Какие действия выполнить в первую очередь?", height=120, key="recommendations_ai_q")
        if st.button("Получить рекомендацию", use_container_width=True, key="recommendations_ai_btn"):
            with st.spinner("Ассистент анализирует данные..."):
                answer, source = llm_answer(role, q, data, selected_employee_id)
            st.caption("Источник: GigaChat" if source == "gigachat" else "Источник: встроенная логика")
            st.info(answer)


def render_settings_page():
    st.header("Настройки")
    st.write("Подключение источников, права доступа, параметры GigaChat и обновления данных.")
    if _gigachat_configured():
        st.success("GigaChat подключён")
    else:
        st.info("GigaChat не настроен. Используется fallback-логика.")
    st.code('GIGACHAT_CREDENTIALS = "<base64>"\nGIGACHAT_SCOPE = "GIGACHAT_API_PERS"\nGIGACHAT_MODEL = "GigaChat"\nGIGACHAT_VERIFY_SSL = false', language="toml")


def main_stable():
    theme = st.sidebar.radio("Тема", ["Темная", "Светлая"], index=0, horizontal=True)
    inject_theme(theme)
    render_nav_brand()

    uploaded_file = st.sidebar.file_uploader("Импорт данных", type=["xlsx"])
    if uploaded_file is not None:
        file_source = uploaded_file
        st.sidebar.success("Файл загружен")
    elif Path(DATA_FILE).exists():
        file_source = DATA_FILE
        st.sidebar.info(f"Файл: {DATA_FILE}")
    else:
        file_source = None
        st.sidebar.warning("Загрузите Excel-файл")
    if file_source is None:
        st.stop()

    data = load_data(file_source)
    selected_employee_id = get_selected_employee(data, "main_employee_select")
    if selected_employee_id is None:
        st.stop()

    auto_role = infer_auto_role(data, selected_employee_id)
    if "active_role" not in st.session_state:
        st.session_state["active_role"] = auto_role if auto_role in ROLE_OPTIONS_STABLE else "Сотрудник"
    if "sidebar_role" not in st.session_state:
        st.session_state["sidebar_role"] = st.session_state["active_role"]
    if "active_page" not in st.session_state or st.session_state["active_page"] not in PAGE_OPTIONS:
        st.session_state["active_page"] = "Обзор"
    if st.session_state.get("pending_role") in ROLE_OPTIONS_STABLE:
        st.session_state["active_role"] = st.session_state["pending_role"]
        st.session_state["sidebar_role"] = st.session_state["pending_role"]
        del st.session_state["pending_role"]
    if st.session_state.get("pending_page") in PAGE_OPTIONS:
        st.session_state["active_page"] = st.session_state["pending_page"]
        del st.session_state["pending_page"]

    role = st.sidebar.selectbox(
        "Роль",
        ROLE_OPTIONS_STABLE,
        index=ROLE_OPTIONS_STABLE.index(st.session_state["sidebar_role"]) if st.session_state["sidebar_role"] in ROLE_OPTIONS_STABLE else 0,
        key="sidebar_role",
    )
    st.session_state["active_role"] = role
    page = st.sidebar.radio("Страницы", PAGE_OPTIONS, format_func=lambda x: f"{NAV_ICONS.get(x, '')}  {x}", key="active_page")

    st.sidebar.divider()
    if _gigachat_configured():
        st.sidebar.success("GigaChat подключён")
    else:
        st.sidebar.info("AI: встроенная логика")

    if page == "Обзор":
        render_overview_page(data, selected_employee_id, role)
    elif page == "Профиль":
        render_start_profile_page(data, selected_employee_id, role)
    elif page == "Календарь":
        render_calendar_page(data, selected_employee_id, role)
    elif page == "Аналитика":
        if role == "Администратор":
            render_admin_view(data)
        elif role == "HR":
            render_hr_view(data)
        elif role == "Проектный менеджер":
            render_pm_view(data)
        elif role == "Аналитик":
            render_analyst_view(data)
        elif role == "Руководитель":
            render_manager_view(data)
        else:
            render_bi_dashboards(data)
    elif page == "Лидерборд":
        render_leaderboard_dashboard(data, role)
    elif page == "Сотрудники":
        render_people_page(data)
    elif page == "Графики":
        render_schedules_page(data)
    elif page == "Задачи":
        render_tasks_page(data)
    elif page == "Риски и ошибки":
        render_risks_errors_page(data)
    elif page == "Рекомендации":
        render_recommendations_page(data, selected_employee_id, role)
    elif page == "Настройки":
        render_settings_page()


if __name__ == "__main__":
    main_stable()
