import streamlit as st
import pandas as pd
import numpy as np
import calendar
from datetime import datetime
from io import BytesIO
import openpyxl
from ortools.sat.python import cp_model

# -------------------------------
# Excelファイルの読み込み
# -------------------------------
st.title("しふと")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])
if not uploaded_file:
    st.stop()

# シート読み込み
xls = pd.ExcelFile(uploaded_file)
#df_avail = pd.read_excel(xls, sheet_name="availability", index_col=0)
df_task = pd.read_excel(xls, sheet_name="tasks", index_col=0)

# 日付リスト取得
dates = list(df_task.index)

# スタッフ名とタスク名取得
staffs = list(df_avail.columns)
tasks = list(df_task.columns)

# -------------------------------
# シフト自動割り当て（簡略版）
# -------------------------------
assignments = {s: {d: "" for d in dates} for s in staffs}
work_count = {s: 0 for s in staffs}

for d in dates:
    for t in tasks:
        candidates = []
        for s in staffs:
            if df_avail.loc[d, s] != "x" and pd.notna(df_task.loc[d, t]):
                candidates.append(s)

        if candidates:
            # 最も割り当て回数が少ない人
            candidates.sort(key=lambda s: work_count[s])
            chosen = candidates[0]
            assignments[chosen][d] = t
            work_count[chosen] += 1

# -------------------------------
# 👤 スタッフ個別カレンダー表示
# -------------------------------
st.header("スタッフ個別のカレンダー表示")

staff_choice = st.selectbox("スタッフを選択してください", staffs)
year = st.selectbox("年を選んでください", sorted({d.year for d in dates}))
month = st.selectbox("月を選んでください", sorted({d.month for d in dates if d.year == year}))

# 指定月の該当日付とシフトを取得
shift_data = {d.day: assignments[staff_choice][d] for d in dates if d.year == year and d.month == month}

# カレンダーHTML生成
cal = calendar.Calendar(firstweekday=6)  # 日曜始まり

def generate_calendar_html(year, month, shifts):
    html = f"<h4>{year}年 {month}月</h4>"
    html += "<table border='1' style='width:100%; text-align:center; border-collapse:collapse;'>"
    html += "<tr>" + "".join(f"<th>{d}</th>" for d in ["日", "月", "火", "水", "木", "金", "土"]) + "</tr>"

    for week in cal.monthdayscalendar(year, month):
        html += "<tr>"
        for day in week:
            if day == 0:
                html += "<td></td>"
            else:
                shift = shifts.get(day, "")
                html += f"<td><strong>{day}</strong><br>{shift}</td>"
        html += "</tr>"
    html += "</table>"
    return html

cal_html = generate_calendar_html(year, month, shift_data)
st.markdown(cal_html, unsafe_allow_html=True)
