import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from datetime import datetime
from collections import defaultdict
import tempfile

st.set_page_config(layout="wide")
st.title("シフト表です。20250807の予定をもとに作りました。")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file:
    # Excelファイルの読み込み
    xls = pd.ExcelFile(uploaded_file)
    task_df = pd.read_excel(xls, sheet_name="task_schedule")
    personal_df = pd.read_excel(xls, sheet_name="personal_schedule")
    staff_df = pd.read_excel(xls, sheet_name="staff")

    # DataFrameの整形
    task_df = task_df.rename(columns={task_df.columns[0]: "date"})
    task_df['date'] = pd.to_datetime(task_df['date'])
    personal_df = personal_df.rename(columns={personal_df.columns[0]: "date"})
    personal_df['date'] = pd.to_datetime(personal_df['date'])
    staff_df = staff_df.fillna('')

    task_df = task_df.set_index("date")
    personal_df = personal_df.set_index("date")

    tasks = task_df.columns.tolist()
    task_dates = task_df.index.tolist()
    workers = personal_df.columns.tolist()
    worker_availability = personal_df.applymap(lambda x: False if str(x).lower() == 'x' else True)
    worker_skills = {
        row[0]: {
            staff_df.columns[i]: False if str(cell).lower() == 'x' else True
            for i, cell in enumerate(row[1:], start=1)
        } for row in staff_df.itertuples(index=False)
    }

    # タスク→日付→作業 の構造を作成
    task_plan = {}
    task_start_dates = {}
    for task in tasks:
        task_plan[task] = []
        for date in task_df.index:
            work = task_df.at[date, task]
            if pd.notna(work):
                task_plan[task].append((date, work))
        if task_plan[task]:
            task_start_dates[task] = min([d for d, _ in task_plan[task]])

    # OR-Tools モデル
    model = cp_model.CpModel()
    assignment = {}

    for task, schedule in task_plan.items():
        for date, work in schedule:
            for worker in workers:
                available = worker_availability.loc[date, worker]
                skilled = worker_skills.get(worker, {}).get(work, False)
                if available and skilled:
                    var = model.NewBoolVar(f"assign_{task}_{date}_{work}_{worker}")
                    assignment[(task, date, work, worker)] = var

    # 制約: 各作業に1人のみ
    for task, schedule in task_plan.items():
        for date, work in schedule:
            model.AddExactlyOne(
                assignment[(task, date, work, worker)]
                for worker in workers
                if (task, date, work, worker) in assignment
            )

    # 制約: 同じ人が同日に複数の作業をしない いや、する。5つまではやってもらう。k[0]はtask, k[1]はdate, k[2]はwork, k[3]はworker
    for date in task_dates:
        for worker in workers:
            model.Add(
                sum(
                    assignment[k]
                    for k in assignment
                    if k[1] == date and k[3] == worker
                ) <= 5
            )
    

    # 公平性: 各スタッフの作業数をカウントし、最大と最小の差を最小化
    task_counts = {
        worker: model.NewIntVar(0, len(assignment), f"count_{worker}") for worker in workers
    }
    for worker in workers:
        model.Add(task_counts[worker] == sum(
            var for (t, d, w, wr), var in assignment.items() if wr == worker
        ))

    max_tasks = model.NewIntVar(0, len(assignment), "max_tasks")
    min_tasks = model.NewIntVar(0, len(assignment), "min_tasks")
    model.AddMaxEquality(max_tasks, list(task_counts.values()))
    model.AddMinEquality(min_tasks, list(task_counts.values()))
    model.Minimize(max_tasks - min_tasks)

    # ソルバー
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        st.success("シフト割当完了")

        # 結果の集計
        shift_table = task_df.copy()
        personal_shifts = defaultdict(list)

        for (task, date, work, worker), var in assignment.items():
            if solver.Value(var):
                shift_table.at[date, task] = f"{work}+{worker}"
                shift_id = f"sample{task_start_dates[task].strftime('%Y/%m/%d')}_{work}"
                personal_shifts[worker].append((date, shift_id))

        # 表示: 全体シフト表
        renamed_cols = {
            task: f"sample{task_start_dates[task].strftime('%Y/%m/%d')}"
            for task in tasks
        }
        shift_table_display = shift_table.rename(columns=renamed_cols)
        shift_table_display.index = shift_table_display.index.map(lambda d: d.strftime('%Y-%m-%d（%a）'))
        st.subheader("全体")
        st.dataframe(shift_table_display)

        # 表示: スタッフ別シフト表
        st.subheader("個人のシフト")
        for worker in workers:
            records = sorted(personal_shifts[worker])
            if records:
                df = pd.DataFrame(records, columns=["日付", "作業"])
                df["日付"] = df["日付"].map(lambda d: d.strftime('%Y-%m-%d（%a）'))
                st.markdown(f"**{worker}**")
                st.dataframe(df)
    else:
        #追加
        st.subheader("原因候補：人員不足")
        st.warning("以下の日付で、割り当て可能なスタッフがいないタスクがあります。")

        found_unassignable = False
        for task, schedule in task_plan.items():
            for date, work in schedule:
                # 割り当て可能なスタッフのリストを作成
                possible_assignees = [
                    worker
                    for worker in workers
                    if worker_availability.loc[date, worker] and worker_skills.get(worker, {}).get(work, False)
                ]
                
                # もし、割り当て可能なスタッフが一人もいない場合
                if not possible_assignees:
                    st.write(f"日付: **{date.strftime('%Y-%m-%d')}**")
                    st.write(f"タスク: **{task}** (作業: **{work}**)")
                    st.write("このタスクをこなせるスタッフがいません。")
                    st.write("---")
                    found_unassignable = True
        
        if not found_unassignable:
            st.info("すべてのタスクに割り当て可能なスタッフは存在します。他の制約（公平性など）が厳しすぎる可能性があります。")