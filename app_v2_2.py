import streamlit as st
from streamlit_dnd import dnd, apply_move
import pandas as pd
from datetime import datetime, date, time, timedelta
from itertools import combinations
from io import BytesIO
from collections import defaultdict
import random
import math

st.set_page_config(page_title="巧固球賽程生成器 V2", page_icon="🏆", layout="wide")

st.markdown("""
<style>
.main .block-container {max-width: 1450px; padding-top: 1.4rem;}
h1 {font-weight: 800;}
div[data-testid="stDataFrame"] {width: 100%;}
</style>
""", unsafe_allow_html=True)

st.title("🏆 巧固球賽程生成器 V2")
st.caption("支援男女組別、單／雙球場、循環賽及淘汰賽；賽程可匯出 Excel 和 CSV。")
# ===== V2.2 拖曳編輯狀態 =====
if "edited_schedules" not in st.session_state:
    st.session_state.edited_schedules = {}

if "schedule_edit_version" not in st.session_state:
    st.session_state.schedule_edit_version = 0

def round_robin(teams, division, group_name="單組"):
    """Circle-method round robin. Returns one fixture list per round."""
    arr = list(teams)
    if len(arr) < 2:
        return []
    if len(arr) % 2:
        arr.append(None)
    n = len(arr)
    rounds = []
    rotating = arr[:]
    for r in range(n - 1):
        fixtures = []
        for i in range(n // 2):
            a, b = rotating[i], rotating[n - 1 - i]
            if a is not None and b is not None:
                fixtures.append({
                    "division": division, "stage": "循環賽", "group": group_name,
                    "home": a, "away": b, "phase": 0, "match_type": "real"
                })
        rounds.append(fixtures)
        rotating = [rotating[0]] + [rotating[-1]] + rotating[1:-1]
    return rounds

def add_knockout(matches, division, stage, group_name, home, away, phase):
    matches.append({
        "division": division, "stage": stage, "group": group_name,
        "home": home, "away": away, "phase": phase, "match_type": "placeholder"
    })

def build_division(teams, division, small_format):
    """Build group fixtures and placeholder knockout matches."""
    rounds = []
    if len(teams) < 2:
        return rounds

    if len(teams) <= 5:
        rr_rounds = round_robin(teams, division)
        rounds.extend(rr_rounds)
        if small_format == "純循環賽" or len(teams) < 4:
            return rounds
        if small_format == "循環賽＋決賽及季軍戰（不設準決賽）":
            add_knockout(rounds[-1] if False else [], division, "", "", "", "", 0) if False else None
            rounds.append([
                {"division": division, "stage": "決賽", "group": "淘汰賽",
                 "home": "循環賽第 1 名", "away": "循環賽第 2 名", "phase": 1, "match_type": "placeholder"},
                {"division": division, "stage": "季軍戰", "group": "淘汰賽",
                 "home": "循環賽第 3 名", "away": "循環賽第 4 名", "phase": 1, "match_type": "placeholder"},
            ])
        else:
            rounds.append([
                {"division": division, "stage": "準決賽", "group": "淘汰賽",
                 "home": "循環賽第 1 名", "away": "循環賽第 4 名", "phase": 1, "match_type": "placeholder"},
                {"division": division, "stage": "準決賽", "group": "淘汰賽",
                 "home": "循環賽第 2 名", "away": "循環賽第 3 名", "phase": 1, "match_type": "placeholder"},
            ])
            rounds.append([
                {"division": division, "stage": "決賽", "group": "淘汰賽",
                 "home": "準決賽勝方 1", "away": "準決賽勝方 2", "phase": 2, "match_type": "placeholder"},
                {"division": division, "stage": "季軍戰", "group": "淘汰賽",
                 "home": "準決賽負方 1", "away": "準決賽負方 2", "phase": 2, "match_type": "placeholder"},
            ])
        return rounds

    if len(teams) <= 7:
        rounds.extend(round_robin(teams, division, "單組"))
        rounds.append([
            {"division": division, "stage": "準決賽", "group": "淘汰賽",
             "home": "單組第 1 名", "away": "單組第 4 名", "phase": 1, "match_type": "placeholder"},
            {"division": division, "stage": "準決賽", "group": "淘汰賽",
             "home": "單組第 2 名", "away": "單組第 3 名", "phase": 1, "match_type": "placeholder"},
        ])
        rounds.append([
            {"division": division, "stage": "決賽", "group": "淘汰賽",
             "home": "準決賽勝方 1", "away": "準決賽勝方 2", "phase": 2, "match_type": "placeholder"},
            {"division": division, "stage": "季軍戰", "group": "淘汰賽",
             "home": "準決賽負方 1", "away": "準決賽負方 2", "phase": 2, "match_type": "placeholder"},
        ])
        return rounds

    # 8+ teams: two groups, top four in each group qualify for quarter-finals.
    shuffled = list(teams)
    random.shuffle(shuffled)
    group_a = shuffled[::2]
    group_b = shuffled[1::2]
    # Keep groups as even as possible and use stable names.
    for r in round_robin(group_a, division, "A 組"):
        rounds.append(r)
    for r in round_robin(group_b, division, "B 組"):
        rounds.append(r)
    rounds.append([
        {"division": division, "stage": "半準決賽", "group": "淘汰賽",
         "home": "A 組第 1 名", "away": "B 組第 4 名", "phase": 1, "match_type": "placeholder"},
        {"division": division, "stage": "半準決賽", "group": "淘汰賽",
         "home": "A 組第 2 名", "away": "B 組第 3 名", "phase": 1, "match_type": "placeholder"},
        {"division": division, "stage": "半準決賽", "group": "淘汰賽",
         "home": "B 組第 1 名", "away": "A 組第 4 名", "phase": 1, "match_type": "placeholder"},
        {"division": division, "stage": "半準決賽", "group": "淘汰賽",
         "home": "B 組第 2 名", "away": "A 組第 3 名", "phase": 1, "match_type": "placeholder"},
    ])
    rounds.append([
        {"division": division, "stage": "準決賽", "group": "淘汰賽",
         "home": "半準決賽勝方 1", "away": "半準決賽勝方 4", "phase": 2, "match_type": "placeholder"},
        {"division": division, "stage": "準決賽", "group": "淘汰賽",
         "home": "半準決賽勝方 2", "away": "半準決賽勝方 3", "phase": 2, "match_type": "placeholder"},
    ])
    rounds.append([
        {"division": division, "stage": "決賽", "group": "淘汰賽",
         "home": "準決賽勝方 1", "away": "準決賽勝方 2", "phase": 3, "match_type": "placeholder"},
        {"division": division, "stage": "季軍戰", "group": "淘汰賽",
         "home": "準決賽負方 1", "away": "準決賽負方 2", "phase": 3, "match_type": "placeholder"},
    ])
    return rounds

def flatten_rounds(rounds):
    result = []
    for round_no, fixtures in enumerate(rounds, 1):
        for fixture in fixtures:
            item = dict(fixture)
            item["round_no"] = round_no
            result.append(item)
    return result


def can_team_play(team, start_dt, duration_min, schedule, min_rest=60):
    """Minimum rest plus at least one completed intervening match."""
    if not team or team == "待定":
        return True

    prior = []
    for row in schedule:
        if row["主隊"] == team or row["客隊"] == team:
            prior.append(row)

    if not prior:
        return True

    last = max(prior, key=lambda x: x["_end"])
    last_end = last["_end"]

    # Minimum rest requirement
    if (start_dt - last_end).total_seconds() < min_rest * 60:
        return False

    # At least one completed match by other teams in between
    intervening = [
        row for row in schedule
        if row["_start"] >= last_end
        and row["_end"] <= start_dt
        and team not in (row["主隊"], row["客隊"])
    ]

    return len(intervening) >= 1


def schedule_fixtures(fixtures_by_division, start_date, day_windows,
                      duration_min, courts, rest_min=60, direction="正向",
                      blocked_intervals=None):
    """
    day_windows: list of (start_time, end_time), one tuple per tournament day.
    direction: 正向 (earliest first) or 逆向 (latest available slot first).
    """
    pending = {d: list(v) for d, v in fixtures_by_division.items()}
    schedule = []
    match_no = 1
    slot = timedelta(minutes=duration_min)
    blocked_intervals = blocked_intervals or []

    # Build valid match-slot candidates within each day's individual time window.
    slots = []
    for day_index, (day_start_time, day_end_time) in enumerate(day_windows):
        day_date = start_date + timedelta(days=day_index)
        dt = datetime.combine(day_date, day_start_time)
        day_end = datetime.combine(day_date, day_end_time)
        while dt + slot <= day_end:
            overlaps_event = any(
                dt < event_end and dt + slot > event_start
                for event_start, event_end in blocked_intervals
            )
            if not overlaps_event:
                slots.append(dt)
            dt += slot

    slots.sort(reverse=(direction == "逆向"))

    def can_team_play_at(team, start_dt):
        if not team or team == "待定":
            return True
        team_rows = [
            row for row in schedule
            if row.get("match_type") == "real"
            and (row["主隊"] == team or row["客隊"] == team)
        ]
        for row in team_rows:
            # Any overlap or less than minimum rest is forbidden.
            if start_dt >= row["_end"]:
                if (start_dt - row["_end"]).total_seconds() < rest_min * 60:
                    return False
                # At least one completed intervening match by other teams.
                between = [
                    x for x in schedule
                    if x.get("match_type") == "real"
                    and x["_start"] >= row["_end"]
                    and x["_end"] <= start_dt
                    and team not in (x["主隊"], x["客隊"])
                ]
                if not between:
                    return False
            elif start_dt + slot <= row["_start"]:
                if (row["_start"] - (start_dt + slot)).total_seconds() < rest_min * 60:
                    return False
                between = [
                    x for x in schedule
                    if x.get("match_type") == "real"
                    and x["_start"] >= start_dt + slot
                    and x["_end"] <= row["_start"]
                    and team not in (x["主隊"], x["客隊"])
                ]
                if not between:
                    return False
            else:
                return False
        return True

    last_division = None
    for current_dt in slots:
        if not any(pending.values()):
            break

        eligible = []
        for division, items in pending.items():
            if not items:
                continue
            phase = (max if direction == "逆向" else min)(x["phase"] for x in items)
            eligible.extend((division, x) for x in items if x["phase"] == phase)

        def div_priority(pair):
            div, _ = pair
            if last_division is None:
                return 0
            return 0 if div != last_division else 1

        eligible.sort(key=lambda pair: (div_priority(pair), pair[1]["round_no"]))
        chosen = []
        used_real_teams = set()

        for required_rest in (120, rest_min):
            for division, fixture in eligible:
                if len(chosen) >= courts:
                    break
                if any(existing is fixture for _, existing in chosen):
                    continue
                if fixture["match_type"] == "real":
                    teams = [fixture["home"], fixture["away"]]
                    if any(t in used_real_teams for t in teams):
                        continue
                    if not all(can_team_play_at(t, current_dt) for t in teams):
                        continue
                    # First pass prefers 120 minutes where feasible.
                    if required_rest == 120:
                        recent = [
                            row for row in schedule
                            if row.get("match_type") == "real"
                            and (row["主隊"] in teams or row["客隊"] in teams)
                        ]
                        if any(
                            current_dt >= row["_end"]
                            and (current_dt - row["_end"]).total_seconds() < 120 * 60
                            for row in recent
                        ):
                            continue
                else:
                    teams = []

                chosen.append((division, fixture))
                used_real_teams.update(teams)

            if len(chosen) >= courts:
                break

        for court_index, (division, fixture) in enumerate(chosen, 1):
            row = {
                "日期": current_dt.date().isoformat(),
                "星期": ["一", "二", "三", "四", "五", "六", "日"][current_dt.weekday()],
                "時間": f'{current_dt.strftime("%H:%M")}–{(current_dt + slot).strftime("%H:%M")}',
                "場次": match_no,
                "球場": f"球場 {court_index}",
                "組別": division,
                "賽事階段": fixture["stage"],
                "分組": fixture["group"],
                "主隊": fixture["home"],
                "主隊得分": "",
                "客隊得分": "",
                "客隊": fixture["away"],
                "_start": current_dt,
                "_end": current_dt + slot,
                "match_type": fixture["match_type"],
            }
            schedule.append(row)
            pending[division].remove(fixture)
            match_no += 1
            last_division = division

    unscheduled = sum(len(v) for v in pending.values())
    return schedule, unscheduled

def make_excel(df, team_rows, title_text, venue="", court_count=1):
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins

    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "賽程表"

    # 顏色
    blue = "D1E3F4"
    pink = "EAD0DF"
    black = "000000"
    white = "FFFFFF"
    gray = "A6A6A6"
    border_color = "000000"

    thin = Side(style="thin", color=border_color)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 每個球場表格的欄位：
    # 日期地點、時間、場次、主隊、主隊比分、客隊比分、客隊
    block_width = 7
    court_count = max(1, court_count)
    total_cols = block_width * court_count

    # 標題
    ws.merge_cells(
        start_row=1, start_column=1,
        end_row=1, end_column=total_cols
    )
    ws.cell(1, 1, title_text)
    ws.cell(1, 1).font = Font(
        name="Microsoft JhengHei", size=18,
        bold=True, color=white
    )
    ws.cell(1, 1).fill = PatternFill("solid", fgColor=black)
    ws.cell(1, 1).alignment = Alignment(
        horizontal="center", vertical="center"
    )
    ws.row_dimensions[1].height = 38

    # 各球場表格標題
    for court in range(court_count):
        start_col = court * block_width + 1
        end_col = start_col + block_width - 1

        ws.merge_cells(
            start_row=2, start_column=start_col,
            end_row=2, end_column=end_col
        )
        cell = ws.cell(2, start_col, f"球場 {court + 1}")
        cell.font = Font(
            name="Microsoft JhengHei", size=14, bold=True
        )
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill("solid", fgColor="D9E1F2")

    # 欄位標題
    headers = ["日期／地點", "時間", "場次", "主隊", "主隊比分", "客隊比分", "客隊"]

    header_row = 3
    for court in range(court_count):
        start_col = court * block_width + 1

        for i, header in enumerate(headers):
            cell = ws.cell(header_row, start_col + i, header)
            cell.font = Font(
                name="Microsoft JhengHei", bold=True, size=11
            )
            cell.fill = PatternFill("solid", fgColor="FFFFFF")
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.border = border

    # 日期／時間排序
    df = df.copy()
    df["_date"] = pd.to_datetime(df["日期"])
    df = df.sort_values(["_date", "_start", "球場"])

    # 每日每個時間段
    grouped_days = list(df.groupby("日期", sort=True))
    row_num = 4

    for day_index, (day, day_df) in enumerate(grouped_days, 1):
        day_df = day_df.sort_values("_start")
        times = sorted(day_df["_start"].unique())

        day_start_row = row_num

        for start_dt in times:
            slot_df = day_df[day_df["_start"] == start_dt]

            for court in range(court_count):
                court_col = court * block_width + 1

                court_label = f"球場 {court + 1}"
                match = slot_df[slot_df["球場"] == court_label]

                event_match = slot_df[slot_df.get("match_type", pd.Series(index=slot_df.index, dtype=object)) == "event"]
                if match.empty and court == 0 and not event_match.empty:
                    match = event_match.iloc[0]
                    values = [
                        "",
                        match["時間"].replace(":", ""),
                        "",
                        match["主隊"],
                        "",
                        "",
                        match["客隊"],
                    ]
                    division = "活動"
                elif match.empty:
                    values = ["", "", "", "", "", "", ""]
                    division = ""
                else:
                    match = match.iloc[0]
                    values = [
                        "",
                        match["時間"].replace(":", ""),
                        match["場次"],
                        match["主隊"],
                        "",
                        "",
                        match["客隊"],
                    ]
                    division = match["組別"]

                for col_offset, value in enumerate(values):
                    cell = ws.cell(
                        row_num, court_col + col_offset, value
                    )

                    fill_color = (
                        blue if division == "男子組"
                        else pink if division == "女子組"
                        else "D9D9D9" if division == "活動"
                        else white
                    )

                    cell.fill = PatternFill(
                        "solid", fgColor=fill_color
                    )
                    cell.font = Font(
                        name="Microsoft JhengHei", size=12
                    )
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True
                    )
                    cell.border = border

                # 主隊及客隊名稱使用較大格
                ws.cell(row_num, court_col + 3).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
                ws.cell(row_num, court_col + 6).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

            ws.row_dimensions[row_num].height = 30
            row_num += 1

        # 日期／地點合併儲存格
        day_end_row = row_num - 1
        date_text = pd.to_datetime(day).strftime("%d/%m/%Y")

        weekday = ["一", "二", "三", "四", "五", "六", "日"][
            pd.to_datetime(day).weekday()
        ]

        date_location = (
            f"DAY {day_index}\n"
            f"{date_text} ({weekday})\n"
            f"{venue}"
        )

        for court in range(court_count):
            date_col = court * block_width + 1

            ws.merge_cells(
                start_row=day_start_row,
                start_column=date_col,
                end_row=day_end_row,
                end_column=date_col
            )

            cell = ws.cell(day_start_row, date_col, date_location)
            cell.font = Font(
                name="Microsoft JhengHei", size=12
            )
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = border

        # 每日完結灰色分隔列
        ws.merge_cells(
            start_row=row_num,
            start_column=1,
            end_row=row_num,
            end_column=total_cols
        )
        end_cell = ws.cell(
            row_num, 1,
            f"第 {day_index} 日賽事 完結"
        )
        end_cell.fill = PatternFill("solid", fgColor=gray)
        end_cell.font = Font(
            name="Microsoft JhengHei", size=12, bold=True
        )
        end_cell.alignment = Alignment(horizontal="center")
        end_cell.border = border
        ws.row_dimensions[row_num].height = 25

        row_num += 1

    # 欄寬：隊伍名稱特別加寬
    widths = [20, 14, 10, 27, 10, 10, 27]

    for court in range(court_count):
        start_col = court * block_width + 1

        for i, width in enumerate(widths):
            ws.column_dimensions[
                get_column_letter(start_col + i)
            ].width = width

    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False

    # 列印設定
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins = PageMargins(
        left=0.25, right=0.25,
        top=0.5, bottom=0.5,
        header=0.2, footer=0.2
    )

    # 隊伍名單工作表
    team_ws = wb.create_sheet("隊伍名單")
    team_ws.append(["組別", "隊伍名稱"])

    for division, team in team_rows:
        team_ws.append([division, team])

    for cell in team_ws[1]:
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor="404F66")

    team_ws.column_dimensions["A"].width = 20
    team_ws.column_dimensions["B"].width = 35
    team_ws.freeze_panes = "A2"

    wb.save(output)
    output.seek(0)

    return output.getvalue()


with st.sidebar:
    st.header("⚙️ 賽事設定")
    tournament_name = st.text_input("賽事名稱", "巧固球錦標賽")
    venue = st.text_input("比賽場地", "")
    start_date = st.date_input("比賽首日", value=date.today())
    days = st.number_input("比賽日數", min_value=1, max_value=14, value=1, step=1)
    direction = st.radio("排程方向", ["正向（由開始時間排起）", "逆向（由結束時間倒推）"])
    direction_mode = "逆向" if direction.startswith("逆向") else "正向"

    st.markdown("### 每日比賽時間")
    day_windows = []
    for day_i in range(int(days)):
        with st.expander(f"第 {day_i + 1} 日時間", expanded=(day_i == 0)):
            day_start = st.time_input(
                f"第 {day_i + 1} 日開始時間", value=time(9, 0),
                key=f"day_start_{day_i}"
            )
            day_end = st.time_input(
                f"第 {day_i + 1} 日結束時間", value=time(18, 0),
                key=f"day_end_{day_i}"
            )
            day_windows.append((day_start, day_end))

    st.markdown("### 開幕禮／閉幕禮（選填）")
    opening_enabled = st.checkbox("加入開幕禮（比賽首日）", value=False)
    opening_time = st.time_input("開幕禮開始時間", value=time(9, 0), key="opening_time")
    opening_duration = st.number_input(
        "開幕禮時長（分鐘）", min_value=5, max_value=240, value=30, step=5,
        key="opening_duration"
    )
    closing_enabled = st.checkbox("加入閉幕禮（最後一日）", value=False)
    closing_time = st.time_input("閉幕禮開始時間", value=time(17, 30), key="closing_time")
    closing_duration = st.number_input(
        "閉幕禮時長（分鐘）", min_value=5, max_value=240, value=30, step=5,
        key="closing_duration"
    )

    duration_choice = st.radio("每場比賽時間（包括場次時段）", ["60 分鐘（15 分鐘 × 3 節）", "45 分鐘（12 分鐘 × 3 節）"])
    duration_min = 60 if duration_choice.startswith("60") else 45
    court_mode = st.radio("球場安排", ["單球場", "雙球場", "同時生成單、雙球場版本"])
    small_format = st.selectbox(
        "隊伍不多於 5 隊時的賽制",
        ["純循環賽", "循環賽＋準決賽、決賽及季軍戰", "循環賽＋決賽及季軍戰（不設準決賽）"]
    )

st.subheader("👥 輸入參賽隊伍")
left, right = st.columns(2)
with left:
    men_count = st.number_input("男子組隊數", min_value=0, max_value=20, value=4, step=1)
    men_text = st.text_area("男子組隊伍名稱（每行一隊；留空則使用預設名稱）",
                            value="\n".join([f"男子 {chr(65+i)}" for i in range(int(men_count))]), height=170)
with right:
    women_count = st.number_input("女子組隊數", min_value=0, max_value=20, value=4, step=1)
    women_text = st.text_area("女子組隊伍名稱（每行一隊；留空則使用預設名稱）",
                              value="\n".join([f"女子 {chr(65+i)}" for i in range(int(women_count))]), height=170)

def parse_teams(raw, count, prefix):
    names = [x.strip() for x in raw.splitlines() if x.strip()]
    if not names:
        names = [f"{prefix} {chr(65+i)}" for i in range(int(count))]
    return names[:int(count)] if len(names) >= int(count) else names + [f"{prefix} {chr(65+i)}" for i in range(len(names), int(count))]

men_teams = parse_teams(men_text, men_count, "男子")
women_teams = parse_teams(women_text, women_count, "女子")

st.info("休息規則：每隊兩場比賽之間最少休息 60 分鐘，並須有至少一場其他隊伍的完整比賽在中間完成；系統會盡量爭取更長休息。")

if st.button("🚀 生成賽程", type="primary", use_container_width=True):
    if int(men_count) < 2 and int(women_count) < 2:
        st.error("至少需要一個組別有 2 隊或以上，才能生成賽程。")
    elif any(end_t <= start_t for start_t, end_t in day_windows):
        st.error("每一日的結束時間都必須遲於開始時間。")
    elif opening_enabled and int(days) >= 1 and not (
        day_windows[0][0] <= opening_time
        and (datetime.combine(start_date, opening_time) + timedelta(minutes=int(opening_duration))).time() <= day_windows[0][1]
    ):
        st.error("開幕禮必須完整位於第一日的比賽時間範圍內。")
    elif closing_enabled and not (
        day_windows[-1][0] <= closing_time
        and (datetime.combine(start_date + timedelta(days=int(days)-1), closing_time) + timedelta(minutes=int(closing_duration))).time() <= day_windows[-1][1]
    ):
        st.error("閉幕禮必須完整位於最後一日的比賽時間範圍內。")
    else:
        divisions = {}
        team_rows = []
        if len(men_teams) >= 2:
            divisions["男子組"] = flatten_rounds(build_division(men_teams, "男子組", small_format))
            team_rows.extend([("男子組", t) for t in men_teams])
        if len(women_teams) >= 2:
            divisions["女子組"] = flatten_rounds(build_division(women_teams, "女子組", small_format))
            team_rows.extend([("女子組", t) for t in women_teams])

        court_versions = [1, 2] if court_mode == "同時生成單、雙球場版本" else ([1] if court_mode == "單球場" else [2])
        for court_count in court_versions:
            # Re-create fixture copies because scheduling mutates pending lists.
            fixture_copy = {d: [dict(x) for x in items] for d, items in divisions.items()}
            blocked_intervals = []
            event_rows = []

            if opening_enabled:
                event_start = datetime.combine(start_date, opening_time)
                event_end = event_start + timedelta(minutes=int(opening_duration))
                blocked_intervals.append((event_start, event_end))
                event_rows.append({
                    "日期": event_start.date().isoformat(),
                    "星期": ["一", "二", "三", "四", "五", "六", "日"][event_start.weekday()],
                    "時間": f"{event_start.strftime('%H:%M')}–{event_end.strftime('%H:%M')}",
                    "場次": "", "球場": "全場", "組別": "活動",
                    "賽事階段": "開幕禮", "分組": "",
                    "主隊": "開幕禮", "主隊得分": "", "客隊得分": "", "客隊": "",
                    "_start": event_start, "_end": event_end, "match_type": "event"
                })

            if closing_enabled:
                closing_date = start_date + timedelta(days=int(days) - 1)
                event_start = datetime.combine(closing_date, closing_time)
                event_end = event_start + timedelta(minutes=int(closing_duration))
                blocked_intervals.append((event_start, event_end))
                event_rows.append({
                    "日期": event_start.date().isoformat(),
                    "星期": ["一", "二", "三", "四", "五", "六", "日"][event_start.weekday()],
                    "時間": f"{event_start.strftime('%H:%M')}–{event_end.strftime('%H:%M')}",
                    "場次": "", "球場": "全場", "組別": "活動",
                    "賽事階段": "閉幕禮", "分組": "",
                    "主隊": "閉幕禮", "主隊得分": "", "客隊得分": "", "客隊": "",
                    "_start": event_start, "_end": event_end, "match_type": "event"
                })

            schedule, unscheduled = schedule_fixtures(
                fixture_copy, start_date, day_windows,
                duration_min, court_count, rest_min=60,
                direction=direction_mode, blocked_intervals=blocked_intervals
            )
            schedule.extend(event_rows)
            schedule.sort(key=lambda row: (row["_start"], row["球場"]))
            # Match numbers always follow the displayed chronological order.
            chronological_match_no = 1
            for scheduled_row in schedule:
                if scheduled_row.get("match_type") != "event":
                    scheduled_row["場次"] = chronological_match_no
                    chronological_match_no += 1
            if not schedule:
                st.warning(f"{court_count} 個球場版本未能排出賽事，請增加比賽日數或延長每日比賽時間。")
                continue

            # ===== 保存賽程，供拖曳編輯及匯出使用 =====
            slot_groups = {}

            for row in schedule:
                slot_key = row["_start"].isoformat()
                if slot_key not in slot_groups:
                    slot_groups[slot_key] = []
                slot_groups[slot_key].append(dict(row))

            slot_items = []

            for index, (slot_key, rows) in enumerate(slot_groups.items()):
                first_row = rows[0]

                match_labels = []
                for row in rows:
                    if row.get("match_type") == "event":
                        match_labels.append(row["賽事階段"])
                    else:
                        match_labels.append(
                            f'{row["球場"]}：{row["主隊"]} 對 {row["客隊"]}'
                        )

                    slot_items.append({
                    "id": f"{court_count}_{index}",
                    "start": slot_key,
                    "label": (
                        f'{first_row["日期"]} '
                        f'{first_row["星期"]} '
                        f'{first_row["時間"]}｜'
                        + "；".join(match_labels)
                    ),
                    "rows": rows,
                    "movable": all(row.get("match_type") != "event" for row in rows)
                })

            st.session_state.edited_schedules[court_count] = {
                "slots": slot_items,
                "unscheduled": unscheduled,
                "team_rows": list(team_rows),
                "tournament_name": tournament_name,
                "venue": venue,
                "court_count": court_count,
                "time_positions": sorted(
                    item["start"] for item in slot_items if item["movable"]
                )
            }
