import streamlit as st
import pandas as pd
from datetime import date, datetime, time, timedelta
from itertools import combinations
from io import BytesIO
import random

st.set_page_config(page_title="巧固球賽程生成器", page_icon="🏆", layout="wide")

st.markdown("""
<style>
.main .block-container {max-width: 1180px; padding-top: 1.8rem;}
.hero {padding: 1.4rem 1.6rem; border-radius: 18px; background: linear-gradient(120deg,#102b46,#176b73); color: white; margin-bottom: 1rem;}
.hero h1 {color:white; margin:0; font-size:2rem;}
.hero p {color:#d8edf0; margin:.35rem 0 0 0;}
.small-note {color: #64748b; font-size:.9rem;}
</style>
<div class="hero">
<h1>🏆 巧固球賽程自動生成器</h1>
<p>TCHOUKBALL TOURNAMENT SCHEDULER · 自動分組、編排賽事及比較場地方案</p>
</div>
""", unsafe_allow_html=True)

st.caption("先設定賽制及隊伍，再安排日期與每日可用時段。生成後可下載 Excel 或 CSV。")

def team_labels(prefix, n):
    return [f"{prefix} {chr(65+i)}" if i < 26 else f"{prefix} {i+1}" for i in range(n)]

def round_robin(teams, label):
    games = []
    for a, b in combinations(teams, 2):
        games.append({"組別": label, "賽事階段": "單循環賽" if label in ("男子組", "女子組") else f"{label} 分組賽",
                      "主隊": a, "客隊": b})
    return games

def split_groups(teams, seed):
    shuffled = list(teams)
    rng = random.Random(seed)
    rng.shuffle(shuffled)
    a, b = [], []
    for i, t in enumerate(shuffled):
        (a if i % 2 == 0 else b).append(t)
    return a, b

def make_group_games(group, group_name, division):
    return [{"組別": division, "賽事階段": f"{group_name} 組分組賽", "主隊": a, "客隊": b}
            for a, b in combinations(group, 2)]

def make_playoffs(division, team_count, format_choice, groups):
    games = []
    if team_count <= 5:
        if format_choice == "純單循環（不設額外賽事）":
            return games
        if format_choice == "單循環＋四強＋冠軍賽＋季軍賽":
            games += [
                {"組別": division, "賽事階段": "四強", "主隊": "單循環第 1 名", "客隊": "單循環第 4 名"},
                {"組別": division, "賽事階段": "四強", "主隊": "單循環第 2 名", "客隊": "單循環第 3 名"},
                {"組別": division, "賽事階段": "季軍賽", "主隊": "四強負方 1", "客隊": "四強負方 2"},
                {"組別": division, "賽事階段": "冠軍賽", "主隊": "四強勝方 1", "客隊": "四強勝方 2"},
            ]
        elif format_choice == "單循環＋冠軍賽＋季軍賽":
            games += [
                {"組別": division, "賽事階段": "季軍賽", "主隊": "單循環第 3 名", "客隊": "單循環第 4 名"},
                {"組別": division, "賽事階段": "冠軍賽", "主隊": "單循環第 1 名", "客隊": "單循環第 2 名"},
            ]
        return games

    a, b = groups
    if team_count <= 7:
        games += [
            {"組別": division, "賽事階段": "四強", "主隊": "A 組第 1 名", "客隊": "B 組第 2 名"},
            {"組別": division, "賽事階段": "四強", "主隊": "B 組第 1 名", "客隊": "A 組第 2 名"},
        ]
    else:
        games += [
            {"組別": division, "賽事階段": "八強", "主隊": "A 組第 1 名", "客隊": "B 組第 4 名"},
            {"組別": division, "賽事階段": "八強", "主隊": "B 組第 2 名", "客隊": "A 組第 3 名"},
            {"組別": division, "賽事階段": "八強", "主隊": "B 組第 1 名", "客隊": "A 組第 4 名"},
            {"組別": division, "賽事階段": "八強", "主隊": "A 組第 2 名", "客隊": "B 組第 3 名"},
        ]
        games += [
            {"組別": division, "賽事階段": "四強", "主隊": "八強勝方 1", "客隊": "八強勝方 2"},
            {"組別": division, "賽事階段": "四強", "主隊": "八強勝方 3", "客隊": "八強勝方 4"},
        ]
    games += [
        {"組別": division, "賽事階段": "季軍賽", "主隊": "四強負方 1", "客隊": "四強負方 2"},
        {"組別": division, "賽事階段": "冠軍賽", "主隊": "四強勝方 1", "客隊": "四強勝方 2"},
    ]
    return games

def build_division(division, teams, format_choice, seed):
    games = []
    groups = None
    if len(teams) <= 5:
        games.extend(round_robin(teams, division))
        games.extend(make_playoffs(division, len(teams), format_choice, None))
    else:
        a, b = split_groups(teams, seed)
        groups = (a, b)
        games.extend(make_group_games(a, "A", division))
        games.extend(make_group_games(b, "B", division))
        games.extend(make_playoffs(division, len(teams), format_choice, groups))
    return games, groups

def parse_time(t):
    return datetime.combine(date.today(), t)

def schedule_matches(matches, days, day_start, day_end, courts, duration, end_mode=False):
    """Greedy scheduler. Keeps tournament stages in order and avoids same-team simultaneous games."""
    scheduled = []
    remaining = [dict(m) for m in matches]

    def phase(m):
        stage = m.get("賽事階段", "")
        if "分組賽" in stage or "單循環賽" in stage:
            return 0
        if stage == "八強":
            return 1
        if stage == "四強":
            return 2
        return 3

    phases = sorted(set(phase(m) for m in remaining))
    for current_phase in phases:
        phase_games = [m for m in remaining if phase(m) == current_phase]
        phase_remaining = list(phase_games)
        for d in days:
            start_dt = parse_time(day_start)
            end_dt = parse_time(day_end)
            cursor = end_dt if end_mode else start_dt
            while phase_remaining:
                if end_mode:
                    slot_end = cursor
                    slot_start = slot_end - timedelta(minutes=duration)
                    if slot_start < start_dt:
                        break
                else:
                    slot_start = cursor
                    slot_end = slot_start + timedelta(minutes=duration)
                    if slot_end > end_dt:
                        break
                selected = []
                used_teams = set()
                for idx, m in enumerate(phase_remaining):
                    teams_here = {m["主隊"], m["客隊"]}
                    if used_teams.isdisjoint(teams_here):
                        selected.append(idx)
                        used_teams.update(teams_here)
                        if len(selected) == courts:
                            break
                if not selected:
                    selected = [0]
                chosen = [phase_remaining[i] for i in selected]
                for court_num, m in enumerate(chosen, start=1):
                    scheduled.append({
                        "日期": d.strftime("%Y-%m-%d"),
                        "開始時間": slot_start.strftime("%H:%M"),
                        "結束時間": slot_end.strftime("%H:%M"),
                        "場地": f"Court {court_num}",
                        **m
                    })
                for idx in reversed(selected):
                    phase_remaining.pop(idx)
                cursor = slot_start if end_mode else slot_end
            if not phase_remaining:
                break
        remaining = [m for m in remaining if m not in phase_games] + phase_remaining
    return scheduled, remaining

def make_excel(sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(writer, sheet_name=name[:31], index=False)
    output.seek(0)
    return output.getvalue()

with st.sidebar:
    st.header("⚙️ 比賽基本設定")
    duration_choice = st.radio("每場比賽時間", ["15 分鐘 × 3 節（每場 60 分鐘）", "12 分鐘 × 3 節（每場 45 分鐘）"], index=0)
    duration = 60 if duration_choice.startswith("15") else 45
    st.divider()
    st.subheader("隊伍數量")
    men_count = st.number_input("男子組隊伍數", min_value=0, max_value=40, value=8, step=1)
    women_count = st.number_input("女子組隊伍數", min_value=0, max_value=40, value=6, step=1)
    custom_names = st.checkbox("自行輸入隊伍名稱")
    st.divider()
    court_mode = st.radio("生成哪種場地版本？", ["只生成單場地", "只生成雙場地", "同時生成兩個版本"], index=2)
    st.caption("雙場地會盡量安排不同隊伍同時比賽；同一隊不會被安排在同一時段的兩場賽事。")

st.header("1. 設定隊伍名稱")
division_teams = {}
for division, count, prefix in [("男子組", int(men_count), "Men"), ("女子組", int(women_count), "Women")]:
    with st.expander(f"{division}（{count} 隊）", expanded=(count > 0)):
        defaults = team_labels(prefix, count)
        if custom_names and count:
            entered = st.text_area(f"每行輸入一隊，{division}", value="\n".join(defaults), key=f"names_{division}", height=min(220, 38*count+20))
            names = [x.strip() for x in entered.splitlines() if x.strip()]
            if len(names) != count:
                st.warning(f"目前有 {len(names)} 個隊名，應該輸入 {count} 個。")
        else:
            names = defaults
            if names:
                st.write("、".join(names))
        division_teams[division] = names

st.header("2. 確認賽制")
formats = {}
for division, teams in division_teams.items():
    n = len(teams)
    with st.expander(f"{division}賽制 — {n} 隊", expanded=(n > 0)):
        if n <= 1:
            st.info("至少需要 2 隊才可以安排比賽。")
            formats[division] = "純單循環（不設額外賽事）"
        elif n <= 5:
            options = ["純單循環（不設額外賽事）"]
            if n >= 4:
                options += ["單循環＋四強＋冠軍賽＋季軍賽", "單循環＋冠軍賽＋季軍賽"]
            chosen = st.radio("單循環後的安排", options, key=f"format_{division}")
            formats[division] = chosen
        else:
            if n <= 7:
                st.success("兩組分組賽；每組頭兩名晉級四強，不設八強。")
            else:
                st.success("兩組分組賽；每組頭四名晉級八強。")
            formats[division] = "分組賽＋淘汰賽"

st.header("3. 日期及每日比賽時段")
c1, c2 = st.columns(2)
with c1:
    start_date = st.date_input("第一日比賽日期", value=date.today())
    num_days = st.number_input("比賽日數", min_value=1, max_value=14, value=2)
    start_time = st.time_input("每日開始時間", value=time(9, 0))
with c2:
    end_time = st.time_input("每日結束時間", value=time(18, 0))
    time_mode = st.radio("排程方向", ["由每日開始時間向後安排", "由每日結束時間向前倒推"])
    st.caption("每場比賽按所選的 45 或 60 分鐘連續安排，不額外加入休息時間。")

if start_time >= end_time:
    st.error("每日結束時間必須遲於開始時間。")
    st.stop()

days = [start_date + timedelta(days=i) for i in range(int(num_days))]
end_mode = time_mode.startswith("由每日結束")
if end_mode:
    st.info("倒推模式會由每日結束時間向前安排；未能放入指定時段的賽事會列為未排入。")

st.header("4. 生成賽程")
seed = st.number_input("分組隨機種子（相同數字會得到相同分組）", min_value=1, max_value=999999, value=54)
if st.button("🏆 生成完整賽程", type="primary", use_container_width=True):
    all_matches = []
    group_summaries = []
    errors = []
    for idx, (division, teams) in enumerate(division_teams.items()):
        if len(teams) != (men_count if division == "男子組" else women_count):
            errors.append(f"{division}隊伍名稱數量不正確。")
            continue
        if len(teams) >= 2:
            games, groups = build_division(division, teams, formats[division], int(seed)+idx)
            all_matches.extend(games)
            if groups:
                group_summaries.append((division, groups))
    if not all_matches:
        st.error("請至少輸入一個組別，而且該組別至少有 2 隊。")
    elif errors:
        st.error(" ".join(errors))
    else:
        st.session_state["all_matches"] = all_matches
        st.session_state["group_summaries"] = group_summaries
        st.session_state["days"] = days
        st.session_state["duration"] = duration
        st.session_state["start_time"] = start_time
        st.session_state["end_time"] = end_time
        st.session_state["end_mode"] = end_mode
        st.session_state["court_mode"] = court_mode
        st.success(f"已建立 {len(all_matches)} 場賽事的對賽清單。")

if "all_matches" in st.session_state:
    matches = st.session_state["all_matches"]
    st.divider()
    st.header("5. 賽程結果")
    for division, (a,b) in st.session_state["group_summaries"]:
        with st.expander(f"{division}分組結果"):
            st.write("**A 組：** " + "、".join(a))
            st.write("**B 組：** " + "、".join(b))
    st.metric("總賽事場數", len(matches))
    if st.session_state["court_mode"] == "只生成單場地":
        court_options = [("單場地", 1)]
    elif st.session_state["court_mode"] == "只生成雙場地":
        court_options = [("雙場地", 2)]
    else:
        court_options = [("單場地", 1), ("雙場地", 2)]

    sheets = {}
    for label, courts in court_options:
        scheduled, unscheduled = schedule_matches(
            matches, st.session_state["days"], st.session_state["start_time"],
            st.session_state["end_time"], courts, st.session_state["duration"],
            st.session_state["end_mode"]
        )
        st.subheader(f"{label}版本")
        if scheduled:
            df = pd.DataFrame(scheduled)
            st.dataframe(df, use_container_width=True, hide_index=True)
            sheets[label] = scheduled
            finish = max(datetime.strptime(x["結束時間"], "%H:%M").time() for x in scheduled)
            st.caption(f"已安排 {len(scheduled)} 場；每場 {st.session_state['duration']} 分鐘。")
        if unscheduled:
            st.warning(f"仍有 {len(unscheduled)} 場未能放入指定日期及時段。請增加比賽日數或延長每日可用時間。")
            with st.expander("查看未能安排的賽事"):
                st.dataframe(pd.DataFrame(unscheduled), use_container_width=True, hide_index=True)
        if scheduled:
            st.download_button(
                f"下載{label} CSV",
                data=pd.DataFrame(scheduled).to_csv(index=False).encode("utf-8-sig"),
                file_name=f"tchoukball_{courts}court_schedule.csv",
                mime="text/csv",
                key=f"csv_{courts}"
            )
    if sheets:
        st.download_button("📥 下載 Excel（包含已生成版本）", data=make_excel(sheets),
                           file_name="tchoukball_tournament_schedule.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    st.info("注意：淘汰賽對賽以排名佔位符表示（例如 A 組第 1 名）。完成分組賽並輸入排名／賽果後，才可確定實際隊名。此版本會先生成賽事輪次及時間，不會自動推算尚未產生的比賽結果。")
