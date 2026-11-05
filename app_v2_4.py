import streamlit as st
import pandas as pd
from datetime import datetime, date, time, timedelta
from io import BytesIO
from collections import defaultdict
import random

# 嘗試載入 streamlit_dnd，若環境未安裝則提供安全備援，避免整個 App 崩潰出唔到表
try:
    from streamlit_dnd import dnd, apply_move
    HAS_DND = True
except ImportError:
    HAS_DND = False

st.set_page_config(page_title="巧固球賽程生成器 V2.4", page_icon="🏆", layout="wide")
st.markdown("""<style>.main .block-container{max-width:1450px;padding-top:1.2rem} h1{font-weight:800}</style>""", unsafe_allow_html=True)
st.title("🏆 巧固球賽程生成器 V2.4")
st.caption("男女組別・單／雙球場・循環賽及淘汰賽・拖曳調動・錦標賽專屬 Excel 排版匯出")

# ------------------------- 隊伍與代號解析 -------------------------
def parse_teams_with_codes(raw, count, default_code_prefix):
    """
    支援三種情況：
    1. 留空不填 -> 自動產生代號與隊名，例如：男子組 MA, MB, MC... / 女子組 WA, WB, WC...
    2. 只輸入隊名（每行一隊），例如：HKUST -> 自動配對代號 MA, MB... / WA, WB...
    3. 輸入「代號 隊名」（以空格、Tab 或直線分隔），例如：UMA HKUST
    """
    lines = [x.strip() for x in raw.splitlines() if x.strip()]
    teams = []
    for i in range(int(count)):
        default_code = f"{default_code_prefix}{chr(65+i)}"
        if i < len(lines):
            line = lines[i]
            for sep in ['\t', '|', '｜']:
                line = line.replace(sep, ' ')
            parts = line.split(None, 1)
            if len(parts) == 2 and len(parts[0]) <= 6 and parts[0].isalnum():
                code, name = parts[0], parts[1].strip()
            else:
                code, name = default_code, line
        else:
            # 留空時直接使用 MA, MB, MC / WA, WB, WC
            code, name = default_code, default_code
        teams.append({'code': code, 'name': name, 'id': f"{code}_{name}"})
    return teams

# ------------------------- 賽制 -------------------------
def round_robin(teams, division, group_name="單組"):
    arr = list(teams)
    if len(arr) < 2:
        return []
    if len(arr) % 2:
        arr.append(None)
    out = []
    for rnd in range(len(arr) - 1):
        games = []
        for i in range(len(arr) // 2):
            a, b = arr[i], arr[-1 - i]
            if a is not None and b is not None:
                games.append(dict(
                    division=division,
                    stage="循環賽",
                    group=group_name,
                    home_code=a['code'],
                    home=a['name'],
                    away_code=b['code'],
                    away=b['name'],
                    phase=0,
                    match_type="real"
                ))
        out.append(games)
        arr = [arr[0], arr[-1]] + arr[1:-1]
    return out

def make_knockout_match(division, stage, home_label, away_label, phase):
    return dict(
        division=division,
        stage=stage,
        group="淘汰賽",
        home_code="",
        home=home_label,
        away_code="",
        away=away_label,
        phase=phase,
        match_type="placeholder"
    )

def build_division(teams, division, small_format):
    if len(teams) < 2:
        return []
    rounds = []
    if len(teams) <= 5:
        rounds = round_robin(teams, division)
        if small_format == "純循環賽" or len(teams) < 4:
            return rounds
        if small_format == "循環賽＋決賽及季軍戰（不設準決賽）":
            rounds.append([
                make_knockout_match(division, "季軍賽", "循環賽第 3 名", "循環賽第 4 名", 1),
                make_knockout_match(division, "冠軍賽", "循環賽第 1 名", "循環賽第 2 名", 2)
            ])
        else:
            rounds.append([
                make_knockout_match(division, "準決賽", "循環賽第 1 名", "循環賽第 4 名", 1),
                make_knockout_match(division, "準決賽", "循環賽第 2 名", "循環賽第 3 名", 1)
            ])
            rounds.append([
                make_knockout_match(division, "季軍賽", "準決賽負方 1", "準決賽負方 2", 2),
                make_knockout_match(division, "冠軍賽", "準決賽勝方 1", "準決賽勝方 2", 3)
            ])
        return rounds

    if len(teams) <= 7:
        rounds = round_robin(teams, division)
        rounds.append([
            make_knockout_match(division, "準決賽", "單組第 1 名", "單組第 4 名", 1),
            make_knockout_match(division, "準決賽", "單組第 2 名", "單組第 3 名", 1)
        ])
    else:
        shuffled = list(teams)
        random.shuffle(shuffled)
        ga, gb = shuffled[::2], shuffled[1::2]
        rounds = round_robin(ga, division, "A 組") + round_robin(gb, division, "B 組")
        rounds.append([
            make_knockout_match(division, "半準決賽", "A 組第 1 名", "B 組第 4 名", 1),
            make_knockout_match(division, "半準決賽", "A 組第 2 名", "B 組第 3 名", 1),
            make_knockout_match(division, "半準決賽", "B 組第 1 名", "A 組第 4 名", 1),
            make_knockout_match(division, "半準決賽", "B 組第 2 名", "A 組第 3 名", 1)
        ])
        rounds.append([
            make_knockout_match(division, "準決賽", "八強勝方 1", "八強勝方 2", 2),
            make_knockout_match(division, "準決賽", "八強勝方 3", "八強勝方 4", 2)
        ])

    rounds.append([
        make_knockout_match(division, "季軍賽", "準決賽負方 1", "準決賽負方 2", 3),
        make_knockout_match(division, "冠軍賽", "準決賽勝方 1", "準決賽勝方 2", 4)
    ])
    return rounds

def flatten_rounds(rounds):
    return [dict(x, round_no=i) for i, r in enumerate(rounds, 1) for x in r]

# ------------------------- 排程引擎 -------------------------
def generate_schedule(divisions, day_configs, duration, courts, rest_min, direction, events):
    pending = {k: list(v) for k, v in divisions.items()}
    slot_len = timedelta(minutes=duration)
    all_slots = []
    for day_i, cfg in enumerate(day_configs):
        d_date = cfg['date']
        st_time, en_time = cfg['start'], cfg['end']
        cur = datetime.combine(d_date, st_time)
        end = datetime.combine(d_date, en_time)
        while cur + slot_len <= end:
            if not any(cur < e['end'] and cur + slot_len > e['start'] for e in events):
                all_slots.append((day_i, cur))
            cur += slot_len

    all_slots.sort(key=lambda x: x[1], reverse=(direction == "逆向"))
    rows = []
    last_div = None
    match_no = 1

    # 為達成「女子季軍 -> 男子季軍 -> 女子冠軍 -> 男子冠軍」交錯順序，排序優先度
    div_order = {'女子組': 0, '男子組': 1}

    for day_i, cur in all_slots:
        if not any(pending.values()):
            break
        eligible = []
        global_min_phase = min(x['phase'] for fixtures in pending.values() if fixtures for x in fixtures)
        for div, fixtures in pending.items():
            if fixtures:
                min_phase = min(x['phase'] for x in fixtures)
                if min_phase == global_min_phase:
                    eligible.extend((div, x) for x in fixtures if x['phase'] == min_phase)

        eligible.sort(key=lambda z: (
            z[1]['phase'],
            div_order.get(z[0], 9) if z[1]['phase'] > 0 else (z[0] == last_div),
            z[1]['round_no']
        ))

        chosen = []
        used = set()
        for target_rest in (max(120, rest_min), rest_min, 0):
            for div, fx in eligible:
                if len(chosen) >= courts:
                    break
                if any(y is fx for _, y in chosen):
                    continue
                teams = [f"{div}_{fx['home_code']}", f"{div}_{fx['away_code']}"] if fx['match_type'] == 'real' else []
                if any(t in used for t in teams):
                    continue
                okay = True
                if target_rest > 0:
                    for team_key in teams:
                        previous = [r for r in rows if r['match_type'] == 'real' and team_key in (r['_home_key'], r['_away_key'])]
                        if previous:
                            last = max(previous, key=lambda r: r['_end'])
                            gap = (cur - last['_end']).total_seconds() / 60
                            between = [
                                r for r in rows
                                if r['match_type'] == 'real'
                                and r['_start'] >= last['_end']
                                and r['_end'] <= cur
                                and team_key not in (r['_home_key'], r['_away_key'])
                            ]
                            if gap < target_rest or not between:
                                okay = False
                                break
                if okay:
                    chosen.append((div, fx))
                    used.update(teams)
            if len(chosen) >= courts:
                break

        for court_i, (div, fx) in enumerate(chosen, 1):
            end = cur + slot_len
            short_div = "男子" if div == "男子組" else ("女子" if div == "女子組" else div)
            rows.append({
                'day_idx': day_i,
                '日期': cur.date().strftime('%d/%m/%Y'),
                '星期': '一二三四五六日'[cur.weekday()],
                '時間': f"{cur:%H%M}-{end:%H%M}",
                '場次': match_no,
                '場次顯示': str(match_no) if fx['match_type'] == 'real' else f"{match_no}({short_div}{fx['stage']})",
                '球場': f"球場 {court_i}",
                '組別': div,
                '賽事階段': fx['stage'],
                '分組': fx['group'],
                '主隊代號': fx['home_code'],
                '主隊': fx['home'],
                '主隊得分': '',
                '客隊得分': '',
                '客隊代號': fx['away_code'],
                '客隊': fx['away'],
                '_home_key': f"{div}_{fx['home_code']}" if fx['match_type'] == 'real' else '',
                '_away_key': f"{div}_{fx['away_code']}" if fx['match_type'] == 'real' else '',
                '_start': cur,
                '_end': end,
                'match_type': fx['match_type'],
                '_phase': fx['phase'],
                '_round': fx['round_no']
            })
            pending[div].remove(fx)
            match_no += 1
            last_div = div

    for e in events:
        s = e['start']
        en = e['end']
        rows.append({
            'day_idx': e['day_idx'],
            '日期': s.date().strftime('%d/%m/%Y'),
            '星期': '一二三四五六日'[s.weekday()],
            '時間': f"{s:%H%M}-{en:%H%M}",
            '場次': '',
            '場次顯示': '',
            '球場': '全場',
            '組別': '活動',
            '賽事階段': e['name'],
            '分組': '',
            '主隊代號': '',
            '主隊': e['name'],
            '主隊得分': '',
            '客隊得分': '',
            '客隊代號': '',
            '客隊': '',
            '_home_key': '',
            '_away_key': '',
            '_start': s,
            '_end': en,
            'match_type': 'event',
            '_phase': -1,
            '_round': -1
        })

    rows.sort(key=lambda r: (r['_start'], r['球場']))
    re_number_matches(rows)
    return rows, sum(map(len, pending.values()))

def re_number_matches(rows):
    n = 1
    for r in rows:
        if r['match_type'] != 'event':
            r['場次'] = n
            short_div = "男子" if r['組別'] == "男子組" else ("女子" if r['組別'] == "女子組" else r['組別'])
            r['場次顯示'] = str(n) if r['match_type'] == 'real' else f"{n}({short_div}{r['賽事階段']})"
            n += 1

def rows_to_slots(rows):
    grouped = defaultdict(list)
    for r in rows:
        grouped[r['_start']].append(dict(r))
    return [
        dict(
            id=f"slot_{i}",
            start=dt,
            day_idx=grouped[dt][0]['day_idx'],
            rows=grouped.get(dt, []),
            movable=not any(x['match_type'] == 'event' for x in grouped.get(dt, []))
        )
        for i, dt in enumerate(sorted(grouped))
    ]

def validate_schedule(slots, rest_min):
    games = []
    for sl in slots:
        for r in sl['rows']:
            if r['match_type'] in ('real', 'placeholder'):
                games.append(r)

    real_games = [g for g in games if g['match_type'] == 'real']
    for i, a in enumerate(real_games):
        for b in real_games[i + 1:]:
            shared = {a['_home_key'], a['_away_key']} & {b['_home_key'], b['_away_key']}
            shared.discard('')
            if not shared:
                continue
            first, second = (a, b) if a['_start'] <= b['_start'] else (b, a)
            gap = (second['_start'] - first['_end']).total_seconds() / 60
            between = [
                x for x in real_games
                if x['_start'] >= first['_end']
                and x['_end'] <= second['_start']
                and not ({x['_home_key'], x['_away_key']} & shared)
            ]
            if gap < rest_min or not between:
                team_names = []
                for k in shared:
                    if k == a['_home_key']:
                        team_names.append(f"{a['主隊代號']} {a['主隊']}".strip())
                    else:
                        team_names.append(f"{a['客隊代號']} {a['客隊']}".strip())
                return f"隊伍「{', '.join(team_names)}」休息不足 {rest_min} 分鐘或中間沒有其他隊伍賽事。"

    # 每個組別的階段不可倒序
    rank = {'循環賽': 0, '半準決賽': 1, '準決賽': 2, '季軍賽': 3, '冠軍賽': 4}
    for div in {g['組別'] for g in games}:
        ordered = sorted([g for g in games if g['組別'] == div], key=lambda g: g['_start'])
        phases = [rank.get(g['賽事階段'], 0) for g in ordered]
        if phases != sorted(phases):
            return f"{div} 的賽事階段次序（循環賽 → 淘汰賽 → 季軍賽 → 冠軍賽）不能倒轉。"
    return None

# ------------------------- 專屬 Excel 排版生成器 -------------------------
def make_excel(rows, men_teams, women_teams, title, subtitle, day_configs, courts):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = '賽程表'

    font_family = 'Microsoft JhengHei'
    f_title = Font(name=font_family, size=12, bold=True, color='FFFFFF')
    f_header = Font(name=font_family, size=11, bold=True, color='000000')
    f_cell = Font(name=font_family, size=11, bold=False, color='000000')
    f_bold = Font(name=font_family, size=11, bold=True, color='000000')

    fill_black = PatternFill('solid', fgColor='000000')
    fill_white = PatternFill('solid', fgColor='FFFFFF')
    fill_men = PatternFill('solid', fgColor='DCE6F1')      # 柔和淺藍
    fill_women = PatternFill('solid', fgColor='F2DCDB')    # 柔和淺粉紅
    fill_bar = PatternFill('solid', fgColor='A6A6A6')      # 灰色分隔列

    thin_side = Side(style='thin', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # 計算總欄數：單球場為 9 欄 (A-I)，雙球場則在右邊擴充第二個球場 (A-P)
    total_schedule_cols = 2 + 7 * courts

    # 1. 頂部黑色大標題
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_schedule_cols)
    ws.cell(1, 1, title)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_schedule_cols)
    ws.cell(2, 1, subtitle)

    for r_idx in (1, 2):
        ws.row_dimensions[r_idx].height = 20
        for c_idx in range(1, total_schedule_cols + 1):
            c = ws.cell(r_idx, c_idx)
            c.font = f_title
            c.fill = fill_black
            c.alignment = align_center

    # 按日分組賽事
    rows_by_day = defaultdict(list)
    for r in sorted(rows, key=lambda x: (x['_start'], x['球場'])):
        rows_by_day[r['day_idx']].append(r)

    cur_row = 3
    for day_i, cfg in enumerate(day_configs):
        day_rows = rows_by_day.get(day_i, [])
        if not day_rows:
            continue

        # 2. 每日兩行表頭 (Header Top & Header Bottom)
        h_top = cur_row
        h_bot = cur_row + 1
        ws.row_dimensions[h_top].height = 19
        ws.row_dimensions[h_bot].height = 20

        ws.merge_cells(start_row=h_top, start_column=1, end_row=h_bot, end_column=1)
        ws.cell(h_top, 1, "日期\n地點")

        ws.merge_cells(start_row=h_top, start_column=2, end_row=h_bot, end_column=2)
        ws.cell(h_top, 2, "時間")

        for c_i in range(courts):
            base_col = 3 + c_i * 7
            ws.merge_cells(start_row=h_top, start_column=base_col, end_row=h_top, end_column=base_col + 6)
            ws.cell(h_top, base_col, "" if courts == 1 else f"球場 {c_i + 1}")

            ws.cell(h_bot, base_col, "場次")
            ws.merge_cells(start_row=h_bot, start_column=base_col + 1, end_row=h_bot, end_column=base_col + 2)
            ws.cell(h_bot, base_col + 1, "主隊")
            ws.merge_cells(start_row=h_bot, start_column=base_col + 3, end_row=h_bot, end_column=base_col + 4)
            ws.cell(h_bot, base_col + 3, "比數")
            ws.merge_cells(start_row=h_bot, start_column=base_col + 5, end_row=h_bot, end_column=base_col + 6)
            ws.cell(h_bot, base_col + 5, "客隊")

        for r_idx in (h_top, h_bot):
            for c_idx in range(1, total_schedule_cols + 1):
                cell = ws.cell(r_idx, c_idx)
                cell.font = f_header
                cell.fill = fill_white
                cell.border = border_all
                cell.alignment = align_center

        cur_row += 2
        day_start_row = cur_row

        slots_in_day = defaultdict(list)
        for r in day_rows:
            slots_in_day[r['_start']].append(r)

        sorted_times = sorted(slots_in_day.keys())
        has_closing_event = False

        for dt in sorted_times:
            slot_items = slots_in_day[dt]
            ws.row_dimensions[cur_row].height = 22
            time_str = slot_items[0]['時間']
            ws.cell(cur_row, 2, time_str).font = f_cell
            ws.cell(cur_row, 2).fill = fill_white
            ws.cell(cur_row, 2).alignment = align_center
            ws.cell(cur_row, 2).border = border_all

            if slot_items[0]['match_type'] == 'event':
                ev_name = slot_items[0]['賽事階段']
                if any(k in ev_name for k in ('頒獎', '閉幕')):
                    has_closing_event = True
                ws.merge_cells(start_row=cur_row, start_column=3, end_row=cur_row, end_column=total_schedule_cols)
                ws.cell(cur_row, 3, ev_name)
                for c_idx in range(3, total_schedule_cols + 1):
                    c = ws.cell(cur_row, c_idx)
                    c.font = f_bold
                    c.fill = fill_bar
                    c.alignment = align_center
                    c.border = border_all
            else:
                for c_i in range(courts):
                    base_col = 3 + c_i * 7
                    court_label = f"球場 {c_i + 1}"
                    match = next((x for x in slot_items if x['球場'] == court_label), None)
                    if match:
                        row_fill = fill_men if match['組別'] == '男子組' else (fill_women if match['組別'] == '女子組' else fill_white)
                        ws.cell(cur_row, base_col, match['場次顯示'])
                        if match['match_type'] == 'real':
                            ws.cell(cur_row, base_col + 1, match['主隊代號'])
                            ws.cell(cur_row, base_col + 2, "" if match['主隊'] == match['主隊代號'] else match['主隊'])
                            ws.cell(cur_row, base_col + 3, match['主隊得分'])
                            ws.cell(cur_row, base_col + 4, match['客隊得分'])
                            ws.cell(cur_row, base_col + 5, match['客隊代號'])
                            ws.cell(cur_row, base_col + 6, "" if match['客隊'] == match['客隊代號'] else match['客隊'])
                        else:
                            for offset in range(1, 7):
                                ws.cell(cur_row, base_col + offset, "")
                        for offset in range(7):
                            c = ws.cell(cur_row, base_col + offset)
                            c.font = f_cell
                            c.fill = row_fill
                            c.alignment = align_center
                            c.border = border_all
                    else:
                        for offset in range(7):
                            c = ws.cell(cur_row, base_col + offset, "")
                            c.font = f_cell
                            c.fill = fill_white
                            c.alignment = align_center
                            c.border = border_all

            cur_row += 1

        if not has_closing_event:
            ws.row_dimensions[cur_row].height = 22
            ws.cell(cur_row, 2, "").fill = fill_white
            ws.cell(cur_row, 2).border = border_all
            chinese_nums = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
            day_cn = chinese_nums[day_i] if day_i < len(chinese_nums) else str(day_i + 1)
            ws.merge_cells(start_row=cur_row, start_column=3, end_row=cur_row, end_column=total_schedule_cols)
            ws.cell(cur_row, 3, f"第{day_cn}日賽事 完結")
            for c_idx in range(3, total_schedule_cols + 1):
                c = ws.cell(cur_row, c_idx)
                c.font = f_bold
                c.fill = fill_bar
                c.alignment = align_center
                c.border = border_all
            cur_row += 1

        day_end_row = cur_row - 1

        ws.merge_cells(start_row=day_start_row, start_column=1, end_row=day_end_row, end_column=1)
        d_obj = cfg['date']
        weekday_cn = '一二三四五六日'[d_obj.weekday()]
        venue_text = cfg['venue'].strip()
        day_label = f"DAY {day_i + 1}\n{d_obj.strftime('%d/%m/%Y')}({weekday_cn})"
        if venue_text:
            day_label += f"\n{venue_text}"
        ws.cell(day_start_row, 1, day_label)
        for r_idx in range(day_start_row, day_end_row + 1):
            c = ws.cell(r_idx, 1)
            c.font = f_cell
            c.fill = fill_white
            c.alignment = align_center
            c.border = border_all

        ws.row_dimensions[cur_row].height = 18
        cur_row += 1

    # 3. 右側並列「男子組／女子組」隊伍代號對照表
    side_col = total_schedule_cols + 2
    side_row = 6

    for div_label, t_list, div_fill in [('男子組', men_teams, fill_men), ('女子組', women_teams, fill_women)]:
        if not t_list:
            continue
        start_r = side_row
        end_r = side_row + len(t_list) - 1
        ws.merge_cells(start_row=start_r, start_column=side_col, end_row=end_r, end_column=side_col)
        ws.cell(start_r, side_col, div_label)

        for idx, t in enumerate(t_list):
            r_curr = start_r + idx
            if ws.row_dimensions[r_curr].height is None:
                ws.row_dimensions[r_curr].height = 22
            ws.cell(r_curr, side_col + 1, t['code'])
            ws.cell(r_curr, side_col + 2, "" if t['name'] == t['code'] else t['name'])
            for c_offset in range(3):
                c = ws.cell(r_curr, side_col + c_offset)
                c.font = f_cell
                c.fill = div_fill
                c.alignment = align_center
                c.border = border_all
        side_row = end_r + 1

    # 4. 設定欄寬與列印版面
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 14
    for c_i in range(courts):
        base_col = 3 + c_i * 7
        col_widths = [15, 7, 14, 6, 6, 7, 14]
        for offset, w in enumerate(col_widths):
            ws.column_dimensions[get_column_letter(base_col + offset)].width = w

    ws.column_dimensions[get_column_letter(total_schedule_cols + 1)].width = 4
    ws.column_dimensions[get_column_letter(side_col)].width = 14
    ws.column_dimensions[get_column_letter(side_col + 1)].width = 9
    ws.column_dimensions[get_column_letter(side_col + 2)].width = 16

    ws.sheet_view.showGridLines = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# ------------------------- 側邊欄設定介面 -------------------------
with st.sidebar:
    st.header('⚙️ 賽事設定')
    tournament_name = st.text_input('賽事名稱（頂部第一行）', '巧固球錦標賽')
    tournament_subtitle = st.text_input('組別副標題（頂部第二行，選填）', '', placeholder='例如：大專組 / 中學組（可留空）')
    default_venue = st.text_area('比賽場地（選填，支援換行）', '', placeholder='例如：明愛胡振中書院\n(東涌松逸街4號)', height=68)

    days = st.number_input('比賽日數', 1, 14, 1)
    direction_label = st.radio('排程方向', ['正向（由開始時間排起）', '逆向（由結束時間倒推）'])
    direction = '逆向' if direction_label.startswith('逆向') else '正向'

    day_configs = []
    st.markdown('### 📅 每日比賽日期與時間')
    base_date = date.today()
    for i in range(int(days)):
        with st.expander(f'DAY {i+1} 設定', expanded=(i == 0)):
            default_d = base_date + timedelta(days=i)
            d_date = st.date_input(f'DAY {i+1} 日期', value=default_d, key=f'day_date_{i}')
            a = st.time_input(f'DAY {i+1} 開始時間', time(9, 0), key=f'day_a_{i}')
            b = st.time_input(f'DAY {i+1} 結束時間', time(18, 0), key=f'day_b_{i}')
            v = st.text_area(f'DAY {i+1} 地點（留空則跟隨上方比賽場地）', value='', key=f'day_v_{i}', height=68)
            day_configs.append({'date': d_date, 'start': a, 'end': b, 'venue': v if v.strip() else default_venue})

    st.markdown('### 🎉 開幕禮／頒獎禮（選填）')
    opening = st.checkbox('首日加入開幕禮', value=False)
    opening_time = st.time_input('開幕禮開始時間', time(9, 0))
    opening_duration = st.number_input('開幕禮時長（分鐘）', 5, 240, 30, 5)

    closing = st.checkbox('最後一日加入頒獎禮／閉幕禮', value=False)
    closing_name = st.text_input('最後一日活動名稱', '頒獎禮')
    closing_time = st.time_input('頒獎禮開始時間', time(17, 30))
    closing_duration = st.number_input('頒獎禮時長（分鐘）', 5, 240, 20, 5)

    duration_choice = st.radio('每場比賽時段', ['60 分鐘（15 分鐘 × 3 節）', '45 分鐘（12 分鐘 × 3 節）'])
    duration = 60 if duration_choice.startswith('60') else 45
    court_mode = st.radio('球場安排', ['單球場', '雙球場', '同時生成單、雙球場版本'])
    small_format = st.selectbox(
        '隊伍不多於 5 隊時的賽制',
        ['純循環賽', '循環賽＋決賽及季軍戰（不設準決賽）', '循環賽＋準決賽、決賽及季軍戰']
    )

# ------------------------- 參賽隊伍輸入區 -------------------------
st.subheader('👥 輸入參賽隊伍')
st.caption('可留空直接生成賽程（男子組預設為 `MA, MB, MC...`，女子組預設為 `WA, WB, WC...`）。如需自訂，每行輸入一隊（例如輸入 `HKUST` 或 `UMA HKUST`）。')

c1, c2 = st.columns(2)
with c1:
    men_count = st.number_input('男子組隊數', 0, 20, 4)
    men_text = st.text_area(
        '男子組隊伍名稱（每行一隊；留空自動用 MA, MB, MC...）',
        value='',
        placeholder='留空則自動使用 MA, MB, MC...\n亦可每行輸入隊名（如：HKUST）或「代號 隊名」（如：UMA HKUST）',
        height=150
    )
with c2:
    women_count = st.number_input('女子組隊數', 0, 20, 4)
    women_text = st.text_area(
        '女子組隊伍名稱（每行一隊；留空自動用 WA, WB, WC...）',
        value='',
        placeholder='留空則自動使用 WA, WB, WC...\n亦可每行輸入隊名（如：HKMU）或「代號 隊名」（如：UWA HKMU）',
        height=150
    )

men_teams = parse_teams_with_codes(men_text, men_count, 'M')
women_teams = parse_teams_with_codes(women_text, women_count, 'W')

rest_min = 60
st.info('休息規則：每隊兩場比賽之間最少休息 60 分鐘，並須有至少一場其他隊伍的完整比賽在中間完成；系統會優先安排更長休息時間。')

if st.button('🚀 生成賽程', type='primary', use_container_width=True):
    err = None
    if len(men_teams) < 2 and len(women_teams) < 2:
        err = '至少需要一個組別有 2 隊或以上。'
    elif any(cfg['end'] <= cfg['start'] for cfg in day_configs):
        err = '每日結束時間必須遲於開始時間。'

    if err:
        st.error(err)
    else:
        divisions = {}
        if len(men_teams) >= 2:
            divisions['男子組'] = flatten_rounds(build_division(men_teams, '男子組', small_format))
        if len(women_teams) >= 2:
            divisions['女子組'] = flatten_rounds(build_division(women_teams, '女子組', small_format))

        events = []
        if opening:
            s = datetime.combine(day_configs[0]['date'], opening_time)
            events.append({'day_idx': 0, 'name': '開幕禮', 'start': s, 'end': s + timedelta(minutes=int(opening_duration))})
        if closing:
            last_idx = len(day_configs) - 1
            s = datetime.combine(day_configs[last_idx]['date'], closing_time)
            events.append({'day_idx': last_idx, 'name': closing_name, 'start': s, 'end': s + timedelta(minutes=int(closing_duration))})

        versions = [1, 2] if court_mode.startswith('同時') else ([1] if court_mode == '單球場' else [2])
        st.session_state.schedule_versions = {}
        for court_count in versions:
            rows, unscheduled = generate_schedule(divisions, day_configs, duration, court_count, rest_min, direction, events)
            st.session_state.schedule_versions[court_count] = {
                'slots': rows_to_slots(rows),
                'unscheduled': unscheduled,
                'men_teams': men_teams,
                'women_teams': women_teams,
                'title': tournament_name,
                'subtitle': tournament_subtitle,
                'day_configs': day_configs,
                'courts': court_count,
                'duration': duration,
                'rest': rest_min
            }
            st.session_state.pop(f"last_dnd_event_{court_count}", None)

# ------------------------- 顯示、拖曳、匯出 -------------------------
if 'schedule_versions' in st.session_state and st.session_state.schedule_versions:
    for court_count, data in st.session_state.schedule_versions.items():
        st.divider()
        st.subheader(f"📋 {'單球場' if court_count == 1 else '雙球場'}賽程表")
        slots = data['slots']

        if data['unscheduled']:
            st.warning(f"⚠️ 仍有 {data['unscheduled']} 場比賽因時間不足未能排入！請增加比賽日數或延長每日結束時間。")

        if HAS_DND:
            st.caption('💡 拖曳下方賽程卡片可調動比賽先後次序；開幕禮及頒獎禮固定不動。')
            list_key = f"schedule_list_{court_count}"
            with st.container(key=list_key, border=True):
                for idx, slot in enumerate(slots):
                    if not slot['rows']:
                        continue
                    fixed = not slot['movable']
                    label_time = f"DAY {slot['day_idx']+1} ({slot['start'].strftime('%d/%m')}) {slot['rows'][0]['時間']}"
                    pieces = []
                    for r in slot['rows']:
                        if r['match_type'] == 'event':
                            pieces.append(f"📌 {r['賽事階段']}")
                        elif r['match_type'] == 'placeholder':
                            pieces.append(f"🏆 {r['場次顯示']}")
                        else:
                            h_disp = r['主隊代號'] if r['主隊'] == r['主隊代號'] else f"{r['主隊代號']} {r['主隊']}"
                            a_disp = r['客隊代號'] if r['客隊'] == r['客隊代號'] else f"{r['客隊代號']} {r['客隊']}"
                            pieces.append(f"{r['組別']}｜{h_disp} vs {a_disp}")
                    with st.container(key=f"card_{court_count}_{idx}", border=True):
                        st.markdown(('🔒 ' if fixed else '⠿ ') + f"**{label_time}**　" + ' ； '.join(pieces))

            event = dnd(
                list_key,
                exclude=[f"card_{court_count}_{i}" for i, s in enumerate(slots) if not s['movable']],
                key=f"dnd_{court_count}",
                handle=True
            )
            last_event_key = f"last_dnd_event_{court_count}"
            if event and st.session_state.get(last_event_key) != event:
                st.session_state[last_event_key] = event
                old_slots = [dict(s, rows=[dict(r) for r in s['rows']]) for s in slots]
                card_slots = [s for s in slots if s['rows']]
                movable_card_slots = [s for s in card_slots if s['movable']]

                if event.from_index < len(movable_card_slots) and event.to_index < len(movable_card_slots):
                    apply_move(event, movable_card_slots)
                    ordered = iter(movable_card_slots)
                    new_slots = [next(ordered) if s['movable'] else s for s in card_slots]

                    time_positions = [(s['day_idx'], s['start'], s['rows'][0]['時間']) for s in card_slots if s['movable']]
                    for (d_idx, pos, t_str), s in zip(time_positions, [s for s in new_slots if s['movable']]):
                        delta = pos - s['start']
                        s['start'] = pos
                        s['day_idx'] = d_idx
                        for r in s['rows']:
                            r['_start'] += delta
                            r['_end'] += delta
                            r['day_idx'] = d_idx
                            r['日期'] = r['_start'].date().strftime('%d/%m/%Y')
                            r['星期'] = '一二三四五六日'[r['_start'].weekday()]
                            r['時間'] = t_str

                    candidate = sorted(new_slots, key=lambda x: x['start'])
                    problem = validate_schedule(candidate, data['rest'])
                    if problem:
                        st.session_state.schedule_versions[court_count]['slots'] = old_slots
                        st.error('呢個調動未能套用：' + problem)
                    else:
                        flat_temp = [r for sl in candidate for r in sl['rows']]
                        re_number_matches(flat_temp)
                        st.session_state.schedule_versions[court_count]['slots'] = candidate
                        st.success('✅ 賽程已成功調動！')
                        st.rerun()

        # 整理表格預覽資料
        flat = [dict(r) for sl in data['slots'] for r in sl['rows']]
        flat.sort(key=lambda r: (r['_start'], r['球場']))
        re_number_matches(flat)

        preview_rows = []
        for r in flat:
            is_real = (r['match_type'] == 'real')
            is_event = (r['match_type'] == 'event')
            preview_rows.append({
                'DAY': f"DAY {r['day_idx'] + 1}",
                '日期': f"{r['日期']}({r['星期']})",
                '時間': r['時間'],
                '場次': r['賽事階段'] if is_event else r['場次顯示'],
                '組別': r['組別'],
                '主隊代號': r['主隊代號'] if is_real else '',
                '主隊': ('' if r['主隊'] == r['主隊代號'] else r['主隊']) if is_real else '',
                '比數': 'vs' if is_real else '',
                '客隊代號': r['客隊代號'] if is_real else '',
                '客隊': ('' if r['客隊'] == r['客隊代號'] else r['客隊']) if is_real else ''
            })

        cols = ['DAY', '日期', '時間', '場次', '組別', '主隊代號', '主隊', '比數', '客隊代號', '客隊']
        visible_df = pd.DataFrame(preview_rows, columns=cols)
        st.dataframe(visible_df, use_container_width=True, hide_index=True)

        # 生成 Excel 並提供下載
        file_prefix = f"{data['title']}_{data['subtitle']}" if data['subtitle'].strip() else data['title']
        excel_bytes = make_excel(
            flat,
            data['men_teams'],
            data['women_teams'],
            data['title'],
            data['subtitle'],
            data['day_configs'],
            court_count
        )
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                '⬇️ 下載 Excel 賽程表（專屬排版）',
                excel_bytes,
                file_name=f"{file_prefix}_{court_count}球場.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
                use_container_width=True,
                key=f"excel_{court_count}"
            )
        with col_dl2:
            st.download_button(
                '⬇️ 下載 CSV 賽程資料',
                visible_df.to_csv(index=False).encode('utf-8-sig'),
                file_name=f"{file_prefix}_{court_count}球場.csv",
                mime='text/csv',
                use_container_width=True,
                key=f"csv_{court_count}"
            )
