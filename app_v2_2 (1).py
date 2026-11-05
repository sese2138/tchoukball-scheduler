import streamlit as st
from streamlit_dnd import dnd, apply_move
import pandas as pd
from datetime import datetime, date, time, timedelta
from itertools import combinations
from io import BytesIO
from collections import defaultdict
import random

st.set_page_config(page_title="巧固球賽程生成器 V2.2", page_icon="🏆", layout="wide")
st.markdown("""<style>.main .block-container{max-width:1450px;padding-top:1.2rem} h1{font-weight:800}</style>""", unsafe_allow_html=True)
st.title("🏆 巧固球賽程生成器 V2.2")
st.caption("男女組別・單／雙球場・循環賽及淘汰賽・拖曳調動・Excel／CSV 匯出")

# ------------------------- 賽制 -------------------------
def round_robin(teams, division, group_name="單組"):
    arr = list(teams)
    if len(arr) < 2: return []
    if len(arr) % 2: arr.append(None)
    out=[]
    for rnd in range(len(arr)-1):
        games=[]
        for i in range(len(arr)//2):
            a,b=arr[i],arr[-1-i]
            if a is not None and b is not None:
                games.append(dict(division=division,stage="循環賽",group=group_name,home=a,away=b,phase=0,match_type="real"))
        out.append(games)
        arr=[arr[0],arr[-1]]+arr[1:-1]
    return out

def knockout(stage, group, home, away, phase):
    return dict(division="",stage=stage,group=group,home=home,away=away,phase=phase,match_type="placeholder")

def build_division(teams, division, small_format):
    if len(teams)<2: return []
    rounds=[]
    if len(teams)<=5:
        rounds=round_robin(teams,division)
        if small_format=="純循環賽" or len(teams)<4: return rounds
        if small_format=="循環賽＋決賽及季軍戰（不設準決賽）":
            rounds.append([dict(division=division,stage="決賽",group="淘汰賽",home="循環賽第 1 名",away="循環賽第 2 名",phase=1,match_type="placeholder"),dict(division=division,stage="季軍戰",group="淘汰賽",home="循環賽第 3 名",away="循環賽第 4 名",phase=1,match_type="placeholder")])
        else:
            rounds.append([dict(division=division,stage="準決賽",group="淘汰賽",home="循環賽第 1 名",away="循環賽第 4 名",phase=1,match_type="placeholder"),dict(division=division,stage="準決賽",group="淘汰賽",home="循環賽第 2 名",away="循環賽第 3 名",phase=1,match_type="placeholder")])
            rounds.append([dict(division=division,stage="決賽",group="淘汰賽",home="準決賽勝方 1",away="準決賽勝方 2",phase=2,match_type="placeholder"),dict(division=division,stage="季軍戰",group="淘汰賽",home="準決賽負方 1",away="準決賽負方 2",phase=2,match_type="placeholder")])
        for r in rounds:
            for x in r: x['division']=division
        return rounds
    if len(teams)<=7:
        rounds=round_robin(teams,division)
        rounds.append([dict(division=division,stage="準決賽",group="淘汰賽",home="單組第 1 名",away="單組第 4 名",phase=1,match_type="placeholder"),dict(division=division,stage="準決賽",group="淘汰賽",home="單組第 2 名",away="單組第 3 名",phase=1,match_type="placeholder")])
    else:
        shuffled=list(teams); random.shuffle(shuffled)
        ga,gb=shuffled[::2],shuffled[1::2]
        rounds=round_robin(ga,division,"A 組")+round_robin(gb,division,"B 組")
        rounds.append([dict(division=division,stage="半準決賽",group="淘汰賽",home="A 組第 1 名",away="B 組第 4 名",phase=1,match_type="placeholder"),dict(division=division,stage="半準決賽",group="淘汰賽",home="A 組第 2 名",away="B 組第 3 名",phase=1,match_type="placeholder"),dict(division=division,stage="半準決賽",group="淘汰賽",home="B 組第 1 名",away="A 組第 4 名",phase=1,match_type="placeholder"),dict(division=division,stage="半準決賽",group="淘汰賽",home="B 組第 2 名",away="A 組第 3 名",phase=1,match_type="placeholder")])
    rounds.append([dict(division=division,stage="決賽",group="淘汰賽",home="準決賽勝方 1",away="準決賽勝方 2",phase=3,match_type="placeholder"),dict(division=division,stage="季軍戰",group="淘汰賽",home="準決賽負方 1",away="準決賽負方 2",phase=3,match_type="placeholder")])
    for r in rounds:
        for x in r: x['division']=division
    return rounds

def flatten_rounds(rounds):
    return [dict(x,round_no=i) for i,r in enumerate(rounds,1) for x in r]

def parse_teams(raw,count,prefix):
    names=[x.strip() for x in raw.splitlines() if x.strip()]
    defaults=[f"{prefix} {chr(65+i)}" for i in range(int(count))]
    return (names+defaults[len(names):])[:int(count)]

# ------------------------- 排程引擎 -------------------------
def generate_schedule(divisions, start_date, windows, duration, courts, rest_min, direction, events):
    pending={k:list(v) for k,v in divisions.items()}
    slot_len=timedelta(minutes=duration)
    all_slots=[]
    for day_i,(st_time,en_time) in enumerate(windows):
        cur=datetime.combine(start_date+timedelta(days=day_i),st_time)
        end=datetime.combine(start_date+timedelta(days=day_i),en_time)
        while cur+slot_len<=end:
            if not any(cur < e['end'] and cur+slot_len > e['start'] for e in events): all_slots.append(cur)
            cur+=slot_len
    all_slots.sort(reverse=direction=="逆向")
    rows=[]; last_div=None; match_no=1
    for cur in all_slots:
        if not any(pending.values()): break
        eligible=[]
        for div,fixtures in pending.items():
            if fixtures:
                min_phase=min(x['phase'] for x in fixtures)
                eligible.extend((div,x) for x in fixtures if x['phase']==min_phase)
        eligible.sort(key=lambda z:(z[0]==last_div,z[1]['round_no']))
        chosen=[]; used=set()
        for target_rest in (max(120,rest_min),rest_min):
            for div,fx in eligible:
                if len(chosen)>=courts: break
                if any(y is fx for _,y in chosen): continue
                teams=[fx['home'],fx['away']] if fx['match_type']=='real' else []
                if any(t in used for t in teams): continue
                okay=True
                for team in teams:
                    previous=[r for r in rows if r['match_type']=='real' and team in (r['主隊'],r['客隊'])]
                    if previous:
                        last=max(previous,key=lambda r:r['_end'])
                        gap=(cur-last['_end']).total_seconds()/60
                        between=[r for r in rows if r['match_type']=='real' and r['_start']>=last['_end'] and r['_end']<=cur and team not in (r['主隊'],r['客隊'])]
                        if gap<target_rest or not between: okay=False; break
                if okay:
                    chosen.append((div,fx)); used.update(teams)
            if len(chosen)>=courts: break
        for court_i,(div,fx) in enumerate(chosen,1):
            end=cur+slot_len
            rows.append({'日期':cur.date().isoformat(),'星期':'一二三四五六日'[cur.weekday()],'時間':f"{cur:%H:%M}–{end:%H:%M}",'場次':match_no,'球場':f"球場 {court_i}",'組別':div,'賽事階段':fx['stage'],'分組':fx['group'],'主隊':fx['home'],'主隊得分':'','客隊得分':'','客隊':fx['away'],'_start':cur,'_end':end,'match_type':fx['match_type'],'_phase':fx['phase'],'_round':fx['round_no']})
            pending[div].remove(fx); match_no+=1; last_div=div
    for e in events:
        s=e['start']; en=e['end']
        rows.append({'日期':s.date().isoformat(),'星期':'一二三四五六日'[s.weekday()],'時間':f"{s:%H:%M}–{en:%H:%M}",'場次':'','球場':'全場','組別':'活動','賽事階段':e['name'],'分組':'','主隊':e['name'],'主隊得分':'','客隊得分':'','客隊':'','_start':s,'_end':en,'match_type':'event','_phase':-1,'_round':-1})
    rows.sort(key=lambda r:(r['_start'],r['球場']))
    n=1
    for r in rows:
        if r['match_type']!='event': r['場次']=n; n+=1
    return rows,sum(map(len,pending.values()))

def rows_to_slots(rows, courts):
    grouped=defaultdict(list)
    for r in rows: grouped[r['_start']].append(dict(r))
    return [dict(id=f"slot_{i}",start=dt,rows=grouped.get(dt,[]),movable=not any(x['match_type']=='event' for x in grouped.get(dt,[]))) for i,dt in enumerate(sorted(grouped))]

def validate_schedule(slots, rest_min):
    games=[]
    for sl in slots:
        for r in sl['rows']:
            if r['match_type']=='real': games.append(r)
    for i,a in enumerate(games):
        for b in games[i+1:]:
            shared=set((a['主隊'],a['客隊'])) & set((b['主隊'],b['客隊']))
            if not shared: continue
            first,second=(a,b) if a['_start']<=b['_start'] else (b,a)
            gap=(second['_start']-first['_end']).total_seconds()/60
            between=[x for x in games if x['_start']>=first['_end'] and x['_end']<=second['_start'] and not (set((x['主隊'],x['客隊'])&set(shared)))]
            if gap<rest_min or not between: return f"隊伍「{', '.join(shared)}」休息不足或中間沒有其他隊伍賽事。"
    # 每個組別的階段不可倒序
    rank={'循環賽':0,'半準決賽':1,'準決賽':2,'決賽':3,'季軍戰':3}
    for div in {g['組別'] for g in games}:
        ordered=sorted([g for g in games if g['組別']==div],key=lambda g:g['_start'])
        phases=[rank.get(g['賽事階段'],0) for g in ordered]
        if phases!=sorted(phases): return f"{div} 的淘汰賽階段次序不能倒轉。"
    return None

def make_excel(rows, team_rows, title, venue, courts):
    from openpyxl import Workbook
    from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
    from openpyxl.utils import get_column_letter
    wb=Workbook(); ws=wb.active; ws.title='賽程表'
    headers=['日期','星期','時間','場次','球場','組別','賽事階段','分組','主隊','主隊得分','客隊得分','客隊']
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(headers)); ws.cell(1,1,f"{title}｜{venue}")
    ws.cell(1,1).font=Font(name='Microsoft JhengHei',size=18,bold=True,color='FFFFFF'); ws.cell(1,1).fill=PatternFill('solid',fgColor='26364A'); ws.cell(1,1).alignment=Alignment(horizontal='center')
    for c,h in enumerate(headers,1):
        x=ws.cell(3,c,h); x.font=Font(bold=True,color='FFFFFF'); x.fill=PatternFill('solid',fgColor='40536B'); x.alignment=Alignment(horizontal='center')
    thin=Side(style='thin',color='B7C1CC')
    for ri,r in enumerate(sorted(rows,key=lambda x:(x['_start'],x['球場'])),4):
        for ci,h in enumerate(headers,1):
            val=r.get(h,''); cell=ws.cell(ri,ci,val); cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
            color='D9D9D9' if r['match_type']=='event' else ('D1E3F4' if r['組別']=='男子組' else 'EAD0DF' if r['組別']=='女子組' else 'FFFFFF')
            cell.fill=PatternFill('solid',fgColor=color)
    widths=[14,8,16,8,12,12,14,12,25,12,12,25]
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes='A4'; ws.sheet_view.showGridLines=False; ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.sheet_properties.pageSetUpPr.fitToPage=True
    tw=wb.create_sheet('隊伍名單'); tw.append(['組別','隊伍名稱'])
    for x in team_rows: tw.append(list(x))
    for c in tw[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='40536B')
    tw.column_dimensions['A'].width=18; tw.column_dimensions['B'].width=32
    out=BytesIO(); wb.save(out); return out.getvalue()

# ------------------------- 設定介面 -------------------------
with st.sidebar:
    st.header('⚙️ 賽事設定')
    tournament_name=st.text_input('賽事名稱','巧固球錦標賽')
    venue=st.text_input('比賽場地','')
    start_date=st.date_input('比賽首日',value=date.today())
    days=st.number_input('比賽日數',1,14,1)
    direction_label=st.radio('排程方向',['正向（由開始時間排起）','逆向（由結束時間倒推）'])
    direction='逆向' if direction_label.startswith('逆向') else '正向'
    windows=[]
    st.markdown('### 每日比賽時間')
    for i in range(int(days)):
        with st.expander(f'第 {i+1} 日時間',expanded=i==0):
            a=st.time_input(f'第 {i+1} 日開始',time(9,0),key=f'day_a_{i}')
            b=st.time_input(f'第 {i+1} 日結束',time(18,0),key=f'day_b_{i}')
            windows.append((a,b))
    st.markdown('### 開幕禮／閉幕禮（選填）')
    opening=st.checkbox('加入開幕禮（首日）'); opening_time=st.time_input('開幕禮開始',time(9,0)); opening_duration=st.number_input('開幕禮時長（分鐘）',5,240,30,5)
    closing=st.checkbox('加入閉幕禮（最後一日）'); closing_time=st.time_input('閉幕禮開始',time(17,30)); closing_duration=st.number_input('閉幕禮時長（分鐘）',5,240,30,5)
    duration_choice=st.radio('每場比賽時段',['60 分鐘（15 分鐘 × 3 節）','45 分鐘（12 分鐘 × 3 節）'])
    duration=60 if duration_choice.startswith('60') else 45
    court_mode=st.radio('球場安排',['單球場','雙球場','同時生成單、雙球場版本'])
    small_format=st.selectbox('隊伍不多於 5 隊時的賽制',['純循環賽','循環賽＋準決賽、決賽及季軍戰','循環賽＋決賽及季軍戰（不設準決賽）'])

st.subheader('👥 輸入參賽隊伍')
c1,c2=st.columns(2)
with c1:
    men_count=st.number_input('男子組隊數',0,20,4); men_text=st.text_area('男子組隊伍名稱（每行一隊；留空使用預設）','\n'.join(f'男子 {chr(65+i)}' for i in range(int(men_count))),height=150)
with c2:
    women_count=st.number_input('女子組隊數',0,20,4); women_text=st.text_area('女子組隊伍名稱（每行一隊；留空使用預設）','\n'.join(f'女子 {chr(65+i)}' for i in range(int(women_count))),height=150)
men_teams=parse_teams(men_text,men_count,'男子'); women_teams=parse_teams(women_text,women_count,'女子')
rest_min=60
st.info('休息規則：每隊兩場比賽之間最少休息 60 分鐘，並須有至少一場其他隊伍的完整比賽在中間完成；系統會盡量安排更長休息。')

if st.button('🚀 生成賽程',type='primary',use_container_width=True):
    err=None
    if len(men_teams)<2 and len(women_teams)<2: err='至少需要一個組別有 2 隊或以上。'
    elif any(b<=a for a,b in windows): err='每日結束時間必須遲於開始時間。'
    elif opening and not (windows[0][0]<=opening_time and (datetime.combine(start_date,opening_time)+timedelta(minutes=int(opening_duration))).time()<=windows[0][1]): err='開幕禮必須完整位於第一日比賽時間內。'
    elif closing and not (windows[-1][0]<=closing_time and (datetime.combine(start_date+timedelta(days=int(days)-1),closing_time)+timedelta(minutes=int(closing_duration))).time()<=windows[-1][1]): err='閉幕禮必須完整位於最後一日比賽時間內。'
    if err: st.error(err)
    else:
        divisions={}; team_rows=[]
        if len(men_teams)>=2: divisions['男子組']=flatten_rounds(build_division(men_teams,'男子組',small_format)); team_rows += [('男子組',t) for t in men_teams]
        if len(women_teams)>=2: divisions['女子組']=flatten_rounds(build_division(women_teams,'女子組',small_format)); team_rows += [('女子組',t) for t in women_teams]
        events=[]
        if opening:
            s=datetime.combine(start_date,opening_time); events.append({'name':'開幕禮','start':s,'end':s+timedelta(minutes=int(opening_duration))})
        if closing:
            s=datetime.combine(start_date+timedelta(days=int(days)-1),closing_time); events.append({'name':'閉幕禮','start':s,'end':s+timedelta(minutes=int(closing_duration))})
        versions=[1,2] if court_mode.startswith('同時') else ([1] if court_mode=='單球場' else [2])
        st.session_state.schedule_versions={}
        for court_count in versions:
            rows,unscheduled=generate_schedule(divisions,start_date,windows,duration,court_count,rest_min,direction,events)
            st.session_state.schedule_versions[court_count]={'slots':rows_to_slots(rows,court_count),'unscheduled':unscheduled,'team_rows':team_rows,'title':tournament_name,'venue':venue,'courts':court_count,'duration':duration,'rest':rest_min}

# ------------------------- 顯示、拖曳、匯出 -------------------------
if 'schedule_versions' in st.session_state and st.session_state.schedule_versions:
    for court_count,data in st.session_state.schedule_versions.items():
        st.divider(); st.subheader(f"{'單球場' if court_count==1 else '雙球場'}賽程")
        slots=data['slots']
        if data['unscheduled']:
            st.warning(f"仍有 {data['unscheduled']} 場未能安排。可增加比賽日數或延長每日時間。")
        st.caption('拖曳賽程卡片可調整先後時間；開幕禮及閉幕禮固定不動。雙球場同一時段會以一張卡片整組移動。')
        # 只提供已排出的時段作為拖曳位置；每個卡片代表同一時段的整組賽事。
        list_key=f"schedule_list_{court_count}"
        with st.container(key=list_key,border=True):
            for idx,slot in enumerate(slots):
                fixed=not slot['movable']
                label_time=slot['start'].strftime('%d/%m %H:%M')
                if not slot['rows']: continue
                pieces=[]
                for r in slot['rows']:
                    if r['match_type']=='event': pieces.append(f"📌 {r['賽事階段']}")
                    else: pieces.append(f"{r['組別']}｜{r['球場']}｜{r['主隊']} vs {r['客隊']}（{r['賽事階段']}）")
                with st.container(key=f"card_{court_count}_{idx}",border=True):
                    st.markdown(('🔒 ' if fixed else '⠿ ') + f"**{label_time}**  " + ' ；  '.join(pieces))
        event=dnd(list_key,exclude=[f"card_{court_count}_{i}" for i,s in enumerate(slots) if not s['movable']],key=f"dnd_{court_count}",handle=True)
        if event:
            old_slots=list(slots)
            # dnd order is applied to a parallel list of slot objects.
            movable_slots=[s for s in slots if s['movable']]
            # The rendered container includes only non-empty cards; all listed cards correspond to slots.
            card_slots=[s for s in slots if s['rows']]
            movable_card_slots=[s for s in card_slots if s['movable']]
            if event.from_index < len(movable_card_slots) and event.to_index < len(movable_card_slots):
                apply_move(event,movable_card_slots)
                ordered=iter(movable_card_slots)
                new_slots=[]
                for s in card_slots:
                    new_slots.append(next(ordered) if s['movable'] else s)
                # Reassign match groups to the original time positions; event positions remain fixed.
                time_positions=[s['start'] for s in card_slots if s['movable']]
                for pos,s in zip(time_positions,[s for s in new_slots if s['movable']]):
                    delta=pos-s['start']
                    s['start']=pos
                    for r in s['rows']:
                        r['_start']+=delta; r['_end']+=delta
                        r['日期']=r['_start'].date().isoformat(); r['星期']='一二三四五六日'[r['_start'].weekday()]; r['時間']=f"{r['_start']:%H:%M}–{r['_end']:%H:%M}"
                candidate=sorted(new_slots,key=lambda x:x['start'])
                problem=validate_schedule(candidate,data['rest'])
                if problem:
                    st.session_state.schedule_versions[court_count]['slots']=old_slots
                    st.error('呢個調動未能套用：'+problem)
                else:
                    st.session_state.schedule_versions[court_count]['slots']=candidate
                    st.success('賽程已調動。')
                st.rerun()
            else:
                st.warning('今次調動未能辨認，請再試一次。')
        flat=[]
        for sl in data['slots']:
            for r in sl['rows']: flat.append(dict(r))
        flat.sort(key=lambda r:(r['_start'],r['球場']))
        n=1
        for r in flat:
            if r['match_type']!='event': r['場次']=n; n+=1
        visible=pd.DataFrame([{k:r[k] for k in ['日期','星期','時間','場次','球場','組別','賽事階段','分組','主隊','主隊得分','客隊得分','客隊']} for r in flat])
        st.dataframe(visible,use_container_width=True,hide_index=True)
        excel=make_excel(flat,data['team_rows'],data['title'],data['venue'],court_count)
        st.download_button('⬇️ 下載 Excel 賽程及隊伍名單',excel,file_name=f"{data['title']}_{court_count}球場.xlsx",mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key=f"excel_{court_count}")
        st.download_button('⬇️ 下載 CSV 賽程',visible.to_csv(index=False).encode('utf-8-sig'),file_name=f"{data['title']}_{court_count}球場.csv",mime='text/csv',key=f"csv_{court_count}")
