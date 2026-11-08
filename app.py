"""범용 실패비용(COPQ) 관리 프로그램 — Streamlit

실행:  streamlit run app.py
"""
from __future__ import annotations

import io
import os
import shutil
from datetime import date, datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import db
import engine
import excel_io

st.set_page_config(page_title="실패비용(COPQ) 관리", page_icon="📉", layout="wide")

# ---- 데이터셋(사업장/제품) 선택: data/ 폴더의 .db 파일 하나 = 데이터셋 하나
_ds = db.list_datasets() or ["default"]
if "dataset" not in st.session_state or st.session_state.dataset not in _ds:
    st.session_state.dataset = _ds[0]
db.use_dataset(st.session_state.dataset)

st.markdown("""
<style>
div[data-testid="stMetric"] {background:#f6f8fb;border:1px solid #e3e8ef;border-radius:10px;padding:10px 14px;}
div[data-testid="stMetricValue"] {font-size:1.45rem;}
.block-container {padding-top:1.6rem;}
.small {color:#6b7280;font-size:0.85rem;}
</style>
""", unsafe_allow_html=True)

C_CONF, C_POT, C_PREV, C_CLAIM, C_REW, C_TGT = "#1F4E79", "#8FB4D9", "#70AD47", "#C55A11", "#A5A5A5", "#C00000"


# ------------------------------------------------------------------ 공통 유틸
def won(v) -> str:
    try:
        return f"{v:,.0f}{CUR}"
    except (TypeError, ValueError):
        return "-"


def pct(v) -> str:
    try:
        return f"{v * 100:,.1f}%"
    except (TypeError, ValueError):
        return "-"


def money_cols(cols):
    return {c: st.column_config.NumberColumn(c, format="localized") for c in cols}


def show(df: pd.DataFrame, rename: dict | None = None, money=(), percent=(), height=None, **kw):
    out = df.rename(columns=rename or excel_io.KOR)
    cfg = {}
    for c in out.columns:
        if not pd.api.types.is_numeric_dtype(out[c]):
            continue
        if c in percent or any(k in c for k in ["률", "대비", "누적비율"]):
            cfg[c] = st.column_config.NumberColumn(c, format="percent")
        elif c in money or any(k in c for k in ["금액", "비용", "단가", "청구", "조정액", "목표", "실적", "한도", "손실", "초과", "REWORK"]):
            cfg[c] = st.column_config.NumberColumn(c, format="localized")
            out[c] = out[c].round(0)
        else:
            cfg[c] = st.column_config.NumberColumn(c, format="localized")
    if height is not None:
        kw["height"] = height
    st.dataframe(out, hide_index=True, use_container_width=True, column_config=cfg, **kw)


def years_available() -> list[int]:
    y = db.query("SELECT DISTINCT CAST(substr(date,1,4) AS INTEGER) y FROM daily ORDER BY y")["y"].tolist()
    t = date.today().year
    return sorted(set(y + [t]))


def latest_month(year: int) -> int:
    r = db.query("SELECT MAX(date) d FROM daily WHERE substr(date,1,4)=?", (str(year),))["d"].iloc[0]
    return int(r[5:7]) if r else date.today().month


def defect_options(process: str | None = None) -> list[str]:
    t = db.read("defect_types", "active=1")
    if process:
        t = t[t["process"] == process]
    return t["defect_type"].tolist()


# ------------------------------------------------------------------ 사이드바
CUR = db.get_setting("currency") or "원"
SYS = db.get_setting("system_name") or "실패비용(COPQ) 관리"
COMPANY = db.get_setting("company") or ""


def _switch_dataset():
    for k in ("year", "month"):
        st.session_state.pop(k, None)


with st.sidebar:
    st.markdown(f"### 📉 {SYS}")
    if COMPANY:
        st.caption(COMPANY)
    st.selectbox("데이터셋 (사업장·제품)", _ds, key="dataset", on_change=_switch_dataset,
                 help="data 폴더의 .db 파일 하나가 데이터셋 하나입니다. 새 데이터셋은 데이터 관리 › 데이터셋에서 만듭니다.")
    ys = years_available()
    if "year" not in st.session_state or st.session_state.year not in ys:
        st.session_state.year = max(ys) if not db.is_empty() else date.today().year
        st.session_state.month = latest_month(st.session_state.year)
    st.selectbox("분석 연도", ys, key="year")
    st.selectbox("기준월", list(range(1, 13)), key="month", format_func=lambda m: f"{m}월")
    st.caption("기준월을 바꾸면 대시보드·파레토·클레임 현황이 함께 바뀝니다.")

YEAR, MONTH = int(st.session_state.year), int(st.session_state.month)


def need_data():
    if db.is_empty():
        page_setup()
        st.stop()


def page_setup():
    """빈 데이터셋 초기 설정 마법사."""
    st.title("시작하기")
    st.write(f"데이터셋 **{st.session_state.dataset}** 이(가) 비어 있습니다. 시작 방법을 고르세요.")
    t1, t2, t3, t4 = st.tabs(["① 직접 설정", "② 표준 엑셀 양식으로 가져오기", "③ 체험용 샘플 데이터", "④ 기존 ECV VER_9.x 엑셀 이관"])
    with t1:
        with st.form("setup"):
            c1, c2, c3 = st.columns([1.2, 1.2, 0.6])
            comp = c1.text_input("회사·사업장명", COMPANY)
            name = c2.text_input("시스템 이름", SYS)
            cur = c3.text_input("통화 단위", CUR)
            procs = st.text_area("공정·라인 (한 줄에 하나, 순서대로)", "조립1\n조립2\n최종검사", height=130)
            wk = st.radio("주차 기준", ["sun", "mon"], horizontal=True,
                          format_func={"sun": "일요일 시작 (엑셀 WEEKNUM 기본)", "mon": "월요일 시작"}.get)
            if st.form_submit_button("시작", type="primary"):
                names = [x.strip() for x in procs.splitlines() if x.strip()]
                if not names:
                    st.error("공정을 하나 이상 입력하세요.")
                else:
                    for k, v in [("company", comp), ("system_name", name), ("currency", cur), ("week_start", wk)]:
                        db.set_setting(k, v)
                    db.ensure_processes(names)
                    db.log_change("설정", f"초기 설정: 공정 {len(names)}개", "")
                    st.success("설정했습니다. **기준정보**에서 부품단가 → 불량유형 → BOM → 목표 순으로 등록하세요.")
                    st.rerun()
    with t2:
        st.write("공정·부품단가·불량유형·BOM·일일불량·귀책판정·목표·조치내역을 시트별로 작성해 한 번에 올립니다. "
                 "빨간 머리글이 필수 열입니다.")
        st.download_button("📄 표준 양식 받기 (예시 포함)", excel_io.standard_template(), "실패비용_표준양식.xlsx")
        f = st.file_uploader("작성한 표준 양식", type=["xlsx"], key="setup_std")
        if f and st.button("가져오기", type="primary", key="setup_std_btn"):
            stats, warns = excel_io.import_standard(f)
            for w in warns:
                st.warning(w)
            st.success("가져오기 완료: " + ", ".join(f"{k} {v}" for k, v in stats.items()))
            st.button("시작하기")
    with t3:
        st.write("가상의 모터 조립 라인(4개 공정, 13개 불량유형, 6개월 실적, 협력사 클레임 포함) 데이터를 만들어 기능을 둘러봅니다. "
                 "체험 후에는 데이터 관리 › 데이터셋에서 새 데이터셋을 만들어 실제 업무를 시작하세요.")
        if st.button("샘플 데이터 만들기", type="primary"):
            stats = excel_io.generate_sample(date.today().year, 6)
            db.log_change("전체", "샘플 생성", str(stats))
            st.rerun()
    with t4:
        st.write("‘ECV 실패비용 SYSTEM VER_9.x’ 형식 엑셀(부품단가DB·유형별BOM·유형별초기원인·일일불량입력·귀책입력·KPI관리 시트)을 그대로 이관합니다.")
        f = st.file_uploader("VER_9.x 엑셀", type=["xlsx", "xlsm"], key="setup_legacy")
        if f and st.button("이관", type="primary", key="setup_legacy_btn"):
            stats = excel_io.import_workbook(f)
            st.success("이관 완료: " + ", ".join(f"{k} {v}" for k, v in stats.items()))
            st.button("시작하기")


# ================================================================== 1. 대시보드
def page_dashboard():
    need_data()
    st.title(f"KPI 대시보드 — {YEAR}년 {MONTH}월 기준")
    agg = engine.daily_agg(YEAR)
    mon = engine.monthly(YEAR, agg)
    lim = engine.kpi_limits(YEAR, MONTH, mon)
    km = engine.kpi_monthly(YEAR, MONTH, mon)
    cs = engine.claim_summary(YEAR, MONTH)

    tot = lim.iloc[0]
    cur = km[km["month"] == MONTH].iloc[0]
    prev = km[km["month"] == MONTH - 1].iloc[0] if MONTH > 1 else None
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("누적 손실비용", won(tot["cum"]), f"연간 한도의 {pct(tot['annual_rate'])}", delta_color="off")
    d = None if prev is None or pd.isna(prev.get("loss_total")) else cur["loss_total"] - prev["loss_total"]
    c2.metric(f"{MONTH}월 손실비용", won(cur.get("loss_total", 0)),
              None if d is None else f"{d:+,.0f}원 (전월비)", delta_color="inverse")
    c3.metric(f"{MONTH}월 목표 대비", pct(cur.get("loss_total", 0) / cur["target"] if cur["target"] else 0),
              f"월간목표 {won(cur['target'])}", delta_color="off")
    c4.metric("클레임 회수율", pct(cs["회수율"]), f"실청구 {won(cs['실청구액'])} / 청구 {won(cs['총청구금액'])}",
              delta_color="off")

    over = lim[(lim["process"] != "Total") & (lim["annual_rate"] >= 1)]
    pace = lim[(lim["process"] != "Total") & (lim["annual_rate"] < 1) & (lim["month_rate"] > 1)]
    if len(over):
        st.error("🚨 **연간 한도 초과**: " + ", ".join(f"{r.process} ({pct(r.annual_rate)})" for r in over.itertuples())
                 + " → 시정조치 필요")
    if len(pace):
        st.warning("⚠️ **월평균이 월간목표 초과(한도 소진 페이스)**: "
                   + ", ".join(f"{r.process} (월목표의 {pct(r.month_rate)})" for r in pace.itertuples()))

    left, right = st.columns([1.25, 1])
    with left:
        st.subheader("실패비용 한도관리")
        t = lim.copy()
        show(t[["process", "annual", "cum", "annual_rate", "remain", "monthly", "avg_month", "month_rate"]],
             rename={"process": "공정", "annual": "연간 목표", "cum": "누적 실적", "annual_rate": "연간 사용률",
                     "remain": "잔여 한도", "monthly": "월간목표", "avg_month": "월평균 손실", "month_rate": "월목표 대비"})
        st.caption("월목표 대비 = 월평균 손실 ÷ 월간목표 (100% 초과 시 연간 한도 초과 페이스)")
    with right:
        st.subheader("공정별 한도 사용률")
        t = lim[lim["process"] != "Total"]
        fig = go.Figure(go.Bar(x=t["annual_rate"] * 100, y=t["process"], orientation="h",
                               marker_color=[C_TGT if v >= 1 else (C_CLAIM if v >= MONTH / 12 else C_CONF) for v in t["annual_rate"]],
                               text=[pct(v) for v in t["annual_rate"]], textposition="outside"))
        fig.add_vline(x=100, line_dash="dash", line_color=C_TGT, annotation_text="한도", annotation_position="top")
        fig.add_vline(x=MONTH / 12 * 100, line_dash="dot", line_color="#666", annotation_text=f"{MONTH}월 경과 기준", annotation_position="bottom")
        fig.update_layout(height=280, margin=dict(l=10, r=40, t=30, b=10), xaxis_title="연간 한도 사용률(%)",
                          yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("월별 실패비용 추이")
    k = km[km["month"] <= MONTH]
    fig = go.Figure()
    fig.add_bar(x=k["month"].map(lambda m: f"{m}월"), y=k["loss_confirmed"], name="손실(확정)", marker_color=C_CONF)
    fig.add_bar(x=k["month"].map(lambda m: f"{m}월"), y=k["loss_potential"], name="손실(보류·잠재)", marker_color=C_POT)
    fig.add_scatter(x=k["month"].map(lambda m: f"{m}월"), y=k["target"], name="월간목표", mode="lines",
                    line=dict(color=C_TGT, dash="dash"))
    fig.add_scatter(x=k["month"].map(lambda m: f"{m}월"), y=k["claim_amt"], name="클레임 청구", mode="lines+markers",
                    line=dict(color=C_CLAIM))
    fig.update_layout(barmode="stack", height=340, margin=dict(l=10, r=10, t=10, b=10), yaxis_tickformat=",",
                      legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True)
    with st.expander("월별 세부 현황 표", expanded=False):
        show(km, rename={"month": "월", "est_loss": "손실비용(추정)", "claim_amt": "클레임액", "prevent_amt": "예방비용",
                         "rework_amt": "REWORK", "loss_potential": "손실비용(보류)", "loss_confirmed": "손실비용(확정)",
                         "loss_total": "손실비용(보류+확정)", "target": "월간목표비용", "over": "목표초과",
                         "mom": "전월대비 증감률", "cum": "누적실적"})

    st.subheader("클레임 현황")
    a, b, c, d2 = st.columns(4)
    a.metric("클레임 건수 / NCR", f"{cs['건수']}건 / {cs['NCR수']}건")
    b.metric("총 청구금액", won(cs["총청구금액"]))
    c.metric("실청구액", won(cs["실청구액"]), pct(cs["회수율"]), delta_color="off")
    d2.metric("조정액(미회수 합의분)", won(cs["조정액"]))
    show(engine.claims_by_ncr(YEAR, MONTH),
         rename={"ncr_no": "NCR No.", "claim_vendor": "클레임 업체", "defect_types": "불량유형", "resp_qty": "귀책수량",
                 "claim_amt": "청구금액", "status": "진행상태", "actual_amount": "실청구액", "share": "분담률",
                 "adjust": "조정액", "agreed_date": "협의완료일"})


# ================================================================== 2. 일일 불량 입력
def page_daily():
    need_data()
    st.title("일일 불량 입력")
    tab1, tab2, tab3 = st.tabs(["✏️ 입력", "🔎 조회·수정", "📤 엑셀 일괄 업로드"])
    types = db.read("defect_types", "active=1")
    uc = engine.unit_costs().set_index("defect_type")["unit_cost"]

    with tab1:
        c1, c2 = st.columns([1, 1])
        d = c1.date_input("발생일", value=date.today(), format="YYYY-MM-DD")
        proc = c2.selectbox("공정", db.get_processes())
        st.caption(f"주차: {engine.excel_weeknum(d)}주차 (엑셀 WEEKNUM 기준) · 해당 공정의 불량유형 목록에 수량만 입력하고 저장하세요.")
        show_all = st.toggle("다른 공정의 불량유형도 표시", value=False)
        used = db.query("SELECT defect_type, COUNT(*) n FROM daily WHERE process=? GROUP BY defect_type", (proc,))
        freq = used.set_index("defect_type")["n"]
        t = types if show_all else types[(types["process"] == proc) | types["defect_type"].isin(freq.index)]
        t = t.assign(_n=t["defect_type"].map(freq).fillna(0)).sort_values(["_n", "defect_type"], ascending=[False, True])
        grid = pd.DataFrame({
            "불량유형": t["defect_type"].values,
            "수량": [None] * len(t),
            "remark": pd.Series([""] * len(t), dtype=object),
            "손실단가": t["defect_type"].map(uc).values,
            "비용구분": t["cost_class"].values,
            "원인분류": t["initial_resp"].values,
            "처리방안": t["disposition"].values,
        })
        ed = st.data_editor(grid, hide_index=True, use_container_width=True, height=420, key=f"grid_{proc}_{d}_{show_all}",
                            disabled=["불량유형", "손실단가", "비용구분", "원인분류", "처리방안"],
                            column_config={"수량": st.column_config.NumberColumn("수량", min_value=0, step=1),
                                           "손실단가": st.column_config.NumberColumn(format="localized")})
        rows = ed[pd.to_numeric(ed["수량"], errors="coerce").fillna(0) > 0]
        est = (rows["수량"].astype(float) * rows["손실단가"]).sum()
        st.write(f"입력 {len(rows)}건 · 수량 {int(rows['수량'].sum()) if len(rows) else 0}개 · 추정 손실 **{won(est)}**")
        if st.button("💾 저장", type="primary", disabled=len(rows) == 0):
            db.executemany("INSERT INTO daily(date,process,defect_type,qty,remark,exclude,created_at) VALUES (?,?,?,?,?,0,?)",
                           [(d.isoformat(), proc, r["불량유형"], int(r["수량"]), r["remark"] or None, db.now())
                            for _, r in rows.iterrows()])
            st.success(f"{len(rows)}건 저장했습니다.")
            st.rerun()

    with tab2:
        c1, c2, c3 = st.columns([1, 1, 2])
        d0 = c1.date_input("시작일", value=date(YEAR, MONTH, 1), format="YYYY-MM-DD")
        d1 = c2.date_input("종료일", value=(pd.Timestamp(date(YEAR, MONTH, 1)) + pd.offsets.MonthEnd(0)).date(), format="YYYY-MM-DD")
        pf = c3.multiselect("공정", db.get_processes(), default=[])
        df = engine.daily_enriched()
        if df.empty:
            st.info("데이터가 없습니다.")
        else:
            df = df[(df["date"].dt.date >= d0) & (df["date"].dt.date <= d1)]
            if pf:
                df = df[df["process"].isin(pf)]
            view = df[["id", "date", "week", "process", "defect_type", "qty", "unit_cost", "est_loss", "cost_class",
                       "guess_cause", "disposition", "remark", "exclude"]].copy()
            view["date"] = view["date"].dt.date
            view["exclude"] = view["exclude"].fillna(0).astype(bool)
            view.insert(0, "삭제", False)
            st.caption(f"{len(view)}건 · 수량 {int(view['qty'].sum()):,} · 추정손실 {won(view['est_loss'].sum())}  "
                       "— 날짜·공정·유형·수량·비고·제외를 직접 수정할 수 있습니다. ‘제외’ 체크 시 실패비용 집계에서 빠집니다.")
            ed = st.data_editor(
                view, hide_index=True, use_container_width=True, height=480, key="daily_edit",
                disabled=["id", "week", "unit_cost", "est_loss", "cost_class", "guess_cause", "disposition"],
                column_config={
                    "date": st.column_config.DateColumn("날짜", format="YYYY-MM-DD"),
                    "process": st.column_config.SelectboxColumn("공정", options=db.get_processes()),
                    "defect_type": st.column_config.SelectboxColumn("불량유형", options=types["defect_type"].tolist(), width="large"),
                    "qty": st.column_config.NumberColumn("수량", min_value=0, step=1),
                    "unit_cost": st.column_config.NumberColumn("손실단가", format="localized"),
                    "est_loss": st.column_config.NumberColumn("추정손실", format="localized"),
                    "exclude": st.column_config.CheckboxColumn("제외"),
                    "id": "ID", "week": "주차", "cost_class": "비용구분", "guess_cause": "원인분류",
                    "disposition": "처리방안", "remark": "비고",
                })
            if st.button("변경사항 저장", type="primary"):
                orig = view.set_index("id")
                upd, dels = [], []
                for _, r in ed.iterrows():
                    if r["삭제"]:
                        dels.append(int(r["id"]))
                        continue
                    o = orig.loc[r["id"]]
                    if any(str(r[c]) != str(o[c]) for c in ["date", "process", "defect_type", "qty", "remark", "exclude"]):
                        upd.append((str(r["date"]), r["process"], r["defect_type"], int(r["qty"]),
                                    r["remark"] if r["remark"] else None, int(bool(r["exclude"])), int(r["id"])))
                if upd:
                    db.executemany("UPDATE daily SET date=?,process=?,defect_type=?,qty=?,remark=?,exclude=? WHERE id=?", upd)
                if dels:
                    db.executemany("DELETE FROM daily WHERE id=?", [(i,) for i in dels])
                st.success(f"수정 {len(upd)}건, 삭제 {len(dels)}건 반영")
                st.rerun()

    with tab3:
        st.download_button("📄 업로드 양식 받기", excel_io.daily_template(), "일일불량_업로드양식.xlsx")
        f = st.file_uploader("작성한 양식 업로드 (열: 날짜, 공정, 불량유형, 수량, remark)", type=["xlsx"])
        if f:
            up = pd.read_excel(f).dropna(subset=["날짜", "불량유형", "수량"])
            up["날짜"] = pd.to_datetime(up["날짜"]).dt.date
            unknown = sorted(set(up["불량유형"]) - set(types["defect_type"]))
            st.dataframe(up, hide_index=True, use_container_width=True)
            if unknown:
                st.error("등록되지 않은 불량유형: " + ", ".join(unknown) + " → 기준정보에서 먼저 등록하세요.")
            elif st.button(f"{len(up)}건 등록", type="primary"):
                db.executemany("INSERT INTO daily(date,process,defect_type,qty,remark,exclude,created_at) VALUES (?,?,?,?,?,0,?)",
                               [(r["날짜"].isoformat(), r["공정"], r["불량유형"], int(r["수량"]),
                                 None if pd.isna(r.get("remark")) else str(r.get("remark")), db.now())
                                for _, r in up.iterrows()])
                st.success("등록 완료")


# ================================================================== 3. 귀책 판정
JCOLS = ["confirmed_cause", "responsibility", "retest_ng", "retest_ok", "reuse_flag", "ncr_no", "claim_yn",
         "claim_vendor", "resp_qty", "claim_no", "actual_amount", "agreed_date", "memo"]


def page_judgment():
    need_data()
    st.title("귀책 판정 · 재검 결과 입력")
    st.caption("주차·공정·불량유형 단위로 집계된 불량에 대해 확정원인, 귀책(사내/공급), 재검 결과, 재사용 여부, 클레임 정보를 입력합니다. "
               "재검합격 수량은 ‘재사용’ 선택 시 REWORK, ‘보류’ 선택 시 손실(잠재)로 집계됩니다.")
    agg = engine.daily_agg(YEAR)
    if agg.empty:
        st.info("일일 불량 데이터가 없습니다.")
        return
    c1, c2, c3 = st.columns([1, 1.4, 1.6])
    months = c1.multiselect("월", list(range(1, 13)), default=[MONTH])
    mode = c2.radio("표시 대상", ["판정 대상", "판정 완료", "전체"], horizontal=True)
    pf = c3.multiselect("공정", db.get_processes())
    g = agg[agg["month"].isin(months)] if months else agg
    if pf:
        g = g[g["process"].isin(pf)]
    j = engine.judgments(YEAR).set_index("key")
    g = g[["key", "month", "week", "process", "defect_type", "disposition", "guess_cause", "qty", "unit_cost"]].copy()
    for c in JCOLS:
        g[c] = g["key"].map(j[c]) if c in j else None
    judged = g["key"].isin(j.index)
    target = g["disposition"].isin(db.options("judge_dispositions")) | g["guess_cause"].isin(db.options("judge_causes"))
    if mode == "판정 대상":
        g = g[target | judged]
    elif mode == "판정 완료":
        g = g[judged]
    for c in ["retest_ng", "retest_ok", "resp_qty"]:
        g[c] = pd.to_numeric(g[c], errors="coerce").fillna(0).astype(int)
    g["claim_yn"] = g["claim_yn"].fillna("") == "Y"
    g["agreed_date"] = pd.to_datetime(g["agreed_date"], errors="coerce").dt.date
    for c in ["confirmed_cause", "responsibility", "reuse_flag", "ncr_no", "claim_vendor", "claim_no", "memo"]:
        g[c] = pd.Series([("" if pd.isna(v) else str(v)) for v in g[c]], index=g.index, dtype=object)
    vendors = sorted(set(db.query("SELECT DISTINCT claim_vendor v FROM judgments WHERE claim_vendor IS NOT NULL")["v"]))
    st.write(f"{len(g)}건 · 판정완료 {int(g['key'].isin(j.index).sum())}건")
    ed = st.data_editor(
        g, hide_index=True, use_container_width=True, height=520, key=f"jud_{months}_{mode}_{pf}",
        disabled=["key", "month", "week", "process", "defect_type", "disposition", "guess_cause", "qty", "unit_cost"],
        column_order=["month", "week", "process", "defect_type", "disposition", "guess_cause", "qty", "unit_cost"] + JCOLS,
        column_config={
            "month": "월", "week": "주차", "process": "공정",
            "defect_type": st.column_config.TextColumn("불량유형", width="medium"),
            "disposition": "처리방안", "guess_cause": "원인분류", "qty": "수량",
            "unit_cost": st.column_config.NumberColumn("단가", format="localized"),
            "confirmed_cause": st.column_config.TextColumn("확정원인", help="자유 입력. 예: " + ", ".join(db.options("opt_confirmed_cause"))),
            "responsibility": st.column_config.TextColumn("귀책", help="사내 또는 공급", validate=r"^(사내|공급)?$"),
            "retest_ng": st.column_config.NumberColumn("재검불량", min_value=0, step=1),
            "retest_ok": st.column_config.NumberColumn("재검합격", min_value=0, step=1),
            "reuse_flag": st.column_config.TextColumn("재사용여부", help="재사용 → REWORK / 보류 → 손실(잠재)", validate=r"^(재사용|보류)?$"),
            "ncr_no": "NCR No.",
            "claim_yn": st.column_config.CheckboxColumn("클레임"),
            "claim_vendor": st.column_config.TextColumn("클레임 업체", help="기존 업체: " + ", ".join(vendors)),
            "resp_qty": st.column_config.NumberColumn("귀책수량", min_value=0, step=1),
            "claim_no": "클레임 No.",
            "actual_amount": st.column_config.NumberColumn("실청구액", format="localized"),
            "agreed_date": st.column_config.DateColumn("협의완료일", format="YYYY-MM-DD"),
            "memo": "메모",
        })
    st.caption("귀책: 사내/공급 · 재사용여부: 재사용/보류 (그 외 값은 입력 오류로 표시) · 칸을 모두 비우고 저장하면 해당 판정이 삭제됩니다.")
    if st.button("💾 판정 저장", type="primary"):
        n_up, n_del = 0, 0
        for _, r in ed.iterrows():
            vals = {c: r[c] for c in JCOLS}
            vals["claim_yn"] = "Y" if r["claim_yn"] else None
            vals["agreed_date"] = None if pd.isna(vals["agreed_date"]) or vals["agreed_date"] in ("", None) else str(vals["agreed_date"])
            vals["actual_amount"] = None if pd.isna(vals["actual_amount"]) else float(vals["actual_amount"])
            filled = any(v not in (None, "", 0) and not (isinstance(v, float) and np.isnan(v)) for v in vals.values())
            exists = r["key"] in j.index
            if not filled:
                if exists:
                    db.execute("DELETE FROM judgments WHERE key=?", (r["key"],))
                    n_del += 1
                continue
            if exists:
                old = j.loc[r["key"]]
                same = all(str(vals[c] if vals[c] not in ("", None) else None) == str(old[c] if not pd.isna(old[c]) and old[c] != "" else None)
                           for c in JCOLS if c not in ("retest_ng", "retest_ok", "resp_qty", "actual_amount")) and \
                    all(int(vals[c] or 0) == int(old[c] or 0) for c in ("retest_ng", "retest_ok", "resp_qty")) and \
                    (vals["actual_amount"] or 0) == (0 if pd.isna(old["actual_amount"]) else old["actual_amount"])
                if same:
                    continue
            db.execute(f"""INSERT INTO judgments(key,year,month,week,process,defect_type,{','.join(JCOLS)},updated_at)
                       VALUES (?,?,?,?,?,?,{','.join('?' * len(JCOLS))},?)
                       ON CONFLICT(key) DO UPDATE SET {','.join(f'{c}=excluded.{c}' for c in JCOLS)},updated_at=excluded.updated_at""",
                       (r["key"], YEAR, int(r["month"]), int(r["week"]), r["process"], r["defect_type"],
                        *[None if vals[c] in ("",) else vals[c] for c in JCOLS], db.now()))
            n_up += 1
        st.success(f"저장 {n_up}건, 판정 삭제 {n_del}건")
        st.rerun()

    with st.expander("집계 결과 확인 (해당 월 일일집계 + COPQ 분리)"):
        a = agg[agg["month"].isin(months)] if months else agg
        show(a[["key", "qty", "resp_qty", "retest_ng", "reuse_qty", "hold_qty", "scrap_amt", "rework_amt",
                "hold_amt", "claim_amt", "prevent_amt", "loss_total", "match"]])


# ================================================================== 4. 클레임
def page_claims():
    need_data()
    st.title("클레임 관리")
    c = engine.claims(YEAR)
    cs = engine.claim_summary(YEAR, 12)
    a, b, cc, d = st.columns(4)
    a.metric("총 청구금액(연)", won(cs["총청구금액"]))
    b.metric("실청구액", won(cs["실청구액"]), pct(cs["회수율"]), delta_color="off")
    cc.metric("청구대기 금액", won(c.loc[c["status"] == "청구대기", "claim_amt"].sum()))
    d.metric("조정액", won(cs["조정액"]))

    tab1, tab2, tab3 = st.tabs(["📒 클레임 원장", "🧾 NCR 일괄 청구 처리", "🏭 업체별 현황"])
    with tab1:
        f1, f2, f3 = st.columns(3)
        sf = f1.multiselect("진행상태", ["청구대기", "청구완료"])
        vf = f2.multiselect("업체", sorted(c["claim_vendor"].unique()))
        mf = f3.multiselect("발생월", sorted(c["month"].unique()))
        x = c.copy()
        if sf: x = x[x["status"].isin(sf)]
        if vf: x = x[x["claim_vendor"].isin(vf)]
        if mf: x = x[x["month"].isin(mf)]
        show(x.rename(columns={"month": "발생월"}), height=420)
        st.caption("청구금액 = 손실단가 × 귀책수량 · 클레임 No.가 입력되면 ‘청구완료’ · 분담률 = 실청구액 ÷ 청구금액 · 조정액 = 청구금액 − 실청구액")
        st.download_button("CSV로 받기", x.rename(columns=excel_io.KOR).to_csv(index=False).encode("utf-8-sig"),
                           f"클레임원장_{YEAR}.csv", "text/csv")

    with tab2:
        st.write("NCR 하나에 묶인 귀책 건 전체에 클레임 No.·분담률·협의완료일을 한 번에 적용합니다. (실청구액 = 청구금액 × 분담률)")
        j = engine.judgments(YEAR)
        j = j[(j["claim_yn"] == "Y") & (j["ncr_no"].fillna("") != "")]
        if j.empty:
            st.info("NCR이 입력된 클레임 건이 없습니다.")
        else:
            ncr = st.selectbox("NCR No.", sorted(j["ncr_no"].unique()))
            uc = engine.unit_costs().set_index("defect_type")["unit_cost"]
            t = j[j["ncr_no"] == ncr].copy()
            t["unit_cost"] = t["defect_type"].map(uc)
            t["claim_amt"] = t["resp_qty"] * t["unit_cost"]
            show(t[["key", "claim_vendor", "defect_type", "resp_qty", "unit_cost", "claim_amt", "claim_no",
                    "actual_amount", "agreed_date"]])
            st.write(f"청구금액 합계: **{won(t['claim_amt'].sum())}**")
            c1, c2, c3 = st.columns(3)
            cno = c1.text_input("클레임 No.", value=(t["claim_no"].dropna().iloc[0] if t["claim_no"].notna().any() else f"{YEAR}-{MONTH:02d}"))
            rate = c2.number_input("분담률(%)", 0.0, 100.0, 100.0, 5.0)
            ad = c3.date_input("협의완료일", value=date.today(), format="YYYY-MM-DD")
            st.write(f"→ 실청구액 합계 **{won(t['claim_amt'].sum() * rate / 100)}**, 조정액 {won(t['claim_amt'].sum() * (1 - rate / 100))}")
            if st.button("적용", type="primary"):
                db.executemany("UPDATE judgments SET claim_no=?, actual_amount=?, agreed_date=?, updated_at=? WHERE key=?",
                               [(cno, float(r.claim_amt) * rate / 100, ad.isoformat(), db.now(), r.key) for r in t.itertuples()])
                db.log_change("클레임", f"NCR {ncr} 청구처리: {cno}, 분담률 {rate}%", "클레임 협의 결과")
                st.success("반영했습니다.")
                st.rerun()

    with tab3:
        if c.empty:
            st.info("데이터 없음")
        else:
            v = c.groupby("claim_vendor", as_index=False).agg(건수=("ncr_no", "count"), NCR=("ncr_no", "nunique"),
                                                              귀책수량=("resp_qty", "sum"), 청구금액=("claim_amt", "sum"),
                                                              실청구액=("actual_amount", "sum"), 조정액=("adjust", "sum"))
            v["회수율"] = np.where(v["청구금액"] > 0, v["실청구액"] / v["청구금액"], 0)
            show(v.rename(columns={"claim_vendor": "업체"}))
            fig = go.Figure()
            fig.add_bar(x=v["claim_vendor"], y=v["청구금액"], name="청구금액", marker_color=C_CLAIM)
            fig.add_bar(x=v["claim_vendor"], y=v["실청구액"], name="실청구액", marker_color=C_CONF)
            fig.update_layout(barmode="group", height=320, yaxis_tickformat=",", margin=dict(t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)


# ================================================================== 5. 월간 파레토
def page_pareto():
    need_data()
    st.title(f"월간 파레토 — {YEAR}년 {MONTH}월")
    agg = engine.daily_agg(YEAR)
    mon = engine.monthly(YEAR, agg)
    tg = engine.targets(YEAR).set_index("process")
    x = mon[mon["month"] == MONTH]
    if x.empty:
        st.info("해당 월 데이터가 없습니다.")
        return
    s = x.groupby("process").agg(qty=("qty", "sum"), est=("est_loss", "sum"), pot=("loss_potential", "sum"),
                                 conf=("loss_confirmed", "sum"), tot=("loss_total", "sum"),
                                 claim=("claim_amt", "sum"), prev=("prevent_amt", "sum"), rew=("rework_amt", "sum"))
    s = s.reindex(list(dict.fromkeys(db.get_processes() + list(s.index)))).fillna(0)
    s.insert(0, "target", tg["monthly"])
    s["rate"] = np.where(s["target"] > 0, s["tot"] / s["target"], 0)
    s.loc["합계"] = s.sum(numeric_only=True)
    s.loc["합계", "rate"] = s.loc["합계", "tot"] / s.loc["합계", "target"] if s.loc["합계", "target"] else 0
    a, b, c = st.columns(3)
    a.metric("월간 목표금액", won(s.loc["합계", "target"]))
    b.metric("월간 총 손실비용", won(s.loc["합계", "tot"]))
    c.metric("목표 대비", pct(s.loc["합계", "rate"]), delta_color="off")
    show(s.reset_index().rename(columns={"index": "공정"}),
         rename={"process": "공정", "target": "월간 목표", "qty": "부적합수량", "est": "추정손실", "pot": "손실비용(보류)",
                 "conf": "손실비용(확정)", "tot": "손실비용 합계", "claim": "클레임청구", "prev": "예방비용",
                 "rew": "REWORK(금액)", "rate": "목표 대비율"})

    c1, c2, c3 = st.columns([1, 1, 1])
    metric = c1.selectbox("파레토 기준", ["loss_total", "est_loss", "qty"],
                          format_func={"loss_total": "손실비용(보류+확정)", "est_loss": "손실비용(추정)", "qty": "불량수량"}.get)
    proc = c2.selectbox("공정", ["전체"] + db.get_processes())
    topn = c3.slider("표시 항목 수", 5, 30, 15)
    p = engine.pareto(mon, MONTH, metric, None if proc == "전체" else proc)
    p = p[p[metric] > 0]
    top = p.head(topn)
    fig = go.Figure()
    fig.add_bar(x=top["defect_type"], y=top[metric], name=excel_io.KOR.get(metric, metric), marker_color=C_CONF,
                text=[f"{v:,.0f}" for v in top[metric]], textposition="outside", hovertext=top["process"])
    fig.add_scatter(x=top["defect_type"], y=top["cum_pct"] * 100, name="누적 %", yaxis="y2", mode="lines+markers",
                    line=dict(color=C_CLAIM))
    fig.add_hline(y=80, yref="y2", line_dash="dot", line_color="#999")
    fig.update_layout(height=430, margin=dict(t=20, b=10), yaxis=dict(tickformat=","),
                      yaxis2=dict(overlaying="y", side="right", range=[0, 105], ticksuffix="%", dtick=20, showgrid=False),
                      legend=dict(orientation="h", y=1.08), xaxis_tickangle=-30)
    st.plotly_chart(fig, use_container_width=True)
    st.subheader("월간 손실비용 상세 (부적합 실패비용 순)")
    show(p[["process", "defect_type", "qty", "est_loss", "loss_total", "claim_amt", "prevent_amt", "rework_amt",
            "loss_potential", "loss_confirmed", "cum_pct", "check"]],
         rename={**excel_io.KOR, "cum_pct": "누적비율"}, percent=("누적비율",))

    st.subheader("월간 협력사 클레임 발행 비용")
    cl = engine.claims(YEAR)
    cl = cl[(cl["month"] == MONTH) & (cl["claim_amt"] > 0)].sort_values("claim_amt", ascending=False)
    if cl.empty:
        st.caption("해당 월 클레임 없음")
    else:
        show(cl[["process", "ncr_no", "claim_vendor", "defect_type", "resp_qty", "unit_cost", "claim_amt", "status"]])

    acts = db.read("actions", "year=? AND (month=? OR month IS NULL) AND status!='완료'", (YEAR, MONTH))
    if len(acts):
        st.subheader("진행 중 조치")
        show(acts[["defect_type", "analysis", "result", "status", "owner", "due_date"]],
             rename={"defect_type": "대상 불량유형", "analysis": "원인 분석", "result": "확인 결과/조치",
                     "status": "진행상태", "owner": "담당", "due_date": "완료예정"})


# ================================================================== 6. 조치내역
def page_actions():
    need_data()
    st.title("월간 조치내역")
    a = db.read("actions", "year=?", (YEAR,))
    a = a[["id", "month", "defect_type", "analysis", "result", "status", "owner", "due_date"]].copy()
    a["due_date"] = pd.to_datetime(a["due_date"], errors="coerce").dt.date
    a.insert(0, "삭제", False)
    st.caption("행 추가는 표 하단의 ‘+’ 버튼. 진행상태가 ‘완료’가 아닌 항목은 월간 파레토 화면 하단에 표시됩니다.")
    ed = st.data_editor(
        a, hide_index=True, use_container_width=True, num_rows="dynamic", height=500, key="act_edit",
        disabled=["id"],
        column_config={
            "id": "No.", "month": st.column_config.NumberColumn("월", min_value=1, max_value=12, step=1),
            "defect_type": st.column_config.TextColumn("대상 불량유형", width="medium"),
            "analysis": st.column_config.TextColumn("원인 분석 내용", width="large"),
            "result": st.column_config.TextColumn("확인 결과 / 조치 사항", width="large"),
            "status": st.column_config.SelectboxColumn("진행 상태", options=db.options("opt_action_status")),
            "owner": "담당", "due_date": st.column_config.DateColumn("완료예정", format="YYYY-MM-DD"),
        })
    if st.button("💾 저장", type="primary"):
        keep = ed[~ed["삭제"].fillna(False)].drop(columns=["삭제"])
        keep = keep[keep["defect_type"].fillna("") != ""]
        db.execute("DELETE FROM actions WHERE year=?", (YEAR,))
        db.executemany("INSERT INTO actions(year,month,defect_type,analysis,result,status,owner,due_date,updated_at) VALUES (?,?,?,?,?,?,?,?,?)",
                       [(YEAR, None if pd.isna(r.month) else int(r.month), r.defect_type, r.analysis, r.result, r.status,
                         r.owner, None if pd.isna(r.due_date) else str(r.due_date), db.now()) for r in keep.itertuples()])
        st.success("저장했습니다.")
        st.rerun()


# ================================================================== 7. 기준정보
def page_master():
    need_data()
    st.title("기준정보 관리")
    t0, t1, t2, t3, t4, t5, t6 = st.tabs(["🏭 공정·라인", "💲 부품단가", "🧩 유형별 BOM", "🏷️ 불량유형",
                                          "🧮 손실단가(자동)", "🎯 목표·한도", "⚙️ 설정"])

    with t0:
        pr = db.read("processes").sort_values(["sort", "name"])
        pr["active"] = pr["active"].fillna(1).astype(bool)
        pr["note"] = pd.Series([("" if pd.isna(v) else str(v)) for v in pr["note"]], index=pr.index, dtype=object)
        st.caption("공정·라인 목록. ‘순서’대로 화면과 보고서에 표시됩니다. 실적이 있는 공정은 삭제 대신 ‘사용’을 해제하세요.")
        ed = st.data_editor(pr, hide_index=True, num_rows="dynamic", key="proc_ed", use_container_width=True,
                            column_config={"name": st.column_config.TextColumn("공정명", required=True),
                                           "sort": st.column_config.NumberColumn("순서", step=1),
                                           "active": st.column_config.CheckboxColumn("사용"), "note": "비고"})
        if st.button("💾 공정 저장", type="primary"):
            ed = ed[ed["name"].fillna("").str.strip() != ""].copy()
            ed["name"] = ed["name"].str.strip()
            ed["sort"] = pd.to_numeric(ed["sort"], errors="coerce").fillna(99).astype(int)
            ed["active"] = ed["active"].fillna(True).astype(int)
            used = set(db.query("SELECT DISTINCT process FROM daily UNION SELECT DISTINCT process FROM defect_types")["process"].dropna())
            gone = (set(pr["name"]) - set(ed["name"])) & used
            if gone:
                st.error("실적·불량유형에 쓰이는 공정은 삭제할 수 없습니다(‘사용’ 해제): " + ", ".join(sorted(gone)))
            elif ed["name"].duplicated().any():
                st.error("공정명이 중복되었습니다.")
            else:
                db.replace_table("processes", ed[["name", "sort", "active", "note"]])
                db.log_change("공정", f"{len(ed)}개 공정 저장", "")
                st.success("저장했습니다.")
                st.rerun()

    with t1:
        p = db.read("parts").sort_values("id")
        p["changed_at"] = pd.to_datetime(p["changed_at"], errors="coerce").dt.date
        p["note"] = pd.Series([("" if pd.isna(v) else str(v)) for v in p["note"]], index=p.index, dtype=object)
        st.caption("단가를 바꾸면 BOM을 통해 모든 불량유형의 손실단가와 과거 손실비용이 다시 계산됩니다(엑셀과 동일). 변경 내역은 이력으로 남습니다.")
        ed = st.data_editor(p, hide_index=True, use_container_width=True, num_rows="dynamic", key="parts_ed",
                            disabled=["id"],
                            column_config={"id": "ID", "name": "부품명",
                                           "price": st.column_config.NumberColumn("부품단가", format="%.3f", min_value=0),
                                           "changed_at": st.column_config.DateColumn("변경일", format="YYYY-MM-DD"),
                                           "note": "비고"})
        reason = st.text_input("변경 사유 (이력 기록용)", key="price_reason")
        if st.button("💾 단가 저장", type="primary"):
            old = p.set_index("name")["price"]
            hist = []
            ed = ed[ed["name"].fillna("") != ""].copy()
            for i, r in ed.iterrows():
                if r["name"] in old and abs(float(old[r["name"]]) - float(r["price"])) > 1e-9:
                    hist.append((r["name"], float(old[r["name"]]), float(r["price"]), db.now(), reason))
                    if pd.isna(r["changed_at"]) or r["changed_at"] == p.set_index("name").loc[r["name"], "changed_at"]:
                        ed.at[i, "changed_at"] = date.today()
            ed["changed_at"] = ed["changed_at"].map(lambda v: None if pd.isna(v) else str(v))
            removed = set(old.index) - set(ed["name"])
            with db.connect() as con:
                for r in ed.itertuples():
                    if pd.isna(r.id):
                        con.execute("INSERT INTO parts(name,price,changed_at,note) VALUES (?,?,?,?)", (r.name, r.price, r.changed_at, r.note))
                    else:
                        con.execute("UPDATE parts SET name=?,price=?,changed_at=?,note=? WHERE id=?", (r.name, r.price, r.changed_at, r.note, int(r.id)))
                for n in removed:
                    con.execute("DELETE FROM parts WHERE name=?", (n,))
                    con.execute("DELETE FROM bom WHERE part_name=?", (n,))
                con.executemany("INSERT INTO price_history(part_name,old_price,new_price,changed_at,memo) VALUES (?,?,?,?,?)", hist)
            for h in hist:
                db.log_change("부품단가DB", f"{h[0]}: {h[1]:,.3f} → {h[2]:,.3f}", reason)
            st.success(f"저장 (단가 변경 {len(hist)}건)")
            st.rerun()
        ph = db.read("price_history").sort_values("id", ascending=False)
        if len(ph):
            with st.expander("단가 변경 이력"):
                show(ph.drop(columns=["id"]), rename={"part_name": "부품", "old_price": "변경 전 단가", "new_price": "변경 후 단가",
                                                        "changed_at": "변경일시", "memo": "사유"})

    with t2:
        m = engine.bom_matrix()
        st.caption("불량 발생 시 폐기되는 부품의 사용수량(1/0). 신규 불량유형은 ‘불량유형’ 탭에서 먼저 등록하면 여기에 행이 생깁니다.")
        pf = st.selectbox("공정 필터", ["전체"] + sorted(m["process"].dropna().unique().tolist()), key="bom_pf")
        mv = m if pf == "전체" else m[m["process"] == pf]
        uc = engine.unit_costs().set_index("defect_type")["unit_cost"]
        mv = mv.copy()
        mv.insert(2, "손실단가", mv["defect_type"].map(uc))
        ed = st.data_editor(mv, hide_index=True, use_container_width=True, height=520, key=f"bom_ed_{pf}",
                            disabled=["process", "defect_type", "손실단가"],
                            column_config={"process": "공정", "defect_type": st.column_config.TextColumn("불량유형", width="medium"),
                                           "손실단가": st.column_config.NumberColumn(format="localized")})
        if st.button("💾 BOM 저장", type="primary"):
            parts = [c for c in ed.columns if c not in ("process", "defect_type", "손실단가")]
            long = ed.melt(id_vars=["defect_type"], value_vars=parts, var_name="part_name", value_name="qty")
            long = long[pd.to_numeric(long["qty"], errors="coerce").fillna(0) > 0]
            with db.connect() as con:
                con.executemany("DELETE FROM bom WHERE defect_type=?", [(t,) for t in ed["defect_type"]])
                con.executemany("INSERT INTO bom(defect_type,part_name,qty) VALUES (?,?,?)",
                                [(r.defect_type, r.part_name, float(r.qty)) for r in long.itertuples()])
            db.log_change("유형별BOM", f"{len(ed)}개 유형 BOM 갱신 ({pf})", "")
            st.success("저장했습니다.")
            st.rerun()

    with t3:
        t = db.read("defect_types")
        st.caption("비용구분: 실패/예방/제외(집계 제외) · 원인분류: 설정 › 선택 목록에서 관리 · 처리방안: 폐기/재검 등. "
                   "일일 입력 시 이 값이 자동 적용됩니다.")
        ed = st.data_editor(
            t, hide_index=True, use_container_width=True, num_rows="dynamic", height=520, key="types_ed",
            column_config={
                "defect_type": st.column_config.TextColumn("불량유형", width="medium", required=True),
                "process": st.column_config.SelectboxColumn("공정", options=db.get_processes(False)),
                "cost_class": st.column_config.SelectboxColumn("비용구분", options=db.COST_CLASSES),
                "initial_cause": "초기원인", "detection": "검출기준",
                "state": st.column_config.SelectboxColumn("불량상태", options=db.options("opt_state")),
                "initial_resp": st.column_config.SelectboxColumn("원인분류", options=db.options("opt_guess_cause")),
                "factor1": "주요인자#1", "factor2": "주요인자#2", "factor3": "주요인자#3",
                "key_proc1": "주요공정#1", "key_proc2": "주요공정#2",
                "disposition": st.column_config.SelectboxColumn("처리방안", options=db.options("opt_disposition")),
                "active": st.column_config.CheckboxColumn("사용"),
            })
        if st.button("💾 불량유형 저장", type="primary"):
            ed = ed[ed["defect_type"].fillna("") != ""].copy()
            ed["active"] = ed["active"].fillna(True).astype(int)
            ed["cost_class"] = ed["cost_class"].fillna("실패")
            ed["disposition"] = ed["disposition"].fillna("폐기")
            used = set(db.query("SELECT DISTINCT defect_type FROM daily")["defect_type"])
            gone = (set(t["defect_type"]) - set(ed["defect_type"])) & used
            if gone:
                st.error("일일불량 데이터가 있는 유형은 삭제할 수 없습니다(‘사용’ 해제로 숨기세요): " + ", ".join(gone))
            else:
                db.replace_table("defect_types", ed)
                db.log_change("불량유형", f"{len(ed)}개 유형 저장", "")
                st.success("저장했습니다.")
                st.rerun()

    with t4:
        u = engine.unit_costs().merge(db.read("defect_types")[["defect_type", "cost_class"]], on="defect_type")
        u["부품수"] = u["defect_type"].map(db.read("bom").groupby("defect_type").size()).fillna(0).astype(int)
        show(u[["process", "defect_type", "cost_class", "부품수", "unit_cost"]], height=400)
        s = u.groupby("process")["unit_cost"].agg(["count", "mean", "max"]).reset_index()
        st.subheader("공정별 소계")
        show(s, rename={"process": "공정", "count": "유형 수", "mean": "평균 실패비용", "max": "최대 실패비용"})

    with t5:
        tg = engine.targets(YEAR)
        st.caption(f"{YEAR}년 공정별 연간 실패비용 한도와 월간 목표. 월간목표 합계가 대시보드의 월간목표선이 됩니다.")
        ed = st.data_editor(tg, hide_index=True, key=f"tg_{YEAR}", disabled=["process"],
                            column_config={"process": "공정",
                                           "annual": st.column_config.NumberColumn("연간 목표(한도)", format="localized"),
                                           "monthly": st.column_config.NumberColumn("월간 목표", format="localized")})
        c1, c2 = st.columns(2)
        st.write(f"연간 합계 {won(ed['annual'].sum())} · 월간 합계 {won(ed['monthly'].sum())} (연간/12 = {won(ed['annual'].sum() / 12)})")
        if c1.button("💾 목표 저장", type="primary"):
            db.executemany("INSERT INTO targets(year,process,annual,monthly) VALUES (?,?,?,?) "
                           "ON CONFLICT(year,process) DO UPDATE SET annual=excluded.annual, monthly=excluded.monthly",
                           [(YEAR, r.process, float(r.annual), float(r.monthly)) for r in ed.itertuples()])
            db.log_change("목표", f"{YEAR}년 목표 저장", "")
            st.success("저장했습니다.")
            st.rerun()
        if c2.button("월간목표 = 연간/12 로 재계산"):
            db.executemany("INSERT INTO targets(year,process,annual,monthly) VALUES (?,?,?,?) "
                           "ON CONFLICT(year,process) DO UPDATE SET annual=excluded.annual, monthly=excluded.monthly",
                           [(YEAR, r.process, float(r.annual), float(r.annual) / 12) for r in ed.itertuples()])
            st.rerun()


    with t6:
        st.subheader("기본 정보")
        with st.form("basic"):
            c1, c2, c3 = st.columns([1.2, 1.2, 0.6])
            comp = c1.text_input("회사·사업장명", COMPANY)
            name = c2.text_input("시스템 이름", SYS)
            cur = c3.text_input("통화 단위", CUR)
            wk0 = db.get_setting("week_start")
            wk = st.radio("주차 기준", ["sun", "mon"], index=0 if wk0 != "mon" else 1, horizontal=True,
                          format_func={"sun": "일요일 시작 (엑셀 WEEKNUM 기본)", "mon": "월요일 시작"}.get,
                          help="귀책 판정은 주차 단위로 저장되므로, 판정 입력 후 바꾸면 기존 판정이 다른 주차와 어긋날 수 있습니다.")
            if st.form_submit_button("저장", type="primary"):
                for k, v in [("company", comp), ("system_name", name), ("currency", cur), ("week_start", wk)]:
                    db.set_setting(k, v)
                if wk != wk0 and db.query("SELECT COUNT(*) n FROM judgments")["n"].iloc[0] > 0:
                    st.warning("주차 기준이 바뀌었습니다. 데이터 관리 › 데이터 검증에서 ‘귀책 키 불일치’를 확인하세요.")
                db.log_change("설정", "기본 정보 변경", "")
                st.rerun()
        st.subheader("선택 목록")
        st.caption("입력 화면의 드롭다운 값입니다. 쉼표로 구분하세요. "
                   "고정 값(계산에 쓰임): 비용구분 = 실패/예방/제외 · 귀책 = 사내/공급 · 재사용여부 = 재사용/보류")
        with st.form("opts"):
            vals = {}
            for k, label in db.SETTING_LABELS.items():
                vals[k] = st.text_input(label, ", ".join(db.options(k)))
            if st.form_submit_button("목록 저장", type="primary"):
                for k, v in vals.items():
                    db.set_setting(k, [x.strip() for x in v.split(",") if x.strip()])
                db.log_change("설정", "선택 목록 변경", "")
                st.success("저장했습니다.")
                st.rerun()


# ================================================================== 8. 데이터 관리
def page_data():
    st.title("데이터 관리")
    t1, t2, t3, t4, t5 = st.tabs(["✅ 데이터 검증", "📊 보고서·내보내기", "📥 가져오기", "🗂️ 데이터셋", "🗄️ 백업·이력"])
    with t1:
        if db.is_empty():
            st.info("데이터 없음")
        else:
            iss = engine.data_issues(YEAR)
            mon = engine.monthly(YEAR)
            bad = mon[mon["check"] != "✅"]
            st.write(f"월불량 검증(클레임+예방+REWORK+잠재+확정 = 추정손실): **{len(mon) - len(bad)}/{len(mon)} 일치**")
            if len(bad):
                show(bad)
            if iss.empty:
                st.success("입력 오류 후보가 없습니다.")
            else:
                st.warning(f"확인이 필요한 항목 {len(iss)}건")
                st.dataframe(iss, hide_index=True, use_container_width=True)
    with t2:
        if not db.is_empty():
            st.write(f"**분석 보고서** — {YEAR}년 {MONTH}월 기준 KPI·파레토·월불량·클레임·일일집계·기준정보")
            st.download_button("📊 분석 보고서(엑셀)", excel_io.export_report(YEAR, MONTH),
                               f"실패비용보고서_{st.session_state.dataset}_{YEAR}{MONTH:02d}.xlsx", type="primary")
            st.write("**전체 데이터(표준 양식)** — 다른 PC·사업장으로 옮기거나, 엑셀에서 일괄 수정 후 다시 가져올 때 사용")
            st.download_button("📦 전체 데이터 내보내기(표준 양식)", excel_io.export_standard(),
                               f"실패비용_전체데이터_{st.session_state.dataset}_{datetime.now():%Y%m%d}.xlsx")
    with t3:
        kind = st.radio("형식", ["표준 양식", "ECV VER_9.x 엑셀"], horizontal=True)
        if kind == "표준 양식":
            st.download_button("📄 빈 표준 양식(예시 포함)", excel_io.standard_template(), "실패비용_표준양식.xlsx")
            mode = st.radio("방식", ["교체 — 양식에 있는 시트의 데이터를 통째로 교체", "추가 — 일일불량·조치내역은 추가, 나머지는 같은 키 갱신"])
            f = st.file_uploader("표준 양식 파일", type=["xlsx"], key="imp_std")
        else:
            st.write("‘ECV 실패비용 SYSTEM VER_9.x’ 엑셀로 현재 데이터셋 전체를 교체합니다.")
            f = st.file_uploader("VER_9.x 엑셀", type=["xlsx", "xlsm"], key="imp_legacy")
        ok = st.checkbox("현재 데이터셋을 변경하는 것에 동의합니다 (자동 백업 후 진행)")
        if f and ok and st.button("가져오기", type="primary"):
            _backup("before_import")
            if kind == "표준 양식":
                stats, warns = excel_io.import_standard(f, replace=mode.startswith("교체"))
                for w in warns:
                    st.warning(w)
                st.success("완료: " + ", ".join(f"{k} {v}" for k, v in stats.items()))
            else:
                stats = excel_io.import_workbook(f)
                st.success("이관 완료: " + ", ".join(f"{k} {v}" for k, v in stats.items()))
                st.session_state.year = stats["기준연도"]
                st.session_state.month = latest_month(stats["기준연도"])
    with t4:
        st.write("사업장·제품·고객사별로 데이터를 분리해 관리합니다. 데이터셋 하나는 data 폴더의 .db 파일 하나입니다.")
        info = []
        for n in db.list_datasets():
            pth = db.dataset_path(n)
            try:
                info.append({"데이터셋": n, "회사·사업장": db.get_setting("company", pth),
                             "공정 수": len(db.get_processes(False, pth)),
                             "일일불량 건수": int(db.query("SELECT COUNT(*) n FROM daily", path=pth)["n"].iloc[0]),
                             "크기(KB)": round(os.path.getsize(pth) / 1024)})
            except Exception:
                info.append({"데이터셋": n})
        st.dataframe(pd.DataFrame(info), hide_index=True, use_container_width=True)
        c1, c2 = st.columns(2)
        with c1.form("new_ds"):
            st.markdown("**새 데이터셋**")
            nm = st.text_input("이름 (영문·숫자·한글, 공백 없이)", placeholder="예: 2공장_모터")
            copy_master = st.checkbox("현재 데이터셋의 기준정보(공정·단가·불량유형·BOM·설정) 복사")
            if st.form_submit_button("만들기", type="primary"):
                nm = nm.strip().replace(" ", "_")
                if not nm or any(ch in nm for ch in '\\/:*?"<>|'):
                    st.error("사용할 수 없는 이름입니다.")
                elif nm in db.list_datasets():
                    st.error("이미 있는 이름입니다.")
                else:
                    src = db.DB_PATH
                    newp = db.dataset_path(nm)
                    db.init_db(newp)
                    if copy_master:
                        for t in ["settings", "processes", "parts", "defect_types", "bom"]:
                            db.replace_table(t, db.read(t, path=src), newp)
                    st.session_state.dataset = nm
                    _switch_dataset()
                    st.rerun()
        with c2.form("del_ds"):
            st.markdown("**데이터셋 삭제**")
            st.caption("삭제 전 data/backups 폴더에 사본을 남깁니다.")
            dn = st.selectbox("대상", [x for x in db.list_datasets()])
            conf = st.text_input("확인을 위해 이름을 다시 입력")
            if st.form_submit_button("삭제"):
                if conf != dn:
                    st.error("이름이 일치하지 않습니다.")
                else:
                    pth = db.dataset_path(dn)
                    shutil.copy(pth, os.path.join(db.backup_dir(), f"{dn}_deleted_{datetime.now():%Y%m%d_%H%M%S}.db"))
                    os.remove(pth)
                    st.session_state.pop("dataset", None)
                    _switch_dataset()
                    st.rerun()
    with t5:
        if os.path.exists(db.DB_PATH):
            with open(db.DB_PATH, "rb") as fh:
                st.download_button("🗄️ 데이터셋 파일(.db) 다운로드", fh.read(), f"{st.session_state.dataset}_{datetime.now():%Y%m%d}.db")
        rf = st.file_uploader("데이터셋 복원 (.db) — 현재 데이터셋을 덮어씀", type=["db"], key="restore")
        if rf and st.button("복원 실행"):
            _backup("before_restore")
            with open(db.DB_PATH, "wb") as fh:
                fh.write(rf.getbuffer())
            db.init_db()
            st.success("복원했습니다.")
            st.rerun()
        st.subheader("변경 이력 (MOC)")
        log = db.read("change_log").sort_values("id", ascending=False)
        st.dataframe(log.rename(columns={"at": "일시", "target": "대상", "content": "내용", "reason": "사유", "who": "작성자"}).drop(columns=["id"]),
                     hide_index=True, use_container_width=True)


def _backup(tag: str) -> None:
    if os.path.exists(db.DB_PATH) and not db.is_empty():
        shutil.copy(db.DB_PATH, os.path.join(db.backup_dir(), f"{st.session_state.dataset}_{tag}_{datetime.now():%Y%m%d_%H%M%S}.db"))


# ------------------------------------------------------------------ 네비게이션
pg = st.navigation({
    "현황": [st.Page(page_dashboard, title="KPI 대시보드", icon="📊", url_path="dashboard", default=True),
             st.Page(page_pareto, title="월간 파레토", icon="📈", url_path="pareto")],
    "입력": [st.Page(page_daily, title="일일 불량 입력", icon="✏️", url_path="daily"),
             st.Page(page_judgment, title="귀책 판정", icon="⚖️", url_path="judgment"),
             st.Page(page_claims, title="클레임 관리", icon="💰", url_path="claims"),
             st.Page(page_actions, title="월간 조치내역", icon="📋", url_path="actions")],
    "관리": [st.Page(page_master, title="기준정보", icon="🛠️", url_path="master"),
             st.Page(page_data, title="데이터 관리", icon="💾", url_path="data")],
})
pg.run()
