"""엑셀 가져오기 / 내보내기

standard_template() / export_standard() / import_standard():
    사업장 무관 '표준 양식' (설정·공정·부품단가·불량유형·BOM·일일불량·귀책판정·목표·조치내역)
    → 새 사업장 도입, 데이터셋 간 이동·백업에 사용
import_workbook():  'ECV 실패비용 SYSTEM VER_9.x' 형식 엑셀 이관 (기존 엑셀 사용처용)
generate_sample():  가상의 체험용 데이터 생성
export_report():    분석 결과 엑셀 보고서
"""
from __future__ import annotations

import io
import random
from datetime import date, datetime, timedelta

import openpyxl
import pandas as pd

import db
import engine


def _header_map(ws, row=1):
    return {str(ws.cell(row, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(row, c).value is not None}


def _d(v):
    if v is None or v == "":
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return str(v)


def _s(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def _i(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 0


def import_workbook(src, replace: bool = True, path: str | None = None) -> dict:
    """VER_9.x 엑셀 → DB. src: 파일경로 또는 file-like. 반환: 테이블별 건수."""
    wb = openpyxl.load_workbook(src, data_only=True)
    db.init_db(path)
    stats = {}

    # 1) 부품단가DB
    ws = wb["부품단가DB"]
    h = _header_map(ws)
    parts = []
    for r in range(2, ws.max_row + 1):
        name = _s(ws.cell(r, h["부품명"]).value)
        if not name:
            continue
        parts.append({"id": _i(ws.cell(r, h["ID"]).value) or None, "name": name,
                      "price": float(ws.cell(r, h["부품단가"]).value or 0),
                      "changed_at": _d(ws.cell(r, h.get("비고", 4)).value), "note": None})
    pdf = pd.DataFrame(parts)

    # 2) 유형별BOM
    ws = wb["유형별BOM"]
    h = _header_map(ws)
    part_cols = [k for k in h if k not in ("공정", "부적합유형")]
    bom, bom_proc = [], {}
    for r in range(2, ws.max_row + 1):
        t = _s(ws.cell(r, h["부적합유형"]).value)
        if not t:
            continue
        bom_proc[t] = _s(ws.cell(r, h["공정"]).value)
        for p in part_cols:
            q = ws.cell(r, h[p]).value
            if q not in (None, 0, "0", ""):
                bom.append({"defect_type": t, "part_name": p, "qty": float(q)})
    bdf = pd.DataFrame(bom)

    # 3) 유형별초기원인 → defect_types
    ws = wb["유형별초기원인"]
    h = _header_map(ws)
    info = {}
    for r in range(2, ws.max_row + 1):
        t = _s(ws.cell(r, h["불량유형"]).value)
        if not t:
            continue
        g = lambda k: _s(ws.cell(r, h[k]).value) if k in h else None
        info[t] = dict(cost_class=g("비용구분") or "실패", initial_cause=g("초기원인"), detection=g("검출기준"),
                       state=g("불량상태"), initial_resp=g("초기귀책"), factor1=g("주요인자#1"),
                       factor2=g("주요인자#2"), factor3=g("주요인자#3"), key_proc1=g("주요공정#1"),
                       key_proc2=g("주요공정#2"), disposition=g("처리방안") or "폐기",
                       line=(g("라인") or "").replace(" ", "_"))
    types = []
    for t in dict.fromkeys(list(bom_proc) + list(info)):
        i = info.get(t, {})
        types.append({"defect_type": t, "process": bom_proc.get(t) or i.get("line"),
                      "cost_class": i.get("cost_class", "실패"), "initial_cause": i.get("initial_cause"),
                      "detection": i.get("detection"), "state": i.get("state") or "단품",
                      "initial_resp": i.get("initial_resp") or "협력사", "factor1": i.get("factor1"),
                      "factor2": i.get("factor2"), "factor3": i.get("factor3"), "key_proc1": i.get("key_proc1"),
                      "key_proc2": i.get("key_proc2"), "disposition": i.get("disposition", "폐기"), "active": 1})
    tdf = pd.DataFrame(types)

    # 4) 일일불량입력(매일)
    ws = wb["일일불량입력(매일)"]
    h = _header_map(ws)
    daily = []
    for r in range(2, ws.max_row + 1):
        dt = ws.cell(r, h["날짜"]).value
        t = _s(ws.cell(r, h["불량유형"]).value)
        if not dt or not t:
            continue
        daily.append({"id": _i(ws.cell(r, h["ID"]).value) or None, "date": _d(dt),
                      "process": _s(ws.cell(r, h["공정"]).value), "defect_type": t,
                      "qty": _i(ws.cell(r, h["수량"]).value), "remark": _s(ws.cell(r, h["remark"]).value),
                      "exclude": 0, "created_at": db.now()})
    ddf = pd.DataFrame(daily)
    year = pd.to_datetime(ddf["date"]).dt.year.mode()[0] if len(ddf) else date.today().year

    # 5) 귀책입력
    ws = wb["귀책입력"]
    h = _header_map(ws)
    jud = []
    for r in range(2, ws.max_row + 1):
        m = ws.cell(r, h["월"]).value
        t = _s(ws.cell(r, h["불량유형"]).value)
        if m is None or not t:
            continue
        g = lambda k: ws.cell(r, h[k]).value if k in h else None
        wk, proc = _i(g("주차")), _s(g("공정"))
        actual = g("실청구액")
        jud.append({"key": engine.make_key(year, m, wk, proc, t), "year": int(year), "month": int(m), "week": wk,
                    "process": proc, "defect_type": t, "confirmed_cause": _s(g("확정원인(귀책)")),
                    "responsibility": _s(g("귀책")), "retest_ng": _i(g("재검불량")), "retest_ok": _i(g("재검합격")),
                    "reuse_flag": _s(g("재사용여부")), "ncr_no": _s(g("NCR No.")),
                    "claim_yn": (_s(g("클레임(Y/N)")) or "").upper() or None,
                    "claim_vendor": _s(g("클레임 업체")), "resp_qty": _i(g("귀책수량")),
                    "claim_no": _s(g("클레임 No.")),
                    "actual_amount": float(actual) if isinstance(actual, (int, float)) else None,
                    "agreed_date": _d(g("협의완료일")), "memo": None, "updated_at": db.now()})
    jdf = pd.DataFrame(jud).drop_duplicates("key", keep="last") if jud else pd.DataFrame()

    # 6) 목표 (KPI관리 한도 + 월간파레토 공정별 월간목표)
    procs = set(tdf["process"].dropna()) | set(ddf["process"].dropna() if len(ddf) else [])
    tg = {}
    if "KPI관리" in wb.sheetnames:
        ws = wb["KPI관리"]
        for r in range(1, 60):
            p = _s(ws.cell(r, 1).value)
            if p in procs and isinstance(ws.cell(r, 2).value, (int, float)):
                tg.setdefault(p, {})["annual"] = float(ws.cell(r, 2).value)
    if "월간파레토(월초)" in wb.sheetnames:
        ws = wb["월간파레토(월초)"]
        for r in range(25, 40):
            p = _s(ws.cell(r, 2).value)
            if p in procs and isinstance(ws.cell(r, 5).value, (int, float)):
                tg.setdefault(p, {})["monthly"] = float(ws.cell(r, 5).value)
    tgt = pd.DataFrame([{"year": int(year), "process": p, "annual": v.get("annual", 0.0),
                         "monthly": v.get("monthly", v.get("annual", 0.0) / 12)} for p, v in tg.items()])

    # 7) 월간 조치내역
    act = []
    if "월간 조치내역" in wb.sheetnames:
        ws = wb["월간 조치내역"]
        for r in range(3, ws.max_row + 1):
            t = _s(ws.cell(r, 2).value)
            if not t:
                continue
            mon = _s(ws.cell(r, 6).value)
            act.append({"year": int(year), "month": _i(mon.replace("월", "")) if mon else None, "defect_type": t,
                        "analysis": _s(ws.cell(r, 3).value), "result": _s(ws.cell(r, 4).value),
                        "status": _s(ws.cell(r, 5).value), "owner": None, "due_date": None, "updated_at": db.now()})
    adf = pd.DataFrame(act)

    if replace:
        db.replace_table("processes", pd.DataFrame(columns=["name"]), path)
        order = list(tg) + list(tdf["process"].dropna()) + (list(ddf["process"].dropna()) if len(ddf) else [])
        db.ensure_processes([x for x in order if x != "단품"] + (["단품"] if "단품" in order else []), path)
        db.execute("UPDATE processes SET active=0 WHERE name='단품'", (), path)  # VER_9 BOM 분류용(실적 공정 아님)
        db.set_setting("week_start", "sun", path)
        db.replace_table("parts", pdf, path)
        db.replace_table("bom", bdf, path)
        db.replace_table("defect_types", tdf, path)
        db.replace_table("daily", ddf, path)
        db.replace_table("judgments", jdf, path)
        if len(tgt):
            db.execute("DELETE FROM targets WHERE year=?", (int(year),), path)
            db.executemany("INSERT INTO targets(year,process,annual,monthly) VALUES (?,?,?,?)",
                           [tuple(r) for r in tgt[["year", "process", "annual", "monthly"]].itertuples(index=False)], path)
        db.replace_table("actions", adf, path)
        db.log_change("전체", "엑셀 이관 (VER_9.x)", "초기 데이터 이관", "", path)

    stats = {"부품": len(pdf), "BOM 매핑": len(bdf), "불량유형": len(tdf), "일일불량": len(ddf),
             "귀책입력": len(jdf), "목표(공정)": len(tgt), "조치내역": len(adf), "기준연도": int(year)}
    return stats


# ------------------------------------------------------------------ 내보내기
KOR = {
    "year": "연", "month": "월", "week": "주차", "process": "공정", "defect_type": "불량유형", "qty": "수량",
    "unit_cost": "손실단가", "cost_class": "비용구분", "state": "불량상태", "guess_cause": "원인분류",
    "disposition": "처리방안", "remark": "비고", "est_loss": "손실비용(추정)", "resp_qty": "귀책수량",
    "retest_ng": "재검불량", "reuse_qty": "재사용수량", "hold_qty": "재사용보류", "pure_scrap": "순수폐기수량",
    "scrap_amt": "폐기금액", "rework_amt": "REWORK금액", "hold_amt": "재사용보류금액", "claim_amt": "클레임금액",
    "prevent_amt": "예방비용", "loss_confirmed": "손실비용(확정)", "loss_potential": "손실비용(잠재)",
    "loss_total": "손실비용(합계)", "check": "검증", "key": "키", "ncr_no": "NCR No.", "claim_vendor": "클레임 업체",
    "claim_no": "클레임 No.", "status": "진행상태", "actual_amount": "실청구액", "share": "분담률",
    "adjust": "조정액", "agreed_date": "협의완료일", "date": "날짜", "id": "ID", "exclude": "제외",
    "confirmed_cause": "확정원인", "responsibility": "귀책", "claim_yn": "클레임(Y/N)", "resp_total": "귀책합계",
    "match": "매칭여부", "annual": "연간목표", "monthly": "월간목표", "cum": "누적실적",
    "annual_rate": "연간 사용률", "remain": "잔여한도", "avg_month": "월평균 손실", "month_rate": "월목표 대비",
    "target": "월간목표", "over": "목표초과", "mom": "전월대비", "defect_types": "불량유형",
}


def export_report(year: int, month: int) -> bytes:
    agg = engine.daily_agg(year)
    mon = engine.monthly(year, agg)
    daily = engine.daily_enriched(year)
    sheets = {
        "KPI_한도관리": engine.kpi_limits(year, month, mon),
        "KPI_월별현황": engine.kpi_monthly(year, month, mon),
        "클레임_NCR별": engine.claims_by_ncr(year, month),
        f"{month}월_파레토": engine.pareto(mon, month),
        "월불량": mon,
        "클레임원장": engine.claims(year),
        "일일집계": agg.drop(columns=["year"], errors="ignore"),
        "일일불량입력": daily.drop(columns=["created_at", "year"], errors="ignore"),
        "실패비용단가": engine.unit_costs(),
        "부품단가DB": db.read("parts"),
        "유형별BOM": engine.bom_matrix(),
    }
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xw:
        wbk = xw.book
        money = wbk.add_format({"num_format": "#,##0"})
        pct = wbk.add_format({"num_format": "0.0%"})
        head = wbk.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1})
        for name, df in sheets.items():
            out = df.rename(columns=KOR)
            out.to_excel(xw, sheet_name=name[:31], index=False)
            ws = xw.sheets[name[:31]]
            for ci, col in enumerate(out.columns):
                ws.write(0, ci, col, head)
                width = min(max(10, len(str(col)) * 2 + 2), 40)
                fmt = None
                if any(k in str(col) for k in ["금액", "비용", "단가", "청구액", "조정액", "목표", "실적", "한도", "손실", "초과"]):
                    fmt = money
                if any(k in str(col) for k in ["률", "대비"]):
                    fmt = pct
                ws.set_column(ci, ci, width, fmt)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(out), 1), max(len(out.columns) - 1, 0))
    return buf.getvalue()


def daily_template() -> bytes:
    """일괄 업로드용 일일불량 양식."""
    t = db.read("defect_types")[["process", "defect_type"]].rename(columns={"process": "공정", "defect_type": "불량유형"})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        pd.DataFrame(columns=["날짜", "공정", "불량유형", "수량", "remark"]).to_excel(xw, sheet_name="일일불량입력", index=False)
        t.to_excel(xw, sheet_name="불량유형목록", index=False)
        xw.sheets["일일불량입력"].set_column(0, 4, 22)
        xw.sheets["불량유형목록"].set_column(0, 1, 36)
    return buf.getvalue()



# ================================================================== 표준 양식
# 시트명: [(엑셀 열 이름, DB 열, 필수여부)]
STD = {
    "공정": ("processes", [("공정명", "name", True), ("순서", "sort", False), ("사용", "active", False), ("비고", "note", False)]),
    "부품단가": ("parts", [("부품명", "name", True), ("단가", "price", True), ("변경일", "changed_at", False), ("비고", "note", False)]),
    "불량유형": ("defect_types", [("불량유형", "defect_type", True), ("공정", "process", True), ("비용구분", "cost_class", False),
                               ("원인분류", "initial_resp", False), ("처리방안", "disposition", False), ("불량상태", "state", False),
                               ("초기원인", "initial_cause", False), ("검출기준", "detection", False),
                               ("주요인자1", "factor1", False), ("주요인자2", "factor2", False), ("주요인자3", "factor3", False),
                               ("주요공정1", "key_proc1", False), ("주요공정2", "key_proc2", False), ("사용", "active", False)]),
    "BOM": ("bom", [("불량유형", "defect_type", True), ("부품명", "part_name", True), ("사용수량", "qty", True)]),
    "일일불량": ("daily", [("날짜", "date", True), ("공정", "process", True), ("불량유형", "defect_type", True),
                       ("수량", "qty", True), ("비고", "remark", False), ("제외", "exclude", False)]),
    "귀책판정": ("judgments", [("연", "year", True), ("월", "month", True), ("주차", "week", True), ("공정", "process", True),
                          ("불량유형", "defect_type", True), ("확정원인", "confirmed_cause", False), ("귀책", "responsibility", False),
                          ("재검불량", "retest_ng", False), ("재검합격", "retest_ok", False), ("재사용여부", "reuse_flag", False),
                          ("NCR No.", "ncr_no", False), ("클레임(Y/N)", "claim_yn", False), ("클레임업체", "claim_vendor", False),
                          ("귀책수량", "resp_qty", False), ("클레임No.", "claim_no", False), ("실청구액", "actual_amount", False),
                          ("협의완료일", "agreed_date", False), ("메모", "memo", False)]),
    "목표": ("targets", [("연도", "year", True), ("공정", "process", True), ("연간목표", "annual", True), ("월간목표", "monthly", False)]),
    "조치내역": ("actions", [("연", "year", True), ("월", "month", False), ("불량유형", "defect_type", True), ("원인분석", "analysis", False),
                         ("조치사항", "result", False), ("진행상태", "status", False), ("담당", "owner", False), ("완료예정", "due_date", False)]),
}

GUIDE = [
    ("시트", "설명", "필수 열"),
    ("설정", "회사명·시스템명·통화단위·주차기준(sun/mon)", "항목, 값"),
    ("공정", "공정/라인 목록. 순서대로 화면에 표시", "공정명"),
    ("부품단가", "실패 시 폐기되는 부품(또는 원가 요소)의 단가", "부품명, 단가"),
    ("불량유형", "불량유형 마스터. 비용구분=실패/예방/제외, 원인분류·처리방안은 설정 목록 값", "불량유형, 공정"),
    ("BOM", "불량유형 1건 발생 시 손실되는 부품과 수량 → 손실단가 = Σ(사용수량×부품단가)", "불량유형, 부품명, 사용수량"),
    ("일일불량", "발생 실적. 제외=1 이면 집계 제외", "날짜, 공정, 불량유형, 수량"),
    ("귀책판정", "연·월·주차·공정·불량유형 단위 판정. 귀책=사내/공급, 재사용여부=재사용/보류", "연, 월, 주차, 공정, 불량유형"),
    ("목표", "공정별 연간 실패비용 한도와 월간 목표(비우면 연간/12)", "연도, 공정, 연간목표"),
    ("조치내역", "개선 조치 이력", "연, 불량유형"),
    ("", "", ""),
    ("규칙", "불량유형·공정·부품명은 시트 간 철자가 같아야 합니다. 일일불량/불량유형에만 있는 공정은 자동 등록됩니다.", ""),
]


def _write_std(xw, frames: dict, settings: dict):
    wbk = xw.book
    head = wbk.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1})
    req = wbk.add_format({"bold": True, "bg_color": "#C00000", "font_color": "white", "border": 1})
    g = pd.DataFrame(GUIDE[1:], columns=GUIDE[0])
    g.to_excel(xw, sheet_name="작성안내", index=False)
    xw.sheets["작성안내"].set_column(0, 0, 12)
    xw.sheets["작성안내"].set_column(1, 1, 90)
    xw.sheets["작성안내"].set_column(2, 2, 36)
    st_df = pd.DataFrame([(k, v if isinstance(v, str) else ", ".join(map(str, v))) for k, v in settings.items()],
                         columns=["항목", "값"])
    st_df.to_excel(xw, sheet_name="설정", index=False)
    xw.sheets["설정"].set_column(0, 0, 24)
    xw.sheets["설정"].set_column(1, 1, 70)
    for sheet, (table, cols) in STD.items():
        df = frames.get(sheet, pd.DataFrame(columns=[c[0] for c in cols]))
        df.to_excel(xw, sheet_name=sheet, index=False)
        ws = xw.sheets[sheet]
        for i, (k, _, r) in enumerate(cols):
            ws.write(0, i, k, req if r else head)
            ws.set_column(i, i, max(12, len(k) * 2 + 4))
        ws.freeze_panes(1, 0)


def standard_template(example: bool = True) -> bytes:
    """빈 표준 양식(예시 몇 줄 포함)."""
    fr = {}
    if example:
        fr["공정"] = pd.DataFrame([["조립1", 1, 1, ""], ["조립2", 2, 1, ""], ["최종검사", 3, 1, ""]],
                                columns=[c[0] for c in STD["공정"][1]])
        fr["부품단가"] = pd.DataFrame([["하우징", 1200, "2026-01-01", ""], ["모터", 8500, "2026-01-01", ""],
                                    ["기어", 650, "2026-01-01", ""]], columns=[c[0] for c in STD["부품단가"][1]])
        fr["불량유형"] = pd.DataFrame([["하우징 찍힘", "조립1", "실패", "협력사", "폐기", "단품"] + [""] * 7 + [1],
                                    ["모터 소음", "최종검사", "실패", "공정", "재검", "완성품"] + [""] * 7 + [1]],
                                   columns=[c[0] for c in STD["불량유형"][1]])
        fr["BOM"] = pd.DataFrame([["하우징 찍힘", "하우징", 1], ["모터 소음", "하우징", 1], ["모터 소음", "모터", 1],
                                 ["모터 소음", "기어", 2]], columns=[c[0] for c in STD["BOM"][1]])
        fr["일일불량"] = pd.DataFrame([["2026-01-05", "조립1", "하우징 찍힘", 3, "", 0]], columns=[c[0] for c in STD["일일불량"][1]])
        fr["목표"] = pd.DataFrame([[2026, "조립1", 6000000, 500000]], columns=[c[0] for c in STD["목표"][1]])
    settings = {"company": "", "system_name": db.DEFAULT_SETTINGS["system_name"], "currency": "원", "week_start": "sun"}
    settings.update({k: v for k, v in db.DEFAULT_SETTINGS.items() if k.startswith(("opt_", "judge_"))})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        _write_std(xw, fr, settings)
    return buf.getvalue()


def export_standard() -> bytes:
    """현재 데이터셋 전체를 표준 양식으로 (백업·다른 PC 이동용)."""
    fr = {}
    for sheet, (table, cols) in STD.items():
        df = db.read(table)
        if table == "processes":
            df = df.sort_values(["sort", "name"])
        for _, c, _r in cols:
            if c not in df.columns:
                df[c] = None
        fr[sheet] = df[[c[1] for c in cols]].rename(columns={c[1]: c[0] for c in cols})
    settings = {k: db.get_setting(k) for k in db.DEFAULT_SETTINGS}
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        _write_std(xw, fr, settings)
    return buf.getvalue()


def _clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return db.py(v)


def import_standard(src, replace: bool = True, path: str | None = None) -> tuple[dict, list[str]]:
    """표준 양식 → DB. replace=True 면 양식에 있는 시트의 테이블을 교체, False 면 추가(일일불량/조치) 및 갱신."""
    xl = pd.read_excel(src, sheet_name=None, dtype=object)
    db.init_db(path)
    stats, warns = {}, []
    data = {}
    for sheet, (table, cols) in STD.items():
        if sheet not in xl:
            continue
        df = xl[sheet].dropna(how="all")
        miss = [k for k, _, r in cols if r and k not in df.columns]
        if miss:
            warns.append(f"[{sheet}] 필수 열 누락: {', '.join(miss)} → 시트 건너뜀")
            continue
        keep = [(k, c) for k, c, _ in cols if k in df.columns]
        df = df[[k for k, _ in keep]].rename(columns=dict(keep))
        req = [c for k, c, r in cols if r]
        bad = df[req].isna().any(axis=1)
        if bad.any():
            warns.append(f"[{sheet}] 필수값이 빈 {int(bad.sum())}행 제외")
            df = df[~bad]
        df = df.map(_clean)
        data[sheet] = df

    # 형 변환·기본값
    if "부품단가" in data:
        d = data["부품단가"]; d["price"] = pd.to_numeric(d["price"], errors="coerce").fillna(0.0)
    if "불량유형" in data:
        d = data["불량유형"]
        for c, dv in [("cost_class", "실패"), ("active", 1)]:
            if c not in d: d[c] = dv
            d[c] = d[c].fillna(dv)
        bad = ~d["cost_class"].isin(db.COST_CLASSES)
        if bad.any():
            warns.append(f"[불량유형] 비용구분은 실패/예방/제외 중 하나여야 합니다 → {int(bad.sum())}건 '실패'로 처리")
            d.loc[bad, "cost_class"] = "실패"
        d["active"] = pd.to_numeric(d["active"], errors="coerce").fillna(1).astype(int)
        data["불량유형"] = d.drop_duplicates("defect_type", keep="last")
    if "BOM" in data:
        d = data["BOM"]; d["qty"] = pd.to_numeric(d["qty"], errors="coerce").fillna(0)
        data["BOM"] = d[d["qty"] > 0].drop_duplicates(["defect_type", "part_name"], keep="last")
    if "일일불량" in data:
        d = data["일일불량"]
        d["date"] = pd.to_datetime(d["date"], errors="coerce").dt.strftime("%Y-%m-%d")
        d["qty"] = pd.to_numeric(d["qty"], errors="coerce").fillna(0).astype(int)
        d["exclude"] = pd.to_numeric(d.get("exclude"), errors="coerce").fillna(0).astype(int) if "exclude" in d else 0
        n0 = len(d); d = d[d["date"].notna() & (d["qty"] > 0)]
        if len(d) < n0:
            warns.append(f"[일일불량] 날짜 오류 또는 수량 0인 {n0 - len(d)}행 제외")
        d["created_at"] = db.now()
        data["일일불량"] = d
    if "귀책판정" in data:
        d = data["귀책판정"]
        for c in ["year", "month", "week"]:
            d[c] = pd.to_numeric(d[c], errors="coerce").astype("Int64")
        for c in ["retest_ng", "retest_ok", "resp_qty"]:
            d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0).astype(int) if c in d else 0
        if "actual_amount" in d:
            d["actual_amount"] = pd.to_numeric(d["actual_amount"], errors="coerce")
        if "claim_yn" in d:
            d["claim_yn"] = d["claim_yn"].map(lambda v: "Y" if str(v).strip().upper() in ("Y", "1", "TRUE", "O") else None)
        d = d.dropna(subset=["year", "month", "week"])
        d["key"] = [engine.make_key(*r) for r in d[["year", "month", "week", "process", "defect_type"]].itertuples(index=False)]
        d["updated_at"] = db.now()
        data["귀책판정"] = d.drop_duplicates("key", keep="last")
    if "목표" in data:
        d = data["목표"]
        d["annual"] = pd.to_numeric(d["annual"], errors="coerce").fillna(0.0)
        d["monthly"] = pd.to_numeric(d.get("monthly"), errors="coerce") if "monthly" in d else None
        d["monthly"] = d["monthly"].fillna(d["annual"] / 12)
        d["year"] = pd.to_numeric(d["year"], errors="coerce").astype(int)
    if "조치내역" in data:
        d = data["조치내역"]; d["updated_at"] = db.now()

    # 참조 무결성
    types = set(data["불량유형"]["defect_type"]) if "불량유형" in data else set(db.read("defect_types", path=path)["defect_type"])
    parts = set(data["부품단가"]["name"]) if "부품단가" in data else set(db.read("parts", path=path)["name"])
    if "BOM" in data:
        b = data["BOM"]
        mp = sorted(set(b["part_name"]) - parts)
        mt = sorted(set(b["defect_type"]) - types)
        if mp: warns.append(f"[BOM] 부품단가에 없는 부품 {len(mp)}개(단가 0 처리): {', '.join(mp[:8])}")
        if mt: warns.append(f"[BOM] 불량유형에 없는 유형 {len(mt)}개: {', '.join(mt[:8])}")
    if "일일불량" in data:
        mt = sorted(set(data["일일불량"]["defect_type"]) - types)
        if mt:
            warns.append(f"[일일불량] 불량유형 마스터에 없는 유형 {len(mt)}개 → 자동 등록(BOM 없음, 단가 0): {', '.join(mt[:8])}")
            add = data["일일불량"].drop_duplicates("defect_type")
            add = add[add["defect_type"].isin(mt)][["defect_type", "process"]].assign(cost_class="실패", active=1)
            data["불량유형"] = pd.concat([data.get("불량유형", pd.DataFrame()), add], ignore_index=True)

    # 저장
    procs = []
    if "공정" in data:
        p = data["공정"]
        p["sort"] = pd.to_numeric(p.get("sort"), errors="coerce") if "sort" in p else None
        p["sort"] = p["sort"].fillna(pd.Series(range(1, len(p) + 1), index=p.index)).astype(int)
        p["active"] = pd.to_numeric(p.get("active"), errors="coerce").fillna(1).astype(int) if "active" in p else 1
        data["공정"] = p.drop_duplicates("name", keep="last")
    order = ["공정", "부품단가", "불량유형", "BOM", "일일불량", "귀책판정", "목표", "조치내역"]
    for sheet in order:
        if sheet not in data:
            continue
        table = STD[sheet][0]
        df = data[sheet]
        if replace or table in ("processes", "parts", "defect_types"):
            if replace:
                db.replace_table(table, df, path)
            else:  # upsert by key
                key = {"processes": "name", "parts": "name", "defect_types": "defect_type"}[table]
                old = db.read(table, path=path)
                old = old[~old[key].isin(df[key])]
                if "id" in old.columns and "id" not in df.columns:
                    old = old.drop(columns=["id"])
                db.replace_table(table, pd.concat([old, df], ignore_index=True), path)
        else:
            if table == "bom":
                db.executemany("DELETE FROM bom WHERE defect_type=?", [(t,) for t in df["defect_type"].unique()], path)
            if table == "judgments":
                db.executemany("DELETE FROM judgments WHERE key=?", [(k,) for k in df["key"]], path)
            if table == "targets":
                db.executemany("DELETE FROM targets WHERE year=? AND process=?",
                               [(int(r.year), r.process) for r in df.itertuples()], path)
            cols = list(df.columns)
            db.executemany(f"INSERT INTO {table} ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                           [tuple(_clean(v) for v in r) for r in df.itertuples(index=False)], path)
        stats[sheet] = len(df)
    # 공정 자동 등록
    for sheet, col in [("불량유형", "process"), ("일일불량", "process"), ("목표", "process"), ("귀책판정", "process")]:
        if sheet in data and col in data[sheet]:
            procs += list(data[sheet][col].dropna())
    db.ensure_processes(procs, path)

    if "설정" in xl:
        sdf = xl["설정"].dropna(how="all")
        if {"항목", "값"} <= set(sdf.columns):
            for r in sdf.itertuples(index=False):
                k, v = str(r[0]).strip(), r[1]
                if k not in db.DEFAULT_SETTINGS or v is None or (isinstance(v, float) and pd.isna(v)):
                    continue
                if isinstance(db.DEFAULT_SETTINGS[k], list):
                    v = [x.strip() for x in str(v).split(",") if x.strip()]
                else:
                    v = str(v).strip()
                db.set_setting(k, v, path)
            stats["설정"] = len(sdf)
    db.log_change("전체", "표준 양식 가져오기 (" + ("교체" if replace else "추가") + ")",
                  ", ".join(f"{k} {v}" for k, v in stats.items()), "", path)
    return stats, warns


# ================================================================== 체험용 샘플
def generate_sample(year: int | None = None, months: int = 6, path: str | None = None, seed: int = 7) -> dict:
    """가상의 모터 조립 라인 데이터 (실제 회사 데이터 아님)."""
    rnd = random.Random(seed)
    year = year or date.today().year
    db.init_db(path)
    for t in ["processes", "parts", "defect_types", "bom", "daily", "judgments", "targets", "actions", "price_history"]:
        db.execute(f"DELETE FROM {t}", (), path)
    db.set_setting("company", "샘플전자(가상)", path)
    db.set_setting("system_name", "실패비용(COPQ) 관리 — 체험용", path)
    procs = [("가공", 1), ("서브조립", 2), ("메인조립", 3), ("최종검사", 4)]
    db.executemany("INSERT INTO processes(name,sort,active) VALUES (?,?,1)", procs, path)
    parts = {"샤프트": 850, "베어링": 420, "하우징": 2300, "로터": 5200, "스테이터": 7400, "커버": 600,
             "기어": 950, "PCB": 6800, "커넥터": 380, "실링": 120, "가공비": 1500, "조립비": 900}
    db.executemany("INSERT INTO parts(name,price,changed_at) VALUES (?,?,?)",
                   [(k, v, f"{year}-01-01") for k, v in parts.items()], path)
    types = [  # 유형, 공정, 비용구분, 원인분류, 처리방안, 상태, BOM
        ("샤프트 외경 NG", "가공", "실패", "공정", "폐기", "단품", {"샤프트": 1, "가공비": 1}),
        ("하우징 버(Burr)", "가공", "실패", "공정", "재검", "단품", {"하우징": 1, "가공비": 1}),
        ("하우징 기공", "가공", "실패", "협력사", "폐기", "단품", {"하우징": 1}),
        ("베어링 압입불량", "서브조립", "실패", "공정", "폐기", "SUB", {"샤프트": 1, "베어링": 2, "로터": 1, "조립비": 1}),
        ("로터 밸런스 NG", "서브조립", "실패", "협력사", "재검", "SUB", {"로터": 1, "샤프트": 1}),
        ("스테이터 절연불량", "서브조립", "실패", "협력사", "폐기", "SUB", {"스테이터": 1}),
        ("기어 소음", "메인조립", "실패", "경계치", "재검", "완성품", {"기어": 2, "하우징": 1, "조립비": 1}),
        ("커넥터 체결불량", "메인조립", "실패", "공정", "폐기", "완성품", {"커넥터": 1, "조립비": 1}),
        ("PCB 기능불량", "메인조립", "실패", "협력사", "재검", "완성품", {"PCB": 1}),
        ("실링 누설", "최종검사", "실패", "공정", "폐기", "완성품", {"실링": 2, "커버": 1, "조립비": 1}),
        ("전류 상한 NG", "최종검사", "실패", "NTF", "재검", "완성품",
         {"로터": 1, "스테이터": 1, "PCB": 1, "하우징": 1, "조립비": 1}),
        ("초도품 검증(파괴시험)", "최종검사", "예방", "공정", "폐기", "완성품",
         {"로터": 1, "스테이터": 1, "PCB": 1, "하우징": 1, "커버": 1}),
        ("설비 셋업 시험품", "가공", "제외", "공정", "폐기", "단품", {"샤프트": 1}),
    ]
    db.executemany("INSERT INTO defect_types(defect_type,process,cost_class,initial_resp,disposition,state,active) "
                   "VALUES (?,?,?,?,?,?,1)", [t[:6] for t in types], path)
    db.executemany("INSERT INTO bom(defect_type,part_name,qty) VALUES (?,?,?)",
                   [(t[0], p, q) for t in types for p, q in t[6].items()], path)
    rate = {t[0]: rnd.uniform(0.2, 3.0) for t in types}
    rows = []
    d = date(year, 1, 1)
    end = date(year, months, 28)
    while d <= end:
        if d.weekday() < 5:
            for t in types:
                if rnd.random() < 0.35:
                    q = max(1, int(rnd.expovariate(1 / rate[t[0]])))
                    rows.append((d.isoformat(), t[1], t[0], q, None, 0, db.now()))
        d += timedelta(days=1)
    db.executemany("INSERT INTO daily(date,process,defect_type,qty,remark,exclude,created_at) VALUES (?,?,?,?,?,?,?)",
                   rows, path)
    targets = {"가공": 900_000, "서브조립": 1_300_000, "메인조립": 1_000_000, "최종검사": 1_600_000}
    db.executemany("INSERT INTO targets(year,process,annual,monthly) VALUES (?,?,?,?)",
                   [(year, k, v, v / 12) for k, v in targets.items()], path)
    # 판정: 협력사 원인 일부를 공급 귀책·클레임으로
    old = db.DB_PATH
    if path:
        db.DB_PATH = path
    try:
        agg = engine.daily_agg(year)
    finally:
        db.DB_PATH = old
    vendors = {"하우징 기공": "가나주조", "로터 밸런스 NG": "다라모터", "스테이터 절연불량": "다라모터", "PCB 기능불량": "마바전자"}
    jud, ncr = [], 0
    for r in agg.itertuples():
        if r.defect_type in vendors and rnd.random() < 0.6:
            ncr += 1
            rq = int(r.qty)
            done = r.month < months - 1
            jud.append((r.key, year, r.month, r.week, r.process, r.defect_type, "협력사", "공급", 0, 0, None,
                        f"NCR-{year % 100:02d}{r.month:02d}-{ncr:03d}", "Y", vendors[r.defect_type], rq,
                        f"CL-{year}-{r.month:02d}" if done else None,
                        float(rq * r.unit_cost * rnd.choice([1.0, 0.7, 0.5])) if done else None,
                        f"{year}-{r.month + 1:02d}-15" if done else None, None, db.now()))
        elif r.disposition == "재검" and rnd.random() < 0.5:
            ok = int(r.qty * rnd.uniform(0.3, 0.8))
            jud.append((r.key, year, r.month, r.week, r.process, r.defect_type, r.guess_cause, "사내", r.qty - ok, ok,
                        rnd.choice(["재사용", "재사용", "보류"]), None, None, None, 0, None, None, None, None, db.now()))
    db.executemany("INSERT INTO judgments(key,year,month,week,process,defect_type,confirmed_cause,responsibility,retest_ng,"
                   "retest_ok,reuse_flag,ncr_no,claim_yn,claim_vendor,resp_qty,claim_no,actual_amount,agreed_date,memo,updated_at) "
                   "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", jud, path)
    db.executemany("INSERT INTO actions(year,month,defect_type,analysis,result,status,owner,due_date,updated_at) VALUES (?,?,?,?,?,?,?,?,?)",
                   [(year, months, "베어링 압입불량", "압입 지그 마모로 동심도 저하", "지그 교체 및 일상점검 항목 추가", "개선 중", "홍길동", f"{year}-{months:02d}-28", db.now()),
                    (year, months, "PCB 기능불량", "협력사 리플로우 온도 프로파일 이탈", "8D 요청, 클레임 협의", "클레임 예정", "김품질", None, db.now())], path)
    db.log_change("전체", "체험용 샘플 데이터 생성", "", "", path)
    return {"일일불량": len(rows), "귀책판정": len(jud), "불량유형": len(types), "부품": len(parts)}
