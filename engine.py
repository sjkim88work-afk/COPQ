"""계산 엔진 — 범용 실패비용(COPQ) 계산 (원본: ECV VER_9.4 엑셀 Power Query 로직).

  실패비용 단가  = Σ(유형별BOM 사용수량 × 부품단가)            [쿼리: 유형별BOM]
  일일집계       = 연|월|주차|공정|불량유형 단위 그룹 + 귀책 병합   [쿼리: 일일불량]
      순수폐기수량  = 수량 − 귀책수량 − 재사용수량 − 재사용보류 − 재검불량
      폐기금액      = (재검불량 + 순수폐기수량) × 단가
      REWORK        = 재사용수량 × 단가
      손실(잠재)    = 재사용보류 × 단가            (예방 제외)
      클레임        = 귀책수량 × 단가
      예방비용      = 비용구분='예방' 이면 추정손실
      손실(확정)    = 비용구분≠'예방' 이면 폐기금액
      손실비용      = 확정 + 잠재
  월불량         = 월·공정·불량유형 그룹 + 검증(클레임+예방+REWORK+잠재+확정 = 추정)
  클레임원장     = 귀책입력 중 클레임(Y) → 청구금액/진행상태/분담률/조정액
  KPI            = 공정별 연간한도 대비 누적, 월별 세부현황, 클레임 현황
"""
from __future__ import annotations

from datetime import date

import numpy as np
import pandas as pd

import db


# ---------------------------------------------------------------- 공통
def excel_weeknum(d: date, monday: bool | None = None) -> int:
    """엑셀 WEEKNUM: 1월 1일이 속한 주 = 1주차.
    monday=False → WEEKNUM(date,1) 일요일 시작 / True → WEEKNUM(date,2) 월요일 시작.
    None 이면 데이터셋 설정(week_start)을 따른다."""
    if monday is None:
        monday = db.get_setting("week_start") == "mon"
    jan1 = date(d.year, 1, 1)
    offset = jan1.weekday() if monday else (jan1.weekday() + 1) % 7
    return (d.timetuple().tm_yday + offset - 1) // 7 + 1


def _maxdate(s):
    v = [x for x in s if isinstance(x, str) and x]
    return max(v) if v else None


def make_key(year, month, week, process, defect_type) -> str:
    return f"{int(year)}|{int(month)}|{int(week)}|{process}|{defect_type}"


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)


# ---------------------------------------------------------------- 기준정보
def unit_costs() -> pd.DataFrame:
    """불량유형별 실패비용 단가 (BOM × 부품단가)."""
    parts = db.read("parts")[["name", "price"]]
    bom = db.read("bom")
    types = db.read("defect_types")[["defect_type", "process"]]
    if bom.empty:
        out = types.copy()
        out["unit_cost"] = 0.0
        return out
    m = bom.merge(parts, left_on="part_name", right_on="name", how="left")
    m["amt"] = _num(m["qty"]) * _num(m["price"])
    uc = m.groupby("defect_type", as_index=False)["amt"].sum().rename(columns={"amt": "unit_cost"})
    out = types.merge(uc, on="defect_type", how="left")
    out["unit_cost"] = out["unit_cost"].fillna(0.0)
    return out


def bom_matrix() -> pd.DataFrame:
    """엑셀과 같은 모양의 1/0 매트릭스 (행=불량유형, 열=부품)."""
    parts = db.read("parts").sort_values("id")["name"].tolist()
    types = db.read("defect_types")[["process", "defect_type"]]
    bom = db.read("bom")
    if bom.empty:
        mat = pd.DataFrame(0, index=types["defect_type"], columns=parts)
    else:
        mat = bom.pivot_table(index="defect_type", columns="part_name", values="qty", aggfunc="sum").reindex(columns=parts)
        mat = mat.reindex(types["defect_type"]).fillna(0)
    mat = mat.reset_index().rename(columns={"index": "defect_type"})
    mat.insert(0, "process", mat["defect_type"].map(types.set_index("defect_type")["process"]))
    return mat


# ---------------------------------------------------------------- 일일 불량
def daily_enriched(year: int | None = None) -> pd.DataFrame:
    """일일불량입력 + 자동조회 컬럼 (엑셀 G~K 열 VLOOKUP 에 해당)."""
    d = db.read("daily")
    if d.empty:
        return d
    d["date"] = pd.to_datetime(d["date"])
    if year:
        d = d[d["date"].dt.year == year]
    d["year"] = d["date"].dt.year
    d["month"] = d["date"].dt.month
    monday = db.get_setting("week_start") == "mon"
    d["week"] = d["date"].dt.date.map(lambda x: excel_weeknum(x, monday))
    t = db.read("defect_types").set_index("defect_type")
    uc = unit_costs().set_index("defect_type")["unit_cost"]
    d["unit_cost"] = d["defect_type"].map(uc).fillna(0.0)
    d["cost_class"] = d["defect_type"].map(t["cost_class"]).fillna("실패")
    d["state"] = d["defect_type"].map(t["state"]).fillna("")
    d["guess_cause"] = d["defect_type"].map(t["initial_resp"]).fillna("")
    d["disposition"] = d["defect_type"].map(t["disposition"]).fillna("")
    d["est_loss"] = d["qty"] * d["unit_cost"]
    return d.sort_values(["date", "id"])


JUDG_NUM = ["retest_ng", "retest_ok", "resp_qty"]


def judgments(year: int | None = None) -> pd.DataFrame:
    j = db.read("judgments") if not year else db.read("judgments", "year=?", (year,))
    for c in JUDG_NUM:
        j[c] = _num(j[c]).astype(int)
    j["actual_amount"] = pd.to_numeric(j["actual_amount"], errors="coerce")
    flag = j["reuse_flag"].fillna("")
    j["reuse_qty"] = np.where(flag == "재사용", j["retest_ok"], 0)
    j["hold_qty"] = np.where(flag == "보류", j["retest_ok"], 0)
    # 엑셀 귀책합계: 공급 → 재사용+보류+귀책수량 / 그 외 → 재검불량+재사용+보류
    j["resp_total"] = np.where(j["responsibility"] == "공급",
                               j["reuse_qty"] + j["hold_qty"] + j["resp_qty"],
                               j["retest_ng"] + j["reuse_qty"] + j["hold_qty"])
    return j


def daily_agg(year: int) -> pd.DataFrame:
    """[쿼리: 일일불량] 주차 단위 집계 + 귀책 병합 + COPQ 분리."""
    d = daily_enriched(year)
    if d.empty:
        return pd.DataFrame()
    d = d[(d["cost_class"] != "제외") & (d["exclude"].fillna(0) == 0)].copy()
    d["remark"] = d["remark"].fillna("")
    g = (d.groupby(["year", "month", "week", "process", "defect_type",
                    "cost_class", "state", "guess_cause", "disposition"], as_index=False)
         .agg(qty=("qty", "sum"), unit_cost=("unit_cost", "max"),
              remark=("remark", lambda s: " / ".join(x for x in s if x))))
    g["key"] = [make_key(*r) for r in g[["year", "month", "week", "process", "defect_type"]].itertuples(index=False)]
    g["est_loss"] = g["qty"] * g["unit_cost"]

    j = judgments(year)
    cols = ["key", "resp_qty", "retest_ng", "reuse_qty", "hold_qty", "confirmed_cause", "responsibility",
            "ncr_no", "claim_yn", "claim_vendor", "resp_total"]
    g = g.merge(j[cols], on="key", how="left")
    for c in ["resp_qty", "retest_ng", "reuse_qty", "hold_qty"]:
        g[c] = g[c].fillna(0).astype(int)

    u = g["unit_cost"]
    g["pure_scrap"] = g["qty"] - g["resp_qty"] - g["reuse_qty"] - g["hold_qty"] - g["retest_ng"]
    g["scrap_amt"] = (g["retest_ng"] + g["pure_scrap"]) * u
    g["rework_amt"] = g["reuse_qty"] * u
    g["hold_amt"] = g["hold_qty"] * u
    g["claim_amt"] = g["resp_qty"] * u
    prev = g["cost_class"] == "예방"
    g["prevent_amt"] = np.where(prev, g["est_loss"], 0.0)
    g["loss_confirmed"] = np.where(~prev, g["scrap_amt"], 0.0)
    g["loss_potential"] = np.where(~prev, g["hold_amt"], 0.0)
    g["loss_total"] = g["loss_confirmed"] + g["loss_potential"]
    g["match"] = np.where(g["resp_total"].fillna(0) > 0,
                          np.where(g["resp_total"] == g["qty"], "OK", "차이있음"), "")
    return g.sort_values(["month", "week", "process", "defect_type"]).reset_index(drop=True)


def monthly(year: int, agg: pd.DataFrame | None = None) -> pd.DataFrame:
    """[쿼리: 월불량]"""
    g = daily_agg(year) if agg is None else agg
    if g.empty:
        return pd.DataFrame(columns=["month", "process", "defect_type", "qty", "est_loss", "claim_amt",
                                     "prevent_amt", "rework_amt", "loss_potential", "loss_confirmed",
                                     "loss_total", "check"])
    m = (g.groupby(["month", "process", "defect_type"], as_index=False)
         .agg(qty=("qty", "sum"), est_loss=("est_loss", "sum"), claim_amt=("claim_amt", "sum"),
              prevent_amt=("prevent_amt", "sum"), rework_amt=("rework_amt", "sum"),
              loss_potential=("hold_amt", "sum"),  # 엑셀 월불량: 재사용보류 금액 합
              loss_confirmed=("loss_confirmed", "sum"), loss_total=("loss_total", "sum")))
    s = m["claim_amt"] + m["prevent_amt"] + m["rework_amt"] + m["loss_potential"] + m["loss_confirmed"]
    diff = (s.round(0) - m["est_loss"].round(0))
    m["check"] = np.where(diff == 0, "✅", "❌ " + (s - m["est_loss"]).round(0).astype(int).astype(str))
    return m.sort_values(["month", "est_loss"], ascending=[True, False]).reset_index(drop=True)


# ---------------------------------------------------------------- 클레임
def claims(year: int) -> pd.DataFrame:
    """[쿼리: 클레임] 귀책입력 중 클레임(Y)."""
    j = judgments(year)
    j = j[j["claim_yn"].fillna("").str.upper() == "Y"].copy()
    cols = ["month", "process", "ncr_no", "claim_vendor", "defect_type", "unit_cost", "resp_qty",
            "claim_amt", "claim_no", "status", "actual_amount", "share", "adjust", "agreed_date"]
    if j.empty:
        return pd.DataFrame(columns=cols)
    uc = unit_costs().set_index("defect_type")["unit_cost"]
    j["unit_cost"] = j["defect_type"].map(uc).fillna(0.0)
    j["claim_amt"] = j["resp_qty"] * j["unit_cost"]
    j["claim_no"] = j["claim_no"].fillna("").astype(str).str.strip()
    j["status"] = np.where(j["claim_no"] != "", "청구완료", "청구대기")
    for c in ["ncr_no", "claim_vendor"]:
        j[c] = j[c].fillna("")
    g = (j.groupby(["month", "process", "defect_type", "claim_vendor", "ncr_no", "claim_no", "unit_cost", "status"],
                   as_index=False, dropna=False)
         .agg(resp_qty=("resp_qty", "sum"), claim_amt=("claim_amt", "sum"),
              actual_amount=("actual_amount", lambda s: s.sum(min_count=1)),
              agreed_date=("agreed_date", _maxdate)))
    g["actual_amount"] = g["actual_amount"].fillna(0.0)
    done = g["status"] == "청구완료"
    g["share"] = np.where(done & (g["claim_amt"] != 0), g["actual_amount"] / g["claim_amt"].replace(0, np.nan), 0.0)
    g["adjust"] = np.where(done, g["claim_amt"] - g["actual_amount"], 0.0)
    g["share"] = g["share"].fillna(0.0)
    return g[cols].sort_values(["month", "claim_vendor"]).reset_index(drop=True)


def claims_by_ncr(year: int, upto_month: int) -> pd.DataFrame:
    """KPI관리 '클레임 세부 현황' (NCR 단위 누적)."""
    c = claims(year)
    c = c[(c["month"] <= upto_month) & (c["ncr_no"] != "")]
    if c.empty:
        return pd.DataFrame(columns=["ncr_no", "claim_vendor", "defect_types", "resp_qty", "claim_amt", "status",
                                     "actual_amount", "share", "adjust", "agreed_date"])
    g = (c.groupby("ncr_no", as_index=False)
         .agg(claim_vendor=("claim_vendor", "first"),
              defect_types=("defect_type", lambda s: ", ".join(dict.fromkeys(s))),
              resp_qty=("resp_qty", "sum"), claim_amt=("claim_amt", "sum"),
              status=("status", lambda s: "청구완료" if (s == "청구완료").all() else ("일부청구" if (s == "청구완료").any() else "청구대기")),
              actual_amount=("actual_amount", "sum"),
              agreed_date=("agreed_date", _maxdate)))
    g["share"] = np.where(g["claim_amt"] > 0, g["actual_amount"] / g["claim_amt"].replace(0, np.nan), 0.0)
    g["adjust"] = g["claim_amt"] - g["actual_amount"]
    return g.sort_values("ncr_no").reset_index(drop=True)


# ---------------------------------------------------------------- KPI
def targets(year: int) -> pd.DataFrame:
    t = db.read("targets", "year=?", (year,))
    base = pd.DataFrame({"process": db.get_processes()})
    t = base.merge(t[["process", "annual", "monthly"]], on="process", how="left").fillna(0.0)
    return t


def kpi_limits(year: int, month: int, mon: pd.DataFrame) -> pd.DataFrame:
    """실패비용 한도관리 (KPI관리 A18:H23)."""
    t = targets(year)
    upto = mon[mon["month"] <= month]
    cum = upto.groupby("process")["loss_total"].sum()
    t["cum"] = t["process"].map(cum).fillna(0.0)
    tot = pd.DataFrame([{"process": "Total", "annual": t["annual"].sum(), "monthly": t["monthly"].sum(),
                         "cum": t["cum"].sum()}])
    t = pd.concat([tot, t], ignore_index=True)
    t["annual_rate"] = np.where(t["annual"] > 0, t["cum"] / t["annual"].replace(0, np.nan), 0.0)
    t["remain"] = t["annual"] - t["cum"]
    t["avg_month"] = t["cum"] / month if month else 0.0
    # 월평균 손실이 월간목표 대비 몇 % 인지 (100% 초과 = 한도 초과 페이스)
    t["month_rate"] = np.where(t["monthly"] > 0, t["avg_month"] / t["monthly"].replace(0, np.nan), 0.0)
    return t.fillna(0.0)


def kpi_monthly(year: int, month: int, mon: pd.DataFrame) -> pd.DataFrame:
    """월별 실패비용 세부 현황 (KPI관리 A26:L39)."""
    monthly_target = targets(year)["monthly"].sum()
    rows = []
    cum = 0.0
    prev = None
    for m in range(1, 13):
        if m > month:
            rows.append({"month": m, "target": monthly_target})
            continue
        x = mon[mon["month"] == m]
        lt = x["loss_total"].sum()
        cum += lt
        rows.append({"month": m, "est_loss": x["est_loss"].sum(), "claim_amt": x["claim_amt"].sum(),
                     "prevent_amt": x["prevent_amt"].sum(), "rework_amt": x["rework_amt"].sum(),
                     "loss_potential": x["loss_potential"].sum(), "loss_confirmed": x["loss_confirmed"].sum(),
                     "loss_total": lt, "target": monthly_target, "over": lt - monthly_target,
                     "mom": (lt - prev) / prev if prev else None, "cum": cum})
        prev = lt
    return pd.DataFrame(rows)


def claim_summary(year: int, month: int) -> dict:
    c = claims(year)
    c = c[c["month"] <= month]
    pos = c[c["claim_amt"] > 0]
    total = pos["claim_amt"].sum()
    actual = c["actual_amount"].sum()
    return {
        "건수": len(pos),
        "NCR수": pos.loc[pos["ncr_no"] != "", "ncr_no"].nunique(),
        "청구완료": int((c["status"] == "청구완료").sum()),
        "총청구금액": total,
        "실청구액": actual,
        "조정액": c["adjust"].sum(),
        "회수율": actual / total if total else 0.0,
    }


# ---------------------------------------------------------------- 파레토
def pareto(mon: pd.DataFrame, month: int, by: str = "loss_total", process: str | None = None) -> pd.DataFrame:
    x = mon[mon["month"] == month]
    if process:
        x = x[x["process"] == process]
    x = x.sort_values(by, ascending=False).reset_index(drop=True)
    tot = x[by].sum()
    x["cum_pct"] = x[by].cumsum() / tot if tot else 0.0
    return x


# ---------------------------------------------------------------- 데이터 검증
def data_issues(year: int) -> pd.DataFrame:
    """입력 오류 후보: 순수폐기 음수, 귀책수량 매칭 차이, 단가 0, 귀책 키가 일일 데이터에 없음."""
    g = daily_agg(year)
    out = []
    if not g.empty:
        over = g["qty"] - g["resp_qty"] - g["reuse_qty"] - g["hold_qty"]
        for r in g[over < 0].itertuples():
            out.append(("수량 초과배분", r.key, f"수량 {r.qty} < 귀책+재사용+보류 합"))
        for r in g[g["match"] == "차이있음"].itertuples():
            out.append(("귀책합계≠수량", r.key, f"귀책합계 {int(r.resp_total)} / 수량 {r.qty}"))
        for r in g[g["unit_cost"] <= 0].drop_duplicates("defect_type").itertuples():
            out.append(("단가 0", r.defect_type, "유형별BOM 매핑 확인 필요"))
        j = judgments(year)
        missing = j[~j["key"].isin(g["key"])]
        for r in missing.itertuples():
            out.append(("귀책 키 불일치", r.key, "해당 주차·공정·유형의 일일불량 데이터 없음"))
        cl = j[(j["claim_yn"].fillna("") == "Y") & ((j["claim_vendor"].fillna("") == "") | (j["resp_qty"] <= 0))]
        for r in cl.itertuples():
            out.append(("클레임 정보 누락", r.key, "클레임 업체 또는 귀책수량 미입력"))
    return pd.DataFrame(out, columns=["구분", "대상", "내용"])
