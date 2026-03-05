from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
import pandas as pd
import re
from datetime import datetime
from typing import Optional, Dict, Any, List

app = FastAPI(title="당직수당 대시보드", version="2.0.0")

# -----------------------------
# In-memory store (Free 플랜에서는 재시작 시 초기화될 수 있음)
# -----------------------------
LATEST: Dict[str, Any] = {"META": {"title": "데이터 없음"}, "RAW": []}
LATEST_META: Dict[str, Any] = {"uploaded_at": None, "filename": None}

# -----------------------------
# Excel parsing helpers
# -----------------------------
def _pick_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = [str(c).strip() for c in df.columns]
    # exact
    for cand in candidates:
        for real in cols:
            if real == cand:
                return real
    # contains
    for cand in candidates:
        for real in cols:
            if cand in real:
                return real
    return None

def _to_month(v) -> Optional[str]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.strftime("%Y-%m")
    s = str(v).strip()
    m = re.search(r"(20\d{2})[./-]?\s*(\d{1,2})", s)
    if not m:
        return None
    y = int(m.group(1))
    mm = int(m.group(2))
    return f"{y:04d}-{mm:02d}"

def _to_date(v) -> Optional[str]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.search(r"(20\d{2})[./-](\d{1,2})[./-](\d{1,2})", s)
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return f"{y:04d}-{mo:02d}-{d:02d}"

def _type_group(raw_type: Any) -> str:
    s = str(raw_type or "")
    if "직책" in s:
        return "직책수당"
    if "출동" in s:
        return "휴일출동"
    if "휴일" in s or "주말" in s:
        return "휴일"
    if "평일" in s:
        return "평일"
    return "기타"

def excel_to_payload(fileobj) -> Dict[str, Any]:
    df = pd.read_excel(fileobj, sheet_name=0, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    c_name = _pick_col(df, ["이름", "성명", "담당자", "name"])
    if not c_name:
        raise ValueError("필수 컬럼(이름/성명)이 없습니다. 엑셀 헤더를 확인해주세요.")

    c_org  = _pick_col(df, ["센터", "소속", "org", "조직"])
    c_hq   = _pick_col(df, ["본부", "hq", "본부명", "부서"])
    c_date = _pick_col(df, ["근무일", "일자", "workDate", "근무일자"])
    c_wm   = _pick_col(df, ["근무월", "workMonth"])
    c_pm   = _pick_col(df, ["지급월", "payMonth"])
    c_cat  = _pick_col(df, ["구분", "category", "카테고리"])
    c_type = _pick_col(df, ["유형", "rawType", "세부유형", "당직유형"])
    c_days = _pick_col(df, ["일수", "days", "당직일수"])
    c_pay  = _pick_col(df, ["수당", "금액", "pay", "당직수당"])
    c_det  = _pick_col(df, ["비고", "detail", "상세", "설명", "내용"])
    c_rng  = _pick_col(df, ["기간", "dateLabel", "근무기간"])

    out = pd.DataFrame()
    out["name"] = df[c_name].astype(str).str.strip()
    out["org"] = df[c_org].astype(str).str.strip() if c_org else "-"
    out["hq"] = df[c_hq].astype(str).str.strip() if c_hq else "-"
    out["workDate"] = df[c_date].apply(_to_date) if c_date else None
    out["workMonth"] = df[c_wm].apply(_to_month) if c_wm else (df[c_date].apply(_to_month) if c_date else None)
    out["payMonth"] = df[c_pm].apply(_to_month) if c_pm else None
    out["category"] = df[c_cat].astype(str).str.strip() if c_cat else "-"
    out["rawType"] = df[c_type].astype(str).str.strip() if c_type else "-"
    out["typeGroup"] = out["rawType"].apply(_type_group)
    out["days"] = pd.to_numeric(df[c_days], errors="coerce").fillna(0) if c_days else 0
    out["pay"] = pd.to_numeric(df[c_pay], errors="coerce").fillna(0) if c_pay else 0
    out["detail"] = df[c_det].astype(str).fillna("").str.strip() if c_det else ""
    out["dateLabel"] = df[c_rng].astype(str).fillna("").str.strip() if c_rng else ""

    out["isAnomaly"] = out["workMonth"].isna() | (out["name"].isna()) | (out["name"] == "")

    months_work = sorted([m for m in out["workMonth"].dropna().unique().tolist()])
    months_pay = sorted([m for m in out["payMonth"].dropna().unique().tolist()])
    types = sorted(out["typeGroup"].dropna().unique().tolist())
    hqs = sorted(out["hq"].dropna().unique().tolist())

    payload = {
        "META": {
            "title": "주말·공휴일 당직근무 이력 대시보드",
            "recordCount": int(len(out)),
            "personCount": int(out["name"].nunique(dropna=True)),
            "totalDays": float(out["days"].sum()),
            "totalPay": float(out["pay"].sum()),
            "monthsWork": months_work,
            "monthsPay": months_pay,
            "types": types,
            "hqs": hqs,
            "anomalyCount": int(out["isAnomaly"].sum()),
            "missingDateCount": int(out["workDate"].isna().sum()),
        },
        "RAW": out.where(pd.notnull(out), None).to_dict(orient="records")
    }
    return payload

# -----------------------------
# HTML pages (Upload + Dashboard)
# -----------------------------
UPLOAD_HTML = """
<!doctype html>
<html lang="ko"><head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>당직수당 대시보드 업로드</title>
<style>
  :root{--bg:#f4f7fb;--card:#fff;--bd:#dbe3ee;--txt:#0f172a;--mut:#64748b;--pri:#2563eb;}
  body{font-family:system-ui,-apple-system,Segoe UI,Arial,"Noto Sans KR",sans-serif;background:var(--bg);margin:0;color:var(--txt)}
  .wrap{max-width:980px;margin:0 auto;padding:22px}
  .card{background:var(--card);border:1px solid var(--bd);border-radius:18px;box-shadow:0 8px 24px rgba(15,23,42,.08);padding:18px}
  h1{margin:4px 0 8px;font-size:22px}
  .m{color:var(--mut);font-size:13px;line-height:1.5}
  .row{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px;align-items:center}
  input[type=file]{flex:1;min-width:260px}
  button{background:var(--pri);color:#fff;border:1px solid var(--pri);border-radius:12px;padding:10px 14px;font-weight:800;cursor:pointer}
  a.btn{display:inline-block;text-decoration:none;background:#fff;color:var(--txt);border:1px solid var(--bd);border-radius:12px;padding:10px 14px;font-weight:800}
  .pill{display:inline-block;background:#eef2ff;color:#3730a3;border:1px solid #c7d2fe;border-radius:999px;padding:6px 10px;font-size:12px;font-weight:700}
  .grid{display:grid;gap:12px;margin-top:12px}
  @media(min-width:900px){.grid{grid-template-columns:1fr 1fr}}
  .box{border:1px solid #eef2f7;border-radius:14px;padding:14px;background:linear-gradient(180deg,#fff,#fbfdff)}
  .k{font-size:12px;color:var(--mut)}
  .v{margin-top:6px;font-size:18px;font-weight:900}
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <div class="pill">업로드형 · 최신 데이터 자동 반영</div>
    <h1 style="margin-top:10px">당직수당 대시보드 · 엑셀 업로드</h1>
    <div class="m">
      누구든지 엑셀 업로드 → 최신 대시보드가 자동 갱신됩니다.<br/>
      업로드 후 <b>대시보드 보기</b>로 이동하세요.
    </div>

    <form class="row" action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="file" accept=".xlsx,.xls" required/>
      <button type="submit">업로드</button>
      <a class="btn" href="/dashboard">대시보드 보기</a>
    </form>

    <div class="grid">
      <div class="box">
        <div class="k">현재 상태</div>
        <div class="v" id="statusLine">로딩 중...</div>
        <div class="m" id="statusSub" style="margin-top:6px">-</div>
      </div>
      <div class="box">
        <div class="k">사용 방법(컴맹용)</div>
        <div class="m" style="margin-top:6px">
          1) 파일 선택 → 2) 업로드 → 3) 대시보드 보기 클릭<br/>
          * 무료 서버는 한동안 안 쓰면 잠들 수 있어요(처음 접속이 느릴 수 있음).
        </div>
      </div>
    </div>
  </div>
</div>

<script>
async function boot(){
  try{
    const meta = await fetch("/data/latest_meta").then(r=>r.json());
    const data = await fetch("/data/latest").then(r=>r.json());
    const M = data.META || {};
    const status = document.getElementById("statusLine");
    const sub = document.getElementById("statusSub");
    if(meta && meta.uploaded_at){
      status.textContent = "최근 업로드: " + meta.uploaded_at;
      sub.textContent = "파일: " + (meta.filename || "-") + " / 레코드 " + (M.recordCount||0) + " / 인원 " + (M.personCount||0);
    }else{
      status.textContent = "아직 업로드된 데이터가 없습니다.";
      sub.textContent = "엑셀을 업로드하면 대시보드가 채워집니다.";
    }
  }catch(e){
    document.getElementById("statusLine").textContent = "서버 깨우는 중일 수 있어요. 잠시 후 새로고침 해주세요.";
    document.getElementById("statusSub").textContent = "";
  }
}
boot();
</script>
</body></html>
"""

DASHBOARD_HTML = r"""
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>당직수당 대시보드</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{
    --bg:#f4f7fb;--card:#fff;--bd:#dbe3ee;--txt:#0f172a;--mut:#64748b;
    --pri:#2563eb;--pri2:#60a5fa;--ok:#0f766e;--warn:#b45309;
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Arial,"Noto Sans KR",sans-serif;background:var(--bg);margin:0;color:var(--txt)}
  .wrap{max-width:1520px;margin:0 auto;padding:18px}
  .card{background:var(--card);border:1px solid var(--bd);border-radius:18px;box-shadow:0 8px 24px rgba(15,23,42,.08);padding:16px}
  .top{display:grid;gap:12px}
  @media(min-width:1100px){.top{grid-template-columns:1fr 420px}}
  h1{margin:6px 0 0;font-size:22px}
  .mut{color:var(--mut);font-size:12px;line-height:1.5}
  .btn{display:inline-block;text-decoration:none;border-radius:12px;padding:10px 12px;font-weight:900;border:1px solid var(--bd);background:#fff;color:var(--txt);cursor:pointer}
  .btnP{background:var(--pri);border-color:var(--pri);color:#fff}
  .row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  .kpiGrid{display:grid;gap:10px;margin-top:12px}
  @media(min-width:1100px){.kpiGrid{grid-template-columns:repeat(4,1fr)}}
  .kpi{border:1px solid #eef2f7;border-radius:16px;padding:14px;background:linear-gradient(180deg,#fff,#fbfdff)}
  .kpi .l{font-size:12px;color:var(--mut)}
  .kpi .v{font-size:28px;font-weight:950;margin-top:6px}
  .grid2{display:grid;gap:12px;margin-top:12px}
  @media(min-width:1100px){.grid2{grid-template-columns:1fr 1fr}}
  .filters{display:grid;gap:10px;margin-top:10px}
  @media(min-width:900px){.filters{grid-template-columns:repeat(6,1fr)}}
  label{display:block;font-size:12px;color:var(--mut);margin-bottom:6px;font-weight:800}
  select,input[type=text]{width:100%;padding:10px 10px;border-radius:12px;border:1px solid var(--bd);background:#fff}
  .tog{display:flex;gap:8px;align-items:center}
  .chips{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px}
  .chip{font-size:12px;color:#111827;background:#f1f5f9;border:1px solid #e2e8f0;padding:6px 10px;border-radius:999px;font-weight:800}
  .warn{background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;border-radius:14px;padding:10px;font-size:12px;margin-top:10px}
  .tblWrap{overflow:auto;border:1px solid #eef2f7;border-radius:16px;background:#fff}
  table{border-collapse:collapse;width:100%}
  th,td{padding:10px;border-bottom:1px solid #eef2f7;font-size:13px}
  th{background:#fafcff;color:var(--mut);font-size:12px;text-align:left}
  td.r,th.r{text-align:right}
  .sticky{position:sticky;left:0;background:#fff}
  .panelTitle{font-weight:950;font-size:16px}
  .small{font-size:12px;color:var(--mut)}
  textarea{width:100%;border-radius:12px;border:1px solid var(--bd);padding:10px;font-size:13px}
  .log{border:1px solid #eef2f7;border-radius:16px;background:#fff;max-height:260px;overflow:auto}
  .logItem{padding:10px;border-bottom:1px solid #eef2f7}
  .logOk{background:#ecfdf5}
  .logInfo{background:#fffbeb}
</style>
</head>
<body>
<div class="wrap">

  <div class="top">
    <div class="card">
      <div class="row" style="justify-content:space-between">
        <div>
          <div class="small" id="badgeLine">반응형 · 업로드형 · 최신 데이터 자동 반영</div>
          <h1 id="title">당직 대시보드</h1>
          <div class="mut" id="metaLine">로딩 중...</div>
        </div>
        <div class="row">
          <a class="btn" href="/">업로드</a>
          <a class="btn btnP" href="/dashboard">새로고침</a>
        </div>
      </div>

      <div id="warnBox" class="warn" style="display:none"></div>

      <div class="filters">
        <div style="grid-column: span 2">
          <label>기간 기준</label>
          <select id="basis">
            <option value="workMonth">근무월 기준</option>
            <option value="payMonth">지급월 기준</option>
          </select>
        </div>
        <div style="grid-column: span 2">
          <label>시작월</label>
          <select id="startMonth"></select>
        </div>
        <div style="grid-column: span 2">
          <label>종료월</label>
          <select id="endMonth"></select>
        </div>

        <div style="grid-column: span 3">
          <label>본부</label>
          <select id="hq"></select>
        </div>
        <div style="grid-column: span 3">
          <label>유형군</label>
          <select id="type"></select>
        </div>

        <div style="grid-column: span 6">
          <label>이름 검색</label>
          <input type="text" id="nameQ" placeholder="예: 권오준" />
        </div>
      </div>

      <div class="row" style="margin-top:10px">
        <div class="tog"><input type="checkbox" id="excludeAnom" checked/> <span class="small">이상치 제외</span></div>
        <div class="tog"><input type="checkbox" id="excludeRole"/> <span class="small">직책수당 제외</span></div>
        <div class="tog"><input type="checkbox" id="weekendOnly"/> <span class="small">휴일/출동만</span></div>

        <div style="flex:1"></div>

        <span class="small">지표</span>
        <select id="metric" style="width:150px">
          <option value="count">건수</option>
          <option value="days">당직일수</option>
          <option value="pay">당직수당</option>
        </select>

        <span class="small">TOP</span>
        <select id="topN" style="width:110px">
          <option>5</option><option selected>10</option><option>15</option><option>20</option><option>30</option><option>50</option>
        </select>

        <button class="btn" id="resetBtn">필터 초기화</button>
      </div>

      <div class="chips" id="chips"></div>

      <div class="kpiGrid" id="kpis"></div>

      <div class="grid2">
        <div class="card">
          <div class="panelTitle">유형별 누적</div>
          <div class="small">선택한 지표 기준</div>
          <canvas id="typeChart" height="140"></canvas>
        </div>
        <div class="card">
          <div class="panelTitle">개인별 TOP</div>
          <div class="small">선택한 지표 기준</div>
          <canvas id="personChart" height="140"></canvas>
        </div>
      </div>

      <div class="card" style="margin-top:12px">
        <div class="panelTitle">기간-유형별 건수 매트릭스</div>
        <div class="small">행=기간, 열=유형군</div>
        <div class="tblWrap" style="margin-top:10px">
          <table id="heatTbl"></table>
        </div>
      </div>

      <div class="grid2" style="margin-top:12px">
        <div class="card">
          <div class="panelTitle">개인별 누적 테이블 (상위 30)</div>
          <div class="tblWrap" style="margin-top:10px">
            <table id="personTbl"></table>
          </div>
        </div>
        <div class="card">
          <div class="panelTitle">최근 상세이력 (최대 25)</div>
          <div class="tblWrap" style="margin-top:10px">
            <table id="recentTbl"></table>
          </div>
        </div>
      </div>

    </div>

    <!-- Right panel -->
    <div class="card">
      <div class="panelTitle">요청사항 (추천 버튼 포함)</div>
      <div class="mut" style="margin-top:6px">
        아래에 적고 <b>적용</b>을 누르면 일부 설정이 바로 바뀝니다. (컴맹용)<br/>
        지원: <b>TOP N</b>, <b>기간(YYYY-MM ~ YYYY-MM)</b>, <b>휴일만</b>, <b>직책수당 제외</b>
      </div>

      <div style="margin-top:10px">
        <label>요청사항 입력</label>
        <textarea id="cmdBox" rows="6" placeholder="예:
개인별 TOP 20으로 변경
휴일만 보기
2025-03 ~ 2025-08
직책수당 제외"></textarea>
      </div>

      <div class="row" style="margin-top:10px">
        <button class="btn btnP" id="applyBtn">적용</button>
        <button class="btn" id="undoBtn">되돌리기</button>
        <button class="btn" id="clearBtn">입력 지우기</button>
      </div>

      <div class="row" style="margin-top:10px">
        <button class="chip" data-add="개인별 TOP 20으로 변경">개인별 TOP 20</button>
        <button class="chip" data-add="휴일만 보기">휴일만</button>
        <button class="chip" data-add="직책수당 제외">직책수당 제외</button>
        <button class="chip" data-add="2025-03 ~ 2025-08">기간 예시</button>
      </div>

      <div style="margin-top:12px" class="panelTitle">적용 내역</div>
      <div class="log" id="logBox" style="margin-top:8px"></div>

      <div class="mut" style="margin-top:10px">
        ⚠️ 인터넷 공개(A안)이므로 누구나 업로드 가능 상태입니다.<br/>
        (원하면 다음 단계에서 “업로드만 비번”으로 바꿀 수 있어요.)
      </div>
    </div>
  </div>

</div>

<script>
let META = {title:"로딩 중..."}; 
let RAW = [];
let META2 = {uploaded_at:null, filename:null};

let typeChart = null;
let personChart = null;

const state = {
  basis: "workMonth",
  startMonth: "",
  endMonth: "",
  hq: "ALL",
  type: "ALL",
  nameQ: "",
  excludeAnom: true,
  excludeRole: false,
  weekendOnly: false,
  metric: "count",
  topN: 10
};

let history = []; // snapshots for undo

function won(n){ return Math.round(n||0).toLocaleString()+"원"; }
function num(v){ 
  if(v==null) return 0;
  const s = String(v).replace(/,/g,'');
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function uniq(arr){ return Array.from(new Set(arr)); }

function heatColor(v,max){
  if(!max) return "rgba(37,99,235,0.05)";
  const a = 0.08 + 0.72*(v/max);
  return `rgba(37,99,235,${a.toFixed(3)})`;
}

function snapshot(){
  return JSON.parse(JSON.stringify(state));
}
function pushLog(ok, text, msg){
  const el = document.createElement("div");
  el.className = "logItem " + (ok ? "logOk":"logInfo");
  el.innerHTML = `<div style="font-weight:900;font-size:12px">${ok?"성공":"안내"} · ${escapeHtml(text)}</div>
                  <div style="font-size:12px;color:#334155;margin-top:4px;line-height:1.4">${escapeHtml(msg)}</div>`;
  const box = document.getElementById("logBox");
  box.prepend(el);
}

function escapeHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;")
    .replaceAll('"',"&quot;").replaceAll("'","&#039;");
}

function setSelectOptions(id, values, allLabel){
  const sel = document.getElementById(id);
  sel.innerHTML = "";
  if(allLabel){
    const opt = document.createElement("option");
    opt.value = "ALL"; opt.textContent = allLabel;
    sel.appendChild(opt);
  }
  for(const v of values){
    const opt = document.createElement("option");
    opt.value = v; opt.textContent = v;
    sel.appendChild(opt);
  }
}

function bindUI(){
  const basis = document.getElementById("basis");
  const startMonth = document.getElementById("startMonth");
  const endMonth = document.getElementById("endMonth");
  const hq = document.getElementById("hq");
  const type = document.getElementById("type");
  const nameQ = document.getElementById("nameQ");
  const excludeAnom = document.getElementById("excludeAnom");
  const excludeRole = document.getElementById("excludeRole");
  const weekendOnly = document.getElementById("weekendOnly");
  const metric = document.getElementById("metric");
  const topN = document.getElementById("topN");
  const resetBtn = document.getElementById("resetBtn");

  basis.addEventListener("change", ()=>{ state.basis = basis.value; initMonths(); renderAll(); });
  startMonth.addEventListener("change", ()=>{ state.startMonth = startMonth.value; renderAll(); });
  endMonth.addEventListener("change", ()=>{ state.endMonth = endMonth.value; renderAll(); });
  hq.addEventListener("change", ()=>{ state.hq = hq.value; renderAll(); });
  type.addEventListener("change", ()=>{ state.type = type.value; renderAll(); });
  nameQ.addEventListener("input", ()=>{ state.nameQ = nameQ.value; renderAll(); });
  excludeAnom.addEventListener("change", ()=>{ state.excludeAnom = excludeAnom.checked; renderAll(); });
  excludeRole.addEventListener("change", ()=>{ state.excludeRole = excludeRole.checked; renderAll(); });
  weekendOnly.addEventListener("change", ()=>{ state.weekendOnly = weekendOnly.checked; renderAll(); });

  metric.addEventListener("change", ()=>{ state.metric = metric.value; renderAll(); });
  topN.addEventListener("change", ()=>{ state.topN = Number(topN.value); renderAll(); });

  resetBtn.addEventListener("click", ()=>{
    state.basis = "workMonth";
    state.hq = "ALL";
    state.type = "ALL";
    state.nameQ = "";
    state.excludeAnom = true;
    state.excludeRole = false;
    state.weekendOnly = false;
    state.metric = "count";
    state.topN = 10;
    applyStateToUI();
    initMonths();
    renderAll();
  });

  // command box
  const cmdBox = document.getElementById("cmdBox");
  document.getElementById("applyBtn").addEventListener("click", ()=>{
    const txt = cmdBox.value.trim();
    if(!txt){ pushLog(false,"(빈 입력)","입력한 내용이 없습니다."); return; }
    history.unshift(snapshot());
    applyCommand(txt);
    applyStateToUI();
    renderAll();
  });
  document.getElementById("undoBtn").addEventListener("click", ()=>{
    const prev = history.shift();
    if(!prev){ pushLog(false,"되돌리기","되돌릴 내역이 없습니다."); return; }
    Object.assign(state, prev);
    applyStateToUI();
    initMonths();
    renderAll();
    pushLog(true,"되돌리기","이전 상태로 복원했습니다.");
  });
  document.getElementById("clearBtn").addEventListener("click", ()=>{ cmdBox.value=""; });

  document.querySelectorAll("[data-add]").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const add = btn.getAttribute("data-add");
      cmdBox.value = cmdBox.value ? (cmdBox.value + "\n" + add) : add;
    });
  });
}

function applyCommand(text){
  const t = text;
  let okAny = false;

  // TOP N
  const mTop = t.match(/TOP\s*(\d+)/i);
  if(mTop){
    const n = Number(mTop[1]);
    if(n>=1 && n<=100){ state.topN = n; okAny = true; pushLog(true, `TOP ${n}`, `개인별 TOP을 ${n}으로 변경했습니다.`); }
  }

  // 기간 YYYY-MM ~ YYYY-MM
  const mRange = t.match(/(20\d{2}-\d{2})\s*~\s*(20\d{2}-\d{2})/);
  if(mRange){
    state.startMonth = mRange[1];
    state.endMonth = mRange[2];
    okAny = true;
    pushLog(true, `${mRange[1]} ~ ${mRange[2]}`, "기간을 변경했습니다.");
  }

  if(t.includes("휴일만")){
    state.weekendOnly = true;
    okAny = true;
    pushLog(true, "휴일만", "휴일/출동만 보기로 설정했습니다.");
  }

  if(t.includes("직책수당 제외")){
    state.excludeRole = true;
    okAny = true;
    pushLog(true, "직책수당 제외", "직책수당을 제외하도록 설정했습니다.");
  }

  if(!okAny){
    pushLog(false, text, "이 대시보드에서 지원하지 않는 요청입니다. (TOP N / 기간 / 휴일만 / 직책수당 제외만 지원)");
  }
}

function applyStateToUI(){
  document.getElementById("basis").value = state.basis;
  document.getElementById("hq").value = state.hq;
  document.getElementById("type").value = state.type;
  document.getElementById("nameQ").value = state.nameQ;
  document.getElementById("excludeAnom").checked = state.excludeAnom;
  document.getElementById("excludeRole").checked = state.excludeRole;
  document.getElementById("weekendOnly").checked = state.weekendOnly;
  document.getElementById("metric").value = state.metric;
  document.getElementById("topN").value = String(state.topN);
  // start/end은 initMonths 후에 세팅
}

function initMonths(){
  const months = (state.basis==="workMonth") ? (META.monthsWork||[]) : (META.monthsPay||[]);
  const startSel = document.getElementById("startMonth");
  const endSel = document.getElementById("endMonth");
  startSel.innerHTML = "";
  endSel.innerHTML = "";
  for(const m of months){
    const o1 = document.createElement("option"); o1.value=m; o1.textContent=m;
    const o2 = document.createElement("option"); o2.value=m; o2.textContent=m;
    startSel.appendChild(o1); endSel.appendChild(o2);
  }
  if(!state.startMonth) state.startMonth = months[0] || "";
  if(!state.endMonth) state.endMonth = months[months.length-1] || "";

  // if user entered months not in list, keep but try best
  startSel.value = state.startMonth || (months[0]||"");
  endSel.value = state.endMonth || (months[months.length-1]||"");
}

function filteredRows(){
  let rows = RAW.slice();

  if(state.excludeAnom) rows = rows.filter(r => !r.isAnomaly);
  if(state.excludeRole) rows = rows.filter(r => r.typeGroup !== "직책수당");
  if(state.weekendOnly) rows = rows.filter(r => (r.typeGroup==="휴일" || r.typeGroup==="휴일출동"));

  const b = state.basis;
  if(state.startMonth) rows = rows.filter(r => (r[b]||"") >= state.startMonth);
  if(state.endMonth) rows = rows.filter(r => (r[b]||"") <= state.endMonth);

  if(state.hq !== "ALL") rows = rows.filter(r => (r.hq||"-") === state.hq);
  if(state.type !== "ALL") rows = rows.filter(r => (r.typeGroup||"-") === state.type);

  const q = (state.nameQ||"").trim();
  if(q) rows = rows.filter(r => String(r.name||"").includes(q));

  return rows;
}

function computeKPIs(rows){
  const recordCount = rows.length;
  const personCount = uniq(rows.map(r=>r.name).filter(Boolean)).length;
  const totalDays = rows.reduce((a,r)=>a + num(r.days), 0);
  const totalPay = rows.reduce((a,r)=>a + num(r.pay), 0);
  const avgPay = personCount ? totalPay / personCount : 0;
  return {recordCount, personCount, totalDays, totalPay, avgPay};
}

function aggByType(rows){
  const map = {};
  for(const r of rows){
    const k = r.typeGroup || "기타";
    if(!map[k]) map[k] = {typeGroup:k, count:0, days:0, pay:0};
    map[k].count += 1;
    map[k].days += num(r.days);
    map[k].pay += num(r.pay);
  }
  return Object.values(map).sort((a,b)=> (b[state.metric]||0) - (a[state.metric]||0));
}

function aggByPerson(rows){
  const map = {};
  const typeMap = {};
  for(const r of rows){
    const n = r.name || "-";
    if(!map[n]) map[n] = {name:n, count:0, days:0, pay:0, repType:"-"};
    map[n].count += 1;
    map[n].days += num(r.days);
    map[n].pay += num(r.pay);

    const tg = r.typeGroup || "기타";
    typeMap[n] = typeMap[n] || {};
    typeMap[n][tg] = (typeMap[n][tg]||0) + 1;
  }
  // repType
  for(const n of Object.keys(typeMap)){
    let best="-", bestN=-1;
    for(const [tg,c] of Object.entries(typeMap[n])){
      if(c>bestN){ bestN=c; best=tg; }
    }
    map[n].repType = best;
  }
  const arr = Object.values(map).sort((a,b)=> (b[state.metric]||0) - (a[state.metric]||0));
  return arr;
}

function heatMatrix(rows){
  const b = state.basis;
  const months = uniq(rows.map(r=>r[b]).filter(Boolean)).sort();
  const types = uniq(rows.map(r=>r.typeGroup).filter(Boolean)).sort();
  const table = {};
  for(const mo of months){
    table[mo] = {};
    for(const t of types) table[mo][t] = 0;
  }
  for(const r of rows){
    const mo = r[b]; const t = r.typeGroup;
    if(!mo || !t) continue;
    table[mo][t] = (table[mo][t]||0) + 1;
  }
  let max=0;
  for(const mo of months) for(const t of types) max = Math.max(max, table[mo][t]||0);
  return {months, types, table, max};
}

function renderChips(){
  const chips = [];
  chips.push(state.basis==="workMonth" ? "근무월 기준" : "지급월 기준");
  if(state.startMonth && state.endMonth) chips.push(`${state.startMonth}~${state.endMonth}`);
  if(state.hq!=="ALL") chips.push(`본부: ${state.hq}`);
  if(state.type!=="ALL") chips.push(`유형: ${state.type}`);
  if((state.nameQ||"").trim()) chips.push(`이름: ${(state.nameQ||"").trim()}`);
  if(state.excludeAnom) chips.push("이상치 제외");
  if(state.excludeRole) chips.push("직책수당 제외");
  if(state.weekendOnly) chips.push("휴일만");

  document.getElementById("chips").innerHTML = chips.map(c=>`<span class="chip">${escapeHtml(c)}</span>`).join("");
}

function renderKPIs(k){
  const items = [
    ["레코드(필터적용)", k.recordCount.toLocaleString()],
    ["인원(필터적용)", k.personCount.toLocaleString()],
    ["총 일수(필터적용)", (Math.round(k.totalDays*10)/10).toLocaleString()],
    ["총 수당(필터적용)", won(k.totalPay)],
  ];
  document.getElementById("kpis").innerHTML = items.map(([l,v]) =>
    `<div class="kpi"><div class="l">${escapeHtml(l)}</div><div class="v">${escapeHtml(v)}</div>
     ${l.includes("총 수당") ? `<div class="mut">인당 평균: ${won(k.avgPay)}</div>` : ``}
     </div>`
  ).join("");
}

function renderTypeChart(typeAgg){
  const labels = typeAgg.map(x=>x.typeGroup);
  const vals = typeAgg.map(x=>x[state.metric]||0);

  if(typeChart) typeChart.destroy();
  typeChart = new Chart(document.getElementById("typeChart"), {
    type: "bar",
    data: { labels, datasets: [{ label: state.metric, data: vals, backgroundColor: "#2563eb", borderRadius: 10 }] },
    options: { responsive:true, plugins:{legend:{display:false}}, scales:{y:{ticks:{color:"#334155"}}, x:{ticks:{color:"#334155"}}} }
  });
}

function renderPersonChart(personAgg){
  const top = personAgg.slice(0, state.topN);
  const labels = top.map(x=>x.name);
  const vals = top.map(x=>x[state.metric]||0);

  if(personChart) personChart.destroy();
  personChart = new Chart(document.getElementById("personChart"), {
    type: "bar",
    data: { labels, datasets: [{ label: state.metric, data: vals, backgroundColor: "#60a5fa", borderRadius: 10 }] },
    options: { responsive:true, plugins:{legend:{display:false}}, scales:{x:{ticks:{display:false}}, y:{ticks:{color:"#334155"}}} }
  });
}

function renderHeatTbl(hm){
  const {months, types, table, max} = hm;
  if(!months.length){
    document.getElementById("heatTbl").innerHTML = `<tr><td style="padding:18px;color:#64748b">데이터가 없습니다. 업로드 또는 필터를 확인하세요.</td></tr>`;
    return;
  }
  let html = `<thead><tr><th style="min-width:120px">기간</th>`;
  for(const t of types) html += `<th style="text-align:center">${escapeHtml(t)}</th>`;
  html += `</tr></thead><tbody>`;
  for(const mo of months){
    html += `<tr><td class="sticky" style="font-weight:900">${escapeHtml(mo)}</td>`;
    for(const t of types){
      const v = (table[mo] && table[mo][t]) ? table[mo][t] : 0;
      html += `<td style="text-align:center;background:${heatColor(v,max)}">${v}</td>`;
    }
    html += `</tr>`;
  }
  html += `</tbody>`;
  document.getElementById("heatTbl").innerHTML = html;
}

function renderPersonTbl(personAgg){
  const top30 = personAgg.slice(0, 30);
  if(!top30.length){
    document.getElementById("personTbl").innerHTML = `<tr><td style="padding:18px;color:#64748b">데이터가 없습니다.</td></tr>`;
    return;
  }
  let html = `<thead><tr>
    <th>이름</th><th class="r">건수</th><th class="r">일수</th><th class="r">수당</th><th>대표유형</th>
  </tr></thead><tbody>`;
  top30.forEach((p,idx)=>{
    html += `<tr>
      <td><span style="display:inline-flex;align-items:center;gap:8px">
        <span style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:999px;background:#eef2ff;color:#3730a3;font-weight:900;font-size:12px">${idx+1}</span>
        <span style="font-weight:900">${escapeHtml(p.name)}</span>
      </span></td>
      <td class="r">${(p.count||0).toLocaleString()}</td>
      <td class="r">${(Math.round((p.days||0)*10)/10).toLocaleString()}</td>
      <td class="r">${won(p.pay||0)}</td>
      <td>${escapeHtml(p.repType||"-")}</td>
    </tr>`;
  });
  html += `</tbody>`;
  document.getElementById("personTbl").innerHTML = html;
}

function renderRecentTbl(rows){
  const r = rows.slice().sort((a,b)=> String(b.workDate||"").localeCompare(String(a.workDate||""))).slice(0,25);
  if(!r.length){
    document.getElementById("recentTbl").innerHTML = `<tr><td style="padding:18px;color:#64748b">데이터가 없습니다.</td></tr>`;
    return;
  }
  let html = `<thead><tr>
    <th>일자</th><th>이름</th><th>본부</th><th>유형군</th><th class="r">일수</th><th class="r">수당</th>
  </tr></thead><tbody>`;
  r.forEach(x=>{
    html += `<tr>
      <td>${escapeHtml(x.workDate||"-")}</td>
      <td style="font-weight:900">${escapeHtml(x.name||"-")}</td>
      <td>${escapeHtml(x.hq||"-")}</td>
      <td>${escapeHtml(x.typeGroup||"-")}</td>
      <td class="r">${(Math.round(num(x.days)*10)/10).toLocaleString()}</td>
      <td class="r">${won(num(x.pay))}</td>
    </tr>`;
  });
  html += `</tbody>`;
  document.getElementById("recentTbl").innerHTML = html;
}

function renderWarn(){
  const warn = document.getElementById("warnBox");
  const anyAnom = META.anomalyCount || 0;
  const missDate = META.missingDateCount || 0;
  if(anyAnom || missDate){
    warn.style.display = "block";
    warn.textContent = `주의: 이상치 ${anyAnom}건 / 날짜누락 ${missDate}건이 감지되었습니다. (필터에서 '이상치 제외'로 숨길 수 있음)`;
  }else{
    warn.style.display = "none";
  }
}

function renderHeader(){
  document.getElementById("title").textContent = META.title || "당직 대시보드";
  const up = META2.uploaded_at ? META2.uploaded_at : "-";
  const fn = META2.filename ? META2.filename : "-";
  document.getElementById("metaLine").textContent = `최신 업로드: ${up} / 파일: ${fn} / 레코드 ${META.recordCount||0} · 인원 ${META.personCount||0}`;
}

function renderAll(){
  const rows = filteredRows();
  renderChips();
  renderWarn();

  const k = computeKPIs(rows);
  renderKPIs(k);

  const typeAgg = aggByType(rows);
  renderTypeChart(typeAgg);

  const personAgg = aggByPerson(rows);
  renderPersonChart(personAgg);
  renderPersonTbl(personAgg);

  const hm = heatMatrix(rows);
  renderHeatTbl(hm);

  renderRecentTbl(rows);
}

async function boot(){
  const data = await fetch("/data/latest").then(r=>r.json());
  const meta = await fetch("/data/latest_meta").then(r=>r.json());

  META = data.META || {title:"데이터 없음"};
  RAW = data.RAW || [];
  META2 = meta || META2;

  // init filter options
  setSelectOptions("hq", (META.hqs||[]), "전체");
  setSelectOptions("type", ["휴일","휴일출동","평일","직책수당","기타"], "전체");

  bindUI();
  applyStateToUI();
  initMonths();
  renderHeader();
  renderAll();

  // if no data
  if(!RAW.length){
    pushLog(false, "안내", "아직 데이터가 없습니다. 업로드 페이지(/)에서 엑셀을 올려주세요.");
  }
}

boot();
</script>
</body>
</html>
"""

# -----------------------------
# Routes
# -----------------------------
@app.get("/", response_class=HTMLResponse)
def home():
    return HTMLResponse(UPLOAD_HTML)

@app.get("/dashboard", response_class=HTMLResponse)
def dashboard():
    return HTMLResponse(DASHBOARD_HTML)

@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    fn = (file.filename or "").lower()
    if not (fn.endswith(".xlsx") or fn.endswith(".xls")):
        raise HTTPException(status_code=400, detail="엑셀(.xlsx/.xls) 파일만 업로드할 수 있습니다.")

    try:
        payload = excel_to_payload(file.file)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 읽기 실패: {e}")

    global LATEST, LATEST_META
    LATEST = payload
    LATEST_META = {
        "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "filename": file.filename
    }

    # 업로드 후 바로 대시보드로 가고 싶으면 아래처럼 HTML 반환
    return HTMLResponse(f"""
      <meta charset="utf-8"/>
      <meta http-equiv="refresh" content="0; url=/dashboard"/>
      업로드 완료! /dashboard로 이동합니다...
    """)

@app.get("/data/latest")
def data_latest():
    return JSONResponse(LATEST)

@app.get("/data/latest_meta")
def data_latest_meta():
    return JSONResponse(LATEST_META)
