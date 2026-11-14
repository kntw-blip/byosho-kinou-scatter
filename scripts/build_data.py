"""令和6年度 病床機能報告オープンデータ(施設票・様式1病棟票)を病院単位に集計し、
data/hospitals.json を生成する。

使い方: python scripts/build_data.py raw/ data/
  raw/ に施設票(facility.xlsx)と病棟票(ward_*.xlsx)、留意事項(notes.zip)を置く。
"""
import glob, json, os, re, sys, zipfile, io, datetime
from collections import Counter, defaultdict
from openpyxl import load_workbook

RAW, OUT = sys.argv[1], sys.argv[2]
HEADER_ROW, DATA_ROW = 4, 6          # 0始まり。5行目が項目名、7行目からデータ
DAYS = 366                           # 令和5年4月〜令和6年3月(うるう年を含む)

PREFS = ("北海道 青森県 岩手県 宮城県 秋田県 山形県 福島県 茨城県 栃木県 群馬県 埼玉県 千葉県 東京都 神奈川県 "
         "新潟県 富山県 石川県 福井県 山梨県 長野県 岐阜県 静岡県 愛知県 三重県 滋賀県 京都府 大阪府 兵庫県 "
         "奈良県 和歌山県 鳥取県 島根県 岡山県 広島県 山口県 徳島県 香川県 愛媛県 高知県 福岡県 佐賀県 長崎県 "
         "熊本県 大分県 宮崎県 鹿児島県 沖縄県").split()
JOBS = ["医師", "歯科医師", "看護師", "准看護師", "看護補助者", "助産師", "理学療法士", "作業療法士",
        "言語聴覚士", "薬剤師", "診療放射線技師", "臨床検査技師", "臨床工学技士", "管理栄養士", "救急救命士"]
FUNCS = ["高度急性期", "急性期", "回復期", "慢性期"]
# 1職種の常勤換算がこの人数を超える報告は入力誤りの疑いが強いとして印を付ける(値は変更しない)。
# 参考: 全国で最も多い大学病院の看護師でも約1,400人。
SUSPECT_FTE = 5000


def norm(s):
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def read(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    header, rows = None, []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == HEADER_ROW:
            header = [norm(v) for v in row]
        elif i >= DATA_ROW and row and row[0] is not None:
            rows.append(row)
    return header, rows


def col(header, label, prefix=False):
    """項目名から列番号を得る(同名が複数ある場合は最初の列)。"""
    label = norm(label)
    for i, h in enumerate(header):
        if h == label or (prefix and h.startswith(label)):
            return i
    raise KeyError(f"列が見つかりません: {label}")


def num(v):
    """数値化。秘匿(*)・未報告・空欄は None。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def add(*vals):
    vals = [v for v in vals if v is not None]
    return sum(vals) if vals else None


def text(v):
    return "" if v is None else str(v).strip()


# 各ファイルの項目名一覧を data/columns.json に書き出す(項目確認用)
def dump_columns():
    out = {}
    for name in ("facility.xlsx", "form2.xlsx", "ward_892.xlsx"):
        path = os.path.join(RAW, name)
        if not os.path.exists(path):
            continue
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i > 8:
                break
            rows.append([None if v is None else str(v)[:60] for v in row])
        out[name] = {"labels": [norm(v) for v in rows[HEADER_ROW]], "rows": rows, "sheets": wb.sheetnames}
    os.makedirs(OUT, exist_ok=True)
    json.dump(out, open(os.path.join(OUT, "columns.json"), "w"), ensure_ascii=False)


dump_columns()

report = {"generated": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")}

# ---------- 施設票 ----------
fh, frows = read(os.path.join(RAW, "facility.xlsx"))
F = lambda label, **k: col(fh, label, **k)
c_id = F("オープンデータ医療機関コード", prefix=True)
fc = {
    "kind": 0, "name": F("医療機関名"), "pref": F("都道府県コード"), "area": F("二次医療圏名"),
    "city": F("市区町村名称"), "owner": F("設置主体", prefix=True), "dpc": F("DPC医療機関群の種類"),
    "tokutei": F("特定機能病院の承認の有無"), "shien": F("地域医療支援病院の承認の有無"),
    "er3": F("三次救急医療施設の認定の有無"), "er2": F("二次救急医療施設の認定の有無"),
}
fnum = {
    "amb": F("救急車の受入件数"), "holiday": F("休日に受診した患者延べ数"),
    "night": F("夜間・時間外に受診した患者延べ数"), "robot": F("内視鏡手術用支援機器"),
}
ct_cols = [i for i, h in enumerate(fh) if h.startswith("CT_")]
mri_cols = [i for i, h in enumerate(fh) if h.startswith("MRI_")]
job_cols = {j: (F(f"施設全体_{j}_常勤"), F(f"施設全体_{j}_非常勤")) for j in JOBS}
report["facility_rows"] = len(frows)
report["kind_counts"] = Counter(text(r[0]) for r in frows)
report["ct_cols"] = [fh[i] for i in ct_cols]
report["mri_cols"] = [fh[i] for i in mri_cols]
report["staff_nonnumeric_samples"] = Counter()

hosp = {}
for r in frows:
    hid = text(r[c_id])
    h = {k: text(r[i]) for k, i in fc.items()}
    for k, i in fnum.items():
        h[k] = num(r[i])
    h["ct"] = add(*[num(r[i]) for i in ct_cols])
    h["mri"] = add(*[num(r[i]) for i in mri_cols])
    full_all = part_all = None
    for j, (cf, cp) in job_cols.items():
        full, part = num(r[cf]), num(r[cp])
        for raw, val in ((r[cf], full), (r[cp], part)):
            if val is None:
                report["staff_nonnumeric_samples"][text(raw)] += 1
        h["fte_" + j] = add(full, part)
        full_all, part_all = add(full_all, full), add(part_all, part)
    h["fte_full"], h["fte_part"] = full_all, part_all
    h["fte_all"] = add(full_all, part_all)
    h["fte_nurse"] = add(h["fte_看護師"], h["fte_准看護師"], h["fte_助産師"])
    h["w"] = defaultdict(float)
    h["seen"] = set()
    hosp[hid] = h

# ---------- 様式1 病棟票 ----------
func_values, unmatched, ward_rows = Counter(), set(), 0
for path in sorted(glob.glob(os.path.join(RAW, "ward_*.xlsx"))):
    wh, wrows = read(path)
    W = lambda label, **k: col(wh, label, **k)
    w_id = W("オープンデータ医療機関コード", prefix=True)
    w_func = W("保有する病棟と機能区分の選択状況（2024", prefix=True)
    wn = {
        "general": W("一般病床_許可病床"), "ryoyo": W("療養病床_許可病床"),
        "max_g": W("一般病床_最大使用病床数"), "max_r": W("療養病床_最大使用病床数"),
        "adm": W("新規入棟患者数（年間）"), "days": W("在棟患者延べ数（年間）"),
        "dis": W("退棟患者数（年間）"), "birth": W("分娩件数", prefix=True),
    }
    for r in wrows:
        ward_rows += 1
        hid = text(r[w_id])
        func = text(r[w_func])
        func_values[func] += 1
        h = hosp.get(hid)
        if h is None:
            unmatched.add(hid)
            continue
        v = {k: num(r[i]) for k, i in wn.items()}
        beds = (v["general"] or 0) + (v["ryoyo"] or 0)
        w = h["w"]
        w["wards"] += 1
        w["general"] += v["general"] or 0
        w["ryoyo"] += v["ryoyo"] or 0
        w["maxuse"] += (v["max_g"] or 0) + (v["max_r"] or 0)
        w["f_" + (func if func in FUNCS else "その他")] += beds
        for k in ("adm", "days", "dis", "birth"):
            if v[k] is not None:
                w[k] += v[k]
                h["seen"].add(k)
            elif text(r[wn[k]]) == "*":
                h["seen"].add(k + "_masked")
report["ward_rows"] = ward_rows
report["function_values"] = func_values
report["ward_hospitals_not_in_facility_sheet"] = len(unmatched)

# ---------- 病院単位レコード ----------
G_STAFF = "職員数(常勤+非常勤の常勤換算)"
FIELDS = [  # key, 表示名, グループ, 単位, 小数桁
    ("beds", "許可病床数(一般+療養)", "病床", "床", 0),
    ("beds_f0", "高度急性期の病床数", "病床", "床", 0),
    ("beds_f1", "急性期の病床数", "病床", "床", 0),
    ("beds_f2", "回復期の病床数", "病床", "床", 0),
    ("beds_f3", "慢性期の病床数", "病床", "床", 0),
    ("beds_other", "休棟中等の病床数", "病床", "床", 0),
    ("beds_general", "一般病床の許可病床数", "病床", "床", 0),
    ("beds_ryoyo", "療養病床の許可病床数", "病床", "床", 0),
    ("beds_maxuse", "最大使用病床数(病棟計)", "病床", "床", 0),
    ("wards", "病棟数", "病床", "棟", 0),
    ("fte_all", "全職種計", G_STAFF, "人", 1),
    ("fte_nurse", "看護職員計(看護師・准看護師・助産師)", G_STAFF, "人", 1),
] + [("fte_" + j, j, G_STAFF, "人", 1) for j in JOBS] + [
    ("fte_full", "全職種計のうち常勤", G_STAFF, "人", 1),
    ("fte_part", "全職種計のうち非常勤(常勤換算)", G_STAFF, "人", 1),
    ("adm", "新規入棟患者数(年間)", "診療実績", "人", 0),
    ("days", "在棟患者延べ数(年間)", "診療実績", "人日", 0),
    ("dis", "退棟患者数(年間)", "診療実績", "人", 0),
    ("occ", "病床稼働率(在棟延べ数÷許可病床×366日)", "診療実績", "%", 1),
    ("alos", "平均在棟日数", "診療実績", "日", 1),
    ("birth", "分娩件数(年間)", "診療実績", "件", 0),
    ("amb", "救急車の受入件数(年間)", "診療実績", "件", 0),
    ("holiday", "休日に受診した患者延べ数(年間)", "診療実績", "人", 0),
    ("night", "夜間・時間外に受診した患者延べ数(年間)", "診療実績", "人", 0),
    ("ct", "CT台数", "医療機器", "台", 0),
    ("mri", "MRI台数", "医療機器", "台", 0),
    ("robot", "内視鏡手術用支援機器の台数", "医療機器", "台", 0),
]
CATS = [("pref", "都道府県"), ("area", "二次医療圏"), ("func", "主機能"), ("owner", "設置主体"),
        ("dpc", "DPC医療機関群"), ("tokutei", "特定機能病院"), ("shien", "地域医療支援病院"),
        ("er3", "三次救急"), ("er2", "二次救急")]

records, no_ward = [], 0
for hid, h in hosp.items():
    w = h["w"]
    if not w.get("wards"):
        no_ward += 1
        continue
    rec = {"id": hid, "name": h["name"], "city": h["city"]}
    try:
        rec["pref"] = PREFS[int(h["pref"]) - 1]
    except (ValueError, IndexError):
        rec["pref"] = "不明"
    for k in ("area", "owner", "dpc", "tokutei", "shien", "er3", "er2"):
        rec[k] = h[k] or "不明"
    fb = [w.get("f_" + f, 0) for f in FUNCS]
    rec["beds"] = w["general"] + w["ryoyo"]
    for i in range(4):
        rec[f"beds_f{i}"] = fb[i]
    rec["beds_other"] = w.get("f_その他", 0)
    rec["func"] = FUNCS[fb.index(max(fb))] if max(fb) > 0 else "休棟中等"
    rec["beds_general"], rec["beds_ryoyo"] = w["general"], w["ryoyo"]
    rec["beds_maxuse"], rec["wards"] = w["maxuse"], w["wards"]
    for k in ("adm", "days", "dis", "birth"):
        rec[k] = w[k] if k in h["seen"] else None
    rec["occ"] = rec["days"] / (rec["beds"] * DAYS) * 100 if rec["days"] is not None and rec["beds"] > 0 else None
    flow = add(rec["adm"], rec["dis"])
    rec["alos"] = rec["days"] / (flow / 2) if rec["days"] and flow else None
    for key, *_ in FIELDS:
        if key not in rec:
            rec[key] = h.get(key)
    rec["suspect"] = int(any((h["fte_" + j] or 0) > SUSPECT_FTE for j in JOBS))
    records.append(rec)
records.sort(key=lambda r: r["id"])
report["hospitals_out"] = len(records)
report["hospitals_without_ward_rows"] = no_ward
report["main_function_counts"] = Counter(r["func"] for r in records)
report["null_counts"] = {k: sum(1 for r in records if r[k] is None) for k, *_ in FIELDS}
report["totals"] = {k: round(sum(r[k] or 0 for r in records), 1)
                    for k in ("beds", "beds_f0", "beds_f1", "beds_f2", "beds_f3", "beds_other",
                              "fte_all", "fte_医師", "fte_看護師")}
report["masked_ward_values"] = {k: sum(1 for h in hosp.values() if k + "_masked" in h["seen"])
                                for k in ("adm", "days", "dis", "birth")}
report["suspect"] = [(r["name"], r["pref"], r["beds"], r["fte_all"]) for r in records if r["suspect"]]
report["sample"] = [r for r in records if "市立函館病院" in r["name"]][:1]

# 留意事項PDFから秘匿処理などの記述を抜き出す(確認用)
try:
    from pypdf import PdfReader
    z = zipfile.ZipFile(os.path.join(RAW, "notes.zip"))
    notes = []
    for n in z.namelist():
        if n.lower().endswith(".pdf"):
            txt = "\n".join(p.extract_text() or "" for p in PdfReader(io.BytesIO(z.read(n))).pages)
            notes += [l.strip() for l in txt.splitlines() if re.search(r"[*＊]|秘匿|未満|非常勤|常勤換算", l)]
    report["notes_excerpt"] = notes[:60]
except Exception as e:  # 確認用なので失敗しても続行
    report["notes_excerpt"] = f"error: {e}"


def rnd(v, d):
    if v is None:
        return None
    v = round(v, d)
    return int(v) if d == 0 or v == int(v) else v


cats = {}
cols = {k: [r[k] for r in records] for k in ("id", "name", "city", "suspect")}
for key, _ in CATS:
    present = {r[key] for r in records}
    order = FUNCS + ["休棟中等"] if key == "func" else PREFS if key == "pref" else sorted(present)
    order = [v for v in order if v in present] + sorted(present - set(order))
    idx = {v: i for i, v in enumerate(order)}
    cats[key] = order
    cols[key] = [idx[r[key]] for r in records]
for key, _, _, _, d in FIELDS:
    cols[key] = [rnd(r[key], d) for r in records]

os.makedirs(OUT, exist_ok=True)
json.dump({
    "meta": {"title": "令和6年度病床機能報告", "asof": "2024年7月1日時点(診療実績は2023年4月〜2024年3月)",
             "source": "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/open_data_00018.html",
             "generated": report["generated"], "n": len(records)},
    "fields": [dict(key=k, label=l, group=g, unit=u, dec=d) for k, l, g, u, d in FIELDS],
    "catLabels": dict(CATS), "cats": cats, "cols": cols,
}, open(os.path.join(OUT, "hospitals.json"), "w"), ensure_ascii=False, separators=(",", ":"))
json.dump(report, open(os.path.join(OUT, "build_report.json"), "w"), ensure_ascii=False, indent=1)
print("hospitals:", len(records))
