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
WARD_JOBS = ["看護師", "准看護師", "看護補助者", "助産師", "理学療法士", "作業療法士", "言語聴覚士",
             "薬剤師", "臨床工学技士", "管理栄養士", "救急救命士"]
# 様式2(具体的な医療の内容・年間合計)から軸に追加する主要項目: key, 公表データの項目名, 表示名
FORM2 = [
    ("ope", "手術総数", "手術総数"),
    ("ope_ga", "全身麻酔の手術総数", "全身麻酔の手術総数"),
    ("ope_cpb", "人工心肺を用いた手術", "人工心肺を用いた手術"),
    ("ope_vats", "胸腔鏡下手術", "胸腔鏡下手術"),
    ("ope_lap", "腹腔鏡下手術", "腹腔鏡下手術"),
    ("ope_robot", "内視鏡手術用支援機器手術", "内視鏡手術用支援機器(ロボット)手術"),
    ("ope_ca", "悪性腫瘍手術", "悪性腫瘍手術"),
    ("path", "病理組織標本作製", "病理組織標本作製"),
    ("rt", "放射線治療", "放射線治療"),
    ("chemo", "化学療法", "化学療法"),
    ("stroke", "超急性期脳卒中加算", "超急性期脳卒中加算"),
    ("evt", "脳血管内手術", "脳血管内手術"),
    ("pci", "経皮的冠動脈形成術", "経皮的冠動脈形成術(PCI)"),
    ("hr_birth", "ハイリスク分娩管理加算", "ハイリスク分娩管理加算"),
    ("er_add", "救急医療管理加算１及び２", "救急医療管理加算1・2"),
    ("er_night", "夜間休日救急搬送医学管理料", "夜間休日救急搬送医学管理料"),
    ("triage", "院内トリアージ実施料", "院内トリアージ実施料"),
    ("vent", "人工呼吸（５時間を超えた場合）", "人工呼吸(5時間超)"),
    ("hd", "人工腎臓、腹膜灌流", "人工腎臓・腹膜灌流"),
    ("dc1", "入退院支援加算１", "入退院支援加算1"),
    ("dementia", "認知症ケア加算１", "認知症ケア加算1"),
    ("reh", "疾患別リハビリテーション料", "疾患別リハビリテーション料(計)"),
    ("reh_heart", "心大血管疾患リハビリテーション料", "心大血管疾患リハビリテーション料"),
    ("reh_brain", "脳血管疾患等リハビリテーション料", "脳血管疾患等リハビリテーション料"),
    ("reh_disuse", "廃用症候群リハビリテーション料", "廃用症候群リハビリテーション料"),
    ("reh_ortho", "運動器リハビリテーション料", "運動器リハビリテーション料"),
    ("reh_resp", "呼吸器リハビリテーション料", "呼吸器リハビリテーション料"),
    ("reh_early", "早期リハビリテーション加算（リハビリテーション料）", "早期リハビリテーション加算"),
    ("dysph", "摂食機能療法", "摂食機能療法"),
]
G_F2 = "医療の内容(様式2・年間の算定件数)"


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


# dump_columns()  # 項目名を確認したいときに有効化

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

# ---------- 様式2 病棟票(年間合計) ----------
# 様式1と様式2では病棟コードの付番が異なるため、病棟は「医療機関コード+病棟名」で突合する。
# (医療機関コード, 病棟名) -> {key: 数値 or None(秘匿「*」・未報告)}。病院合計は医療機関コードで直接合算する。
form2, form2_dup, form2_hosp = {}, set(), defaultdict(lambda: {"wards": 0})
f2h, f2rows = read(os.path.join(RAW, "form2.xlsx"))
f2_hid = col(f2h, "オープンデータ医療機関コード", prefix=True)
f2_wname = col(f2h, "病棟名")
f2_month = col(f2h, "報告月")
f2_cols = [(k, col(f2h, lab)) for k, lab, _ in FORM2]
report["form2_rows"] = len(f2rows)
report["form2_month_values"] = Counter(text(r[f2_month]) for r in f2rows)
for r in f2rows:
    vals = {k: num(r[i]) for k, i in f2_cols}
    key = (text(r[f2_hid]), norm(r[f2_wname]))
    if key in form2:
        form2_dup.add(key)
    form2[key] = vals
    agg = form2_hosp[key[0]]
    agg["wards"] += 1
    for k, val in vals.items():
        if val is None:
            agg["m_" + k] = agg.get("m_" + k, 0) + 1
        else:
            agg[k] = agg.get(k, 0) + val
for key in form2_dup:   # 同名病棟が複数ある場合は突合しない
    del form2[key]
report["form2_wards"] = len(form2)
report["form2_duplicate_names"] = len(form2_dup)

# ---------- 様式1 病棟票 ----------
func_values, unmatched, ward_rows = Counter(), set(), 0
wards = []
for path in sorted(glob.glob(os.path.join(RAW, "ward_*.xlsx"))):
    wh, wrows = read(path)
    W = lambda label, **k: col(wh, label, **k)
    w_id = W("オープンデータ医療機関コード", prefix=True)
    w_func = W("保有する病棟と機能区分の選択状況（2024", prefix=True)
    w_ward, w_wname = W("病棟コード", prefix=True), W("病棟名")
    w_fee, w_dept = W("算定する入院基本料・特定入院料"), W("主とする診療科")
    w_jobs = {j: (W(f"{j}_常勤"), W(f"{j}_非常勤")) for j in WARD_JOBS}
    w_er = W("うち、予定外の救急医療入院の患者（年間）")
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
        f2 = form2.get((hid, norm(r[w_wname])))
        # 病棟単位のレコード
        wr = {"id": text(r[w_ward]), "hid": hid, "name": h["name"] + " " + text(r[w_wname]), "city": h["city"],
              "pref": h["pref"], "area": h["area"], "func": func if func in FUNCS else "休棟中等",
              "fee": text(r[w_fee]) or "不明", "dept": text(r[w_dept]) or "不明",
              "beds": beds, "beds_general": v["general"] or 0, "beds_ryoyo": v["ryoyo"] or 0,
              "beds_maxuse": (v["max_g"] or 0) + (v["max_r"] or 0),
              "adm": v["adm"], "days": v["days"], "dis": v["dis"], "adm_er": num(r[w_er])}
        tot = None
        for j, (cf, cp) in w_jobs.items():
            wr["fte_" + j] = add(num(r[cf]), num(r[cp]))
            tot = add(tot, wr["fte_" + j])
        wr["fte_all"] = tot
        wr["fte_nurse"] = add(wr["fte_看護師"], wr["fte_准看護師"], wr["fte_助産師"])
        wr["occ"] = v["days"] / (beds * DAYS) * 100 if v["days"] is not None and beds > 0 else None
        flow = add(v["adm"], v["dis"])
        wr["alos"] = v["days"] / (flow / 2) if v["days"] and flow else None
        for k, _, _ in FORM2:
            wr[k] = f2[k] if f2 is not None else None
        wr["suspect"] = int(any((wr["fte_" + j] or 0) > SUSPECT_FTE for j in WARD_JOBS))
        wards.append(wr)
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
] + [(k, lab, G_F2, "件", 0) for k, _, lab in FORM2]
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
    # 様式2: 秘匿のない病棟の合計。合計が0で秘匿病棟がある場合(=1件以上だが不明)と、様式2の報告がない場合は欠損
    agg = form2_hosp.get(hid)
    for k, _, _ in FORM2:
        tot, masked = (agg.get(k, 0), agg.get("m_" + k, 0)) if agg else (0, 0)
        rec[k] = None if agg is None or (masked and tot == 0) else tot
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
report["form2_hospitals_partly_masked"] = {k: sum(1 for a in form2_hosp.values() if a.get("m_" + k) and a.get(k, 0) > 0)
                                          for k, _, _ in FORM2}
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


def write(fname, recs, fields, catdefs, meta):
    cats = {}
    cols = {k: [r[k] for r in recs] for k in ("id", "name", "city", "suspect")}
    for key, _ in catdefs:
        present = {r[key] for r in recs}
        order = FUNCS + ["休棟中等"] if key == "func" else PREFS if key == "pref" else sorted(present)
        order = [v for v in order if v in present] + sorted(present - set(order))
        idx = {v: i for i, v in enumerate(order)}
        cats[key] = order
        cols[key] = [idx[r[key]] for r in recs]
    for key, _, _, _, d in fields:
        cols[key] = [rnd(r[key], d) for r in recs]
    meta = dict({"title": "令和6年度病床機能報告", "asof": "2024年7月1日時点(診療実績は2023年4月〜2024年3月)",
                 "source": "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/open_data_00018.html",
                 "generated": report["generated"], "n": len(recs)}, **meta)
    json.dump({"meta": meta, "fields": [dict(key=k, label=l, group=g, unit=u, dec=d) for k, l, g, u, d in fields],
               "catLabels": dict(catdefs), "cats": cats, "cols": cols},
              open(os.path.join(OUT, fname), "w"), ensure_ascii=False, separators=(",", ":"))


os.makedirs(OUT, exist_ok=True)
write("hospitals.json", records, FIELDS, CATS, {"unit": "病院", "colorLabel": "主機能", "x": "beds", "y": "fte_all"})

# ---------- 病棟単位 ----------
G_WSTAFF = "病棟の職員数(常勤+非常勤の常勤換算)"
WFIELDS = [
    ("beds", "許可病床数(一般+療養)", "病床", "床", 0),
    ("beds_general", "一般病床の許可病床数", "病床", "床", 0),
    ("beds_ryoyo", "療養病床の許可病床数", "病床", "床", 0),
    ("beds_maxuse", "最大使用病床数", "病床", "床", 0),
    ("fte_all", "病棟職員計(医師を除く11職種)", G_WSTAFF, "人", 1),
    ("fte_nurse", "看護職員計(看護師・准看護師・助産師)", G_WSTAFF, "人", 1),
] + [("fte_" + j, j, G_WSTAFF, "人", 1) for j in WARD_JOBS] + [
    ("adm", "新規入棟患者数(年間)", "診療実績", "人", 0),
    ("adm_er", "うち予定外の救急医療入院(年間)", "診療実績", "人", 0),
    ("days", "在棟患者延べ数(年間)", "診療実績", "人日", 0),
    ("dis", "退棟患者数(年間)", "診療実績", "人", 0),
    ("occ", "病床稼働率(在棟延べ数÷許可病床×366日)", "診療実績", "%", 1),
    ("alos", "平均在棟日数", "診療実績", "日", 1),
] + [(k, lab, G_F2, "件", 0) for k, _, lab in FORM2]
WCATS = [("pref", "都道府県"), ("area", "二次医療圏"), ("func", "機能区分"),
         ("fee", "入院基本料・特定入院料"), ("dept", "主とする診療科")]
for w in wards:
    try:
        w["pref"] = PREFS[int(w["pref"]) - 1]
    except (ValueError, IndexError):
        w["pref"] = "不明"
    w["area"] = w["area"] or "不明"
wards.sort(key=lambda r: r["id"])
write("wards.json", wards, WFIELDS, WCATS, {"unit": "病棟", "colorLabel": "病棟の機能区分", "x": "beds", "y": "fte_nurse"})
report["wards_out"] = len(wards)
report["ward_function_counts"] = Counter(w["func"] for w in wards)
report["ward_null_counts"] = {k: sum(1 for w in wards if w[k] is None) for k, *_ in WFIELDS}
report["ward_suspect"] = [(w["name"], w["fte_all"]) for w in wards if w["suspect"]]
report["ward_max_fte"] = sorted(((w["fte_all"] or 0, w["name"], w["beds"]) for w in wards), reverse=True)[:8]
report["form2_totals"] = {k: sum(r[k] or 0 for r in records) for k, _, _ in FORM2}
report["wards_matched_to_form2"] = sum(1 for w in wards if w["ope"] is not None or w["reh"] is not None)
report["hosp_sample_f2"] = {k: v for r in records if "市立函館病院" in r["name"] for k, v in r.items() if k in dict((a, 1) for a, _, _ in FORM2)}
report["ward_sample"] = [w for w in wards if "市立函館病院" in w["name"]][:2]
json.dump(report, open(os.path.join(OUT, "build_report.json"), "w"), ensure_ascii=False, indent=1)
print("hospitals:", len(records))
