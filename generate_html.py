"""
クリニック数ダッシュボード HTMLレポート生成スクリプト
実行すると index.html を生成します
"""
import pandas as pd
import plotly.express as px
import calendar
from datetime import date
from pathlib import Path

# ── ファイルパス ──────────────────────────────────────
BOX_PATH  = Path(r"C:\Users\宮城杏奈\Box\総合企画部_特殊案件\その他\院情報一覧カウント\院情報一覧_カウント自動化.xlsx")
LOCAL_PATH = Path.home() / "Documents" / "クリニックDB" / "院情報一覧_カウント自動化.xlsx"
OUTPUT_HTML = Path(__file__).parent / "index.html"

ORANGE_TWIST_COUNT = 24
C_HEADER = "#2C3E50"
C_TOTAL  = "#FFF2CC"
C_ORANGE = "#E67E22"
C_BLUE   = "#2980B9"
C_GREEN  = "#D9EAD3"

TARGET_BRANDS = [
    "湘南美容クリニック","湘南美容クリニック（ﾍﾞﾄﾅﾑ）","湘南歯科クリニック","湘南AGAクリニック","湘南美容皮フ科",
    "湘南皮膚科クリニック","SBC NEO Skin Clinic","肌の青空クリニック","湘南内科皮フ科クリニック",
    "湘南美容皮フ科内科クリニック","イテウォン","湘南メディカル記念病院","新宿近視クリニック",
    "西新宿整形外科","SBC横浜駅前整形外科","六本木レディース","神奈川レディース","リッツ美容外科",
    "リゼクリニック","ゴリラクリニック","JUNCLINIC","SBC東京医療大学付属クリニック",
    "SBC東京接骨院","SBC BODY ARCHI","SkinGo!","Wen & Weng Family Clinic",
    "The Chelsea Clinic","Chelsea Aesthetics","Chelsea Asthetics","Gangnam Laser Clinic","Rochor Centre Clinic"
]
EXCLUDE_PR = ["SBC東京医療大学付属クリニック","SBC東京接骨院","SBC BODY ARCHI"]
OVERSEAS_KW = ["DS","DSS","RCC","WWFC","WWMG","SkinGo","Chelsea","ﾍﾞﾄﾅﾑ","Vietnam","Shoubikai"]
HOUJIN_ORDER = [
    "医療法人 湘美会","医療法人社団 孝和会","医療法人社団 樹慶会","医療法人社団 愛恵会",
    "医療法人社団 風林会","医療法人社団 菜寿会","医療法人社団 孝仁会",
    "一般社団法人美央斗会","一般社団法人MASA","健美会",
    "医療法人社団 リッツ美容外科","株式会社 湘南美容クリニック",
    "DS","DSS","DSS(FC)","RCC","WWFC","WWMG",
    "学校法人 SBC東京医療大学","医療法人 きびたき会","SBC東京接骨院","株式会社 MG"
]
EXCLUDE_HOUJIN = ["学校法人 SBC東京医療大学","㈻SBC東京医療大学附属","医療法人 きびたき会","医療法人社団 百花会","SBC東京接骨院","株式会社 MG"]

JIKEI_COLS = [
    ("湘南美容クリニック","美容外科・皮膚科"),("湘南美容クリニック","オンライン"),("湘南美容クリニック","スキンLab"),
    ("湘南歯科クリニック",None),("湘南AGAクリニック",None),("湘南美容皮フ科",None),("湘南皮膚科クリニック",None),
    ("SBC NEO Skin Clinic",None),("肌の青空クリニック",None),("湘南内科皮フ科クリニック",None),
    ("湘南美容皮フ科内科クリニック",None),("イテウォン",None),("湘南メディカル記念病院",None),
    ("新宿近視クリニック",None),("西新宿整形外科",None),("SBC横浜駅前整形外科",None),
    ("六本木レディース",None),("神奈川レディース",None),("リッツ美容外科",None),
    ("リゼクリニック",None),("ゴリラクリニック",None),("JUNCLINIC",None),
    ("SBC東京医療大学付属クリニック",None),("SBC東京接骨院",None),("SBC BODY ARCHI",None),
    ("湘南美容クリニック（ﾍﾞﾄﾅﾑ）",None),("The Chelsea Clinic",None),("Chelsea Aesthetics",None),
    ("Gangnam Laser Clinic",None),("SkinGo!",None),("Wen & Weng Family Clinic",None),("Rochor Centre Clinic",None),
]

# 法人グループ定義
HOUJIN_GROUPS = [
    ("①6医療法人合計", ["医療法人 湘美会","医療法人社団 孝和会","医療法人社団 菜寿会",
                       "医療法人社団 愛恵会","医療法人社団 樹慶会","医療法人社団 リッツ美容外科",
                       "一般社団法人MASA","健美会","法人無し（個人開設）","個人/その他"]),
    ("②3医療法人合計", ["医療法人社団 風林会","医療法人 きびたき会","医療法人社団 百花会"]),
    ("③医療法人社団十二会", ["医療法人社団十二会"]),
    ("④2医療法人合計", ["医療法人社団美咲会","一般社団法人美央斗会"]),
    ("⑤㈻SBC東京医療大学附属", ["㈻SBC東京医療大学附属","学校法人 SBC東京医療大学"]),
    ("⑥株式会社SBC湘南接骨院", ["株式会社SBC湘南接骨院"]),
    ("⑦SBCメディカルグループ株式会社", ["SBCメディカルグループ株式会社","株式会社 MG"]),
    ("⑧Shoubikai Medical Vietnam Co., Ltd.", ["Shoubikai Medical Vietnam Co., Ltd."]),
    ("⑨WWMG", ["WWMG"]),
    ("⑩DS", ["DS"]),
    ("⑪DSS", ["DSS"]),
    ("⑫DSS(FC)", ["DSS(FC)"]),
    ("⑬WWFC", ["WWFC"]),
    ("⑭RCC", ["RCC"]),
]

# ブランド列定義（時系列表用・指定順）
JIKEI_BRAND_COLS = [
    ("湘南美容クリニック","美容外科・皮膚科"),
    ("湘南美容クリニック","オンライン"),
    ("湘南美容クリニック","スキンLab"),
    ("湘南歯科クリニック",None),
    ("湘南AGAクリニック",None),
    ("湘南美容皮フ科",None),
    ("湘南皮膚科クリニック",None),
    ("SBC NEO Skin Clinic",None),
    ("肌の青空クリニック",None),
    ("湘南内科皮フ科クリニック",None),
    ("湘南美容皮フ科内科クリニック",None),
    ("イテウォン",None),
    ("湘南メディカル記念病院",None),
    ("新宿近視クリニック",None),
    ("西新宿整形外科",None),
    ("SBC横浜駅前整形外科",None),
    ("六本木レディース",None),
    ("神奈川レディース",None),
    ("リッツ美容外科",None),
    ("リゼクリニック",None),
    ("ゴリラクリニック",None),
    ("JUNCLINIC",None),
    ("SBC東京医療大学付属クリニック",None),
    ("SBC東京接骨院",None),
    ("SBC BODY ARCHI",None),
    ("湘南美容クリニック（ﾍﾞﾄﾅﾑ）",None),
    ("The Chelsea Clinic",None),
    ("Chelsea Aesthetics",None),
    ("Gangnam Laser Clinic",None),
    ("SkinGo!",None),
    ("Wen & Weng Family Clinic",None),
    ("Rochor Centre Clinic",None),
]
EXISTING_GROUP_END = 19  # リッツ美容外科まで（既存グループ合計）


def get_path():
    if BOX_PATH.exists(): return BOX_PATH
    if LOCAL_PATH.exists(): return LOCAL_PATH
    raise FileNotFoundError("Excelファイルが見つかりません")

def to_ts(v):
    if isinstance(v, pd.Timestamp): return v
    if pd.isna(v): return None
    try: return pd.Timestamp(v)
    except: return None

def load_data():
    path = get_path()
    df = pd.read_excel(path, sheet_name="クリニック一覧")
    for c in ["開院日","MA日","移転拡張日","業態転換日","閉院日"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    def parse_limit(v):
        if isinstance(v, pd.Timestamp) and not pd.isna(v): return v.strftime("%Y/%m")
        if pd.isna(v): return ""
        return str(v).strip()
    excl = "猶予期限（業態転換後の院が開院するまでの期間）"
    if excl in df.columns:
        df[excl] = df[excl].apply(parse_limit)
    return df

def load_brand_settings():
    try:
        path = get_path()
        raw = pd.read_excel(path, sheet_name="ブランド設定", header=None)
        header_row = None
        for i, row in raw.iterrows():
            if any("ブランド名" in str(v) for v in row.values):
                header_row = i; break
        if header_row is None: return [], [], []
        df = pd.read_excel(path, sheet_name="ブランド設定", header=header_row)
        df = df.dropna(subset=["ブランド名"])
        if "表示順" in df.columns: df = df.sort_values("表示順")
        cols, flags, excl_pr = [], [], []
        for _, row in df.iterrows():
            b = str(row.get("ブランド名","") or "").strip()
            g_val = row.get("業態","")
            g = str(g_val).strip() if pd.notna(g_val) and str(g_val).strip() not in ("","nan") else None
            ex = str(row.get("既存グループ合計","") or "").strip()
            ep = str(row.get("広報IR除外","") or "").strip()
            if b:
                cols.append((b,g)); flags.append(ex=="○")
                if ep=="○": excl_pr.append(b)
        return cols, flags, excl_pr
    except: return [], [], []

def get_brand(row):
    b = str(row.get("ブランド","") or "").strip()
    return "Chelsea Aesthetics" if b=="Chelsea Asthetics" else b

def get_region(row):
    rv = row.get("海外／国内","")
    r = "" if pd.isna(rv) else str(rv).strip()
    if not r:
        b = str(row.get("ブランド","") or ""); h = str(row.get("法人名","") or "")
        return "海外" if any(kw in b or kw in h for kw in OVERSEAS_KW) else "国内"
    return r

def check_active(row, target):
    ma = to_ts(row.get("MA日")); op = to_ts(row.get("開院日"))
    base = ma if ma is not None else op
    if base is None: return False
    d_conv = to_ts(row.get("業態転換日")); d_close = to_ts(row.get("閉院日"))
    final_end = d_conv if d_conv is not None else d_close
    base_only = pd.Timestamp(base.year, base.month, base.day)
    return base_only <= target and (final_end is None or final_end > target)

def month_end_ts(y, m): return pd.Timestamp(y, m, calendar.monthrange(y,m)[1], 23,59,59)
def month_start_ts(y, m): return pd.Timestamp(y, m, 1, 0, 0, 0)

def aggregate(df, me, ms, target_brands, exclude_pr):
    r1, r2, r3 = {}, {}, {}
    for _, row in df.iterrows():
        brand = get_brand(row); gyoutai = str(row.get("業態","") or "").strip()
        houjin = str(row.get("法人名","") or "個人/その他").strip()
        region = get_region(row)
        if not brand or not gyoutai or brand not in target_brands: continue
        is_pr = brand not in exclude_pr
        if check_active(row, me):
            k1 = brand+"|"+gyoutai
            r1.setdefault(k1,{"all":0,"pr":0}); r1[k1]["all"]+=1
            if is_pr: r1[k1]["pr"]+=1
            k2 = region+"|"+houjin
            r2.setdefault(k2,{"all":0,"pr":0}); r2[k2]["all"]+=1
            if is_pr: r2[k2]["pr"]+=1
        if check_active(row, ms):
            if houjin in HOUJIN_ORDER:
                r3.setdefault(houjin,{"all":0})
                if houjin not in EXCLUDE_HOUJIN: r3[houjin]["all"]+=1
    return r1, r2, r3

def get_base_date(row):
    ma = to_ts(row.get("MA日")); op = to_ts(row.get("開院日"))
    return ma if ma is not None else op

def build_history(df, target_brands, exclude_pr):
    """年別・月別の開院・閉院・業態転換データを構築"""
    today = date.today()
    all_dates = pd.concat([df["開院日"].dropna(), df["MA日"].dropna(),
                           df["閉院日"].dropna(), df["業態転換日"].dropna()])
    data_max_year = int(all_dates.dt.year.max()) if not all_dates.empty else today.year
    hist_years = list(range(2025, max(data_max_year, today.year) + 1))
    result = {}
    for year in hist_years:
        monthly = {}
        for month in range(1, 13):
            ms = pd.Timestamp(year, month, 1)
            me = pd.Timestamp(year, month, calendar.monthrange(year, month)[1])
            opened, closed, convert = [], [], []
            for _, row in df.iterrows():
                brand = get_brand(row)
                if not brand or brand not in target_brands: continue
                base = get_base_date(row)
                d_conv = to_ts(row.get("業態転換日")); d_close = to_ts(row.get("閉院日"))
                # 開院
                tenkan_mae = str(row.get("転換前業態","") or "").strip()
                is_converted = tenkan_mae not in ("","nan")
                if base is not None:
                    base_only = pd.Timestamp(base.year, base.month, base.day)
                    if ms <= base_only <= me:
                        opened.append({
                            "種別": "🔵 業態転換" if is_converted else "🟢 新規開院",
                            "開院日": base_only.strftime("%Y/%m/%d"),
                            "院名": row.get("正式名称",""), "ブランド": brand,
                            "業態": str(row.get("業態","") or ""),
                            "法人名": str(row.get("法人名","") or ""),
                            "国内／海外": get_region(row),
                        })
                # 業態転換
                if d_conv is not None:
                    conv_only = pd.Timestamp(d_conv.year, d_conv.month, d_conv.day)
                    if ms <= conv_only <= me:
                        before = str(row.get("転換前業態","") or "").strip()
                        if not before or before=="nan": before = str(row.get("業態","") or "").strip()
                        after_n = str(row.get("転換後院名","") or "").strip()
                        after_g = str(row.get("転換後業態","") or "").strip()
                        convert.append({
                            "転換前院名": row.get("正式名称",""), "転換前ブランド": brand,
                            "転換前業態": before if before else "―", "　": "→",
                            "転換後院名": after_n if after_n else "―",
                            "転換後業態": after_g if after_g else "―",
                            "法人名": str(row.get("法人名","") or ""),
                            "国内／海外": get_region(row),
                            "業態転換日": conv_only.strftime("%Y/%m/%d"),
                        })
                # 閉院
                close_ref = d_close if d_close is not None else d_conv
                if close_ref is not None:
                    close_only = pd.Timestamp(close_ref.year, close_ref.month, close_ref.day)
                    if ms <= close_only <= me:
                        closed.append({
                            "種別": "🔵 業態転換" if d_conv is not None else "🔴 閉院",
                            "閉院日": close_only.strftime("%Y/%m/%d"),
                            "院名": row.get("正式名称",""), "ブランド": brand,
                            "業態": str(row.get("業態","") or ""),
                            "法人名": str(row.get("法人名","") or ""),
                            "国内／海外": get_region(row),
                        })
            monthly[month] = {"opened": opened, "closed": closed, "convert": convert}
        result[year] = monthly
    return result, hist_years

def build_history_html(history, hist_years, mode="openclose", prefix=""):
    """開院・閉院または業態転換のアコーディオンHTMLを生成"""
    html = ""
    uid = 0
    for year in hist_years:
        monthly = history[year]
        if mode == "openclose":
            ytotal_o = sum(len(v["opened"]) for v in monthly.values())
            ytotal_c = sum(len(v["closed"]) for v in monthly.values())
            if ytotal_o == 0 and ytotal_c == 0: continue
            year_label = f"📅 {year}年　｜　🟢 開院 {ytotal_o} 院　　🔴 閉院 {ytotal_c} 院"
            color = C_HEADER
        else:
            ytotal_k = sum(len(v["convert"]) for v in monthly.values())
            if ytotal_k == 0: continue
            year_label = f"📅 {year}年　｜　🔵 業態転換 {ytotal_k} 院"
            color = C_BLUE

        html += f'<div style="background:{color};color:white;padding:10px 16px;font-weight:bold;font-size:16px;border-radius:6px;margin-top:20px">{year_label}</div>'

        for month in range(1, 13):
            opened  = monthly[month]["opened"]
            closed  = monthly[month]["closed"]
            convert = monthly[month]["convert"]
            uid += 1

            if mode == "openclose":
                n_o = len(opened); n_c = len(closed)
                if n_o == 0 and n_c == 0: continue
                label = f"{year}年{month}月　🟢 開院 {n_o}院　　🔴 閉院 {n_c}院"
                inner = ""
                if n_o > 0:
                    inner += f'<div style="border-left:4px solid #27AE60;background:#D5F5E3;padding:5px 10px;font-weight:bold;margin-bottom:6px">🟢 開院　{n_o}院</div>'
                    inner += df_to_html_table(pd.DataFrame(opened), highlight_last=False)
                    inner += "<br>"
                if n_c > 0:
                    inner += f'<div style="border-left:4px solid #E74C3C;background:#FADBD8;padding:5px 10px;font-weight:bold;margin-bottom:6px">🔴 閉院　{n_c}院</div>'
                    inner += df_to_html_table(pd.DataFrame(closed), highlight_last=False)
            else:
                n_k = len(convert)
                if n_k == 0: continue
                label = f"{year}年{month}月　🔵 業態転換 {n_k}院"
                inner = f'<div style="border-left:4px solid {C_BLUE};background:#D6EAF8;padding:5px 10px;font-weight:bold;margin-bottom:6px">🔵 業態転換　{n_k}院</div>'
                inner += df_to_html_table(pd.DataFrame(convert), highlight_last=False)

            html += f"""
<div style="margin-top:8px;border:1px solid #ddd;border-radius:6px;overflow:hidden">
  <div onclick="toggleAcc('{prefix}acc{uid}','{prefix}arr{uid}')" style="padding:10px 16px;cursor:pointer;background:#f8f9fa;display:flex;align-items:center;gap:8px">
    <span id="{prefix}arr{uid}" style="font-size:12px">▶</span> {label}
  </div>
  <div id="{prefix}acc{uid}" style="display:none;padding:16px">
    {inner}
  </div>
</div>"""
    return html

def df_to_html_table(df, highlight_last=True, orange_last=False):
    rows_html = ""
    for i, (_, row) in enumerate(df.iterrows()):
        is_last = i == len(df) - 1
        is_second_last = i == len(df) - 2
        is_third_last = i == len(df) - 3
        if orange_last and is_last:
            style = f'background:{C_ORANGE};color:white;font-weight:bold'
        elif orange_last and is_second_last:
            style = f'background:{C_ORANGE};color:white'
        elif highlight_last and is_last:
            style = f'background:{C_TOTAL};font-weight:bold'
        elif highlight_last and (is_second_last or is_third_last) and orange_last:
            style = f'background:{C_TOTAL};font-weight:bold'
        else:
            style = 'background:white'
        cells = "".join(f'<td style="padding:6px 10px;border:1px solid #ddd">{v}</td>' for v in row)
        rows_html += f'<tr style="{style}">{cells}</tr>'
    headers = "".join(f'<th style="padding:8px 10px;background:{C_HEADER};color:white;border:1px solid #555">{c}</th>' for c in df.columns)
    return f'<table style="border-collapse:collapse;width:100%;font-size:14px"><thead><tr>{headers}</tr></thead><tbody>{rows_html}</tbody></table>'


def build_jikei_table_with_houjin(df, target_brands, exclude_pr):
    """2021/01から今月まで、月別×ブランド・法人グループの集計レコードリストを返す"""
    today = date.today()
    sy, sm = 2021, 1
    records = []

    while (sy, sm) <= (today.year, today.month):
        tme = month_end_ts(sy, sm)
        rec = {"年月": f"{sy}/{sm:02d}"}

        # ブランド×業態ごとのカウント
        brand_counts = {}
        for _, row in df.iterrows():
            brand = get_brand(row)
            gyoutai = str(row.get("業態", "") or "").strip()
            houjin = str(row.get("法人名", "") or "個人/その他").strip()
            if not brand or brand not in target_brands:
                continue
            if check_active(row, tme):
                key = (brand, gyoutai)
                brand_counts[key] = brand_counts.get(key, 0) + 1

        # JIKEI_BRAND_COLS の各列カウント
        col_counts = []
        for (brand, gyoutai) in JIKEI_BRAND_COLS:
            if gyoutai is not None:
                cnt = brand_counts.get((brand, gyoutai), 0)
            else:
                # 業態指定なし → そのブランド全業態合計
                cnt = sum(v for (b, g), v in brand_counts.items() if b == brand)
            col_counts.append(cnt)
            label = f"{brand}({gyoutai})" if gyoutai else brand
            rec[label] = cnt

        # 既存グループ合計（最初の EXISTING_GROUP_END ブランド）
        existing_sum = sum(col_counts[:EXISTING_GROUP_END])
        rec["既存グループ合計"] = existing_sum

        # 全拠点合計
        total_all = sum(col_counts)
        rec["全拠点合計"] = total_all

        # OrangeTwist
        rec["OrangeTwist"] = ORANGE_TWIST_COUNT

        # IR・広報用（exclude_pr 以外のブランド合計 + OrangeTwist）
        pr_sum = 0
        for _, row in df.iterrows():
            brand = get_brand(row)
            if not brand or brand not in target_brands:
                continue
            if brand in exclude_pr:
                continue
            if check_active(row, tme):
                pr_sum += 1
        rec["IR・広報用"] = pr_sum + ORANGE_TWIST_COUNT

        # 法人グループ別カウント
        houjin_group_counts = {}
        for group_name, members in HOUJIN_GROUPS:
            cnt = 0
            for _, row in df.iterrows():
                brand = get_brand(row)
                houjin = str(row.get("法人名", "") or "個人/その他").strip()
                if not brand or brand not in target_brands:
                    continue
                if houjin not in members:
                    continue
                if check_active(row, tme):
                    cnt += 1
            houjin_group_counts[group_name] = cnt
            rec[group_name] = cnt

        records.append(rec)

        sm += 1
        if sm > 12:
            sm = 1
            sy += 1

    return records


def jikei_to_html_table(records):
    """時系列法人・ブランド別集計テーブルのHTMLを生成"""
    if not records:
        return "<p>データなし</p>"

    # ヘッダー列構成
    brand_labels = []
    for (brand, gyoutai) in JIKEI_BRAND_COLS:
        label = f"{brand}({gyoutai})" if gyoutai else brand
        brand_labels.append(label)

    houjin_labels = [g for g, _ in HOUJIN_GROUPS]

    # 2行ヘッダー構成
    # 第1行: グループ見出し (colspan)
    # 第2行: 個別列名

    # ブランド列数
    n_brand = len(brand_labels)

    # 色設定
    th_base = f'style="padding:6px 8px;background:{C_HEADER};color:white;border:1px solid #444;white-space:nowrap;font-size:11px;position:sticky;top:0;z-index:2"'
    th_orange = f'style="padding:6px 8px;background:{C_ORANGE};color:white;border:1px solid #444;white-space:nowrap;font-size:11px;position:sticky;top:0;z-index:2"'
    th_blue = f'style="padding:6px 8px;background:{C_BLUE};color:white;border:1px solid #444;white-space:nowrap;font-size:11px;position:sticky;top:0;z-index:2"'
    th_green = f'style="padding:6px 8px;background:#1E8449;color:white;border:1px solid #444;white-space:nowrap;font-size:11px;position:sticky;top:0;z-index:2"'
    th_sticky = f'style="padding:6px 8px;background:{C_HEADER};color:white;border:1px solid #444;white-space:nowrap;font-size:11px;position:sticky;top:0;left:0;z-index:3"'

    # 第1行ヘッダー
    row1 = f'<th rowspan="2" {th_sticky}>年月</th>'
    row1 += f'<th colspan="{n_brand}" {th_base}>ブランド</th>'
    row1 += f'<th rowspan="2" {th_green}>既存グループ合計</th>'
    row1 += f'<th rowspan="2" {th_base}>全拠点合計</th>'
    row1 += f'<th rowspan="2" {th_orange}>OrangeTwist</th>'
    row1 += f'<th rowspan="2" {th_orange}>IR・広報用</th>'
    row1 += f'<th colspan="{len(houjin_labels)}" {th_blue}>法人別</th>'

    # 第2行ヘッダー
    row2 = ""
    for label in brand_labels:
        row2 += f'<th {th_base}>{label}</th>'
    for label in houjin_labels:
        row2 += f'<th {th_blue}>{label}</th>'

    thead = f"<thead><tr>{row1}</tr><tr>{row2}</tr></thead>"

    # tbody
    tbody_rows = ""
    for rec in records:
        ym = rec["年月"]
        # 12月は薄青背景
        is_dec = ym.endswith("/12")
        row_bg = "#EBF5FB" if is_dec else "white"

        td_base = f'style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-size:12px;background:{row_bg}"'
        td_green = f'style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-size:12px;background:#D5F5E3;font-weight:bold"'
        td_orange = f'style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-size:12px;background:#FAD7A0;font-weight:bold"'
        td_sticky = f'style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-size:12px;background:{row_bg};position:sticky;left:0;z-index:1;font-weight:bold;white-space:nowrap"'

        cells = f'<td {td_sticky}>{ym}</td>'

        for label in brand_labels:
            cells += f'<td {td_base}>{rec.get(label, 0)}</td>'

        cells += f'<td {td_green}>{rec.get("既存グループ合計", 0)}</td>'
        cells += f'<td {td_base}>{rec.get("全拠点合計", 0)}</td>'
        cells += f'<td {td_orange}>{rec.get("OrangeTwist", ORANGE_TWIST_COUNT)}</td>'
        cells += f'<td {td_orange}>{rec.get("IR・広報用", 0)}</td>'

        for label in houjin_labels:
            cells += f'<td {td_base}>{rec.get(label, 0)}</td>'

        tbody_rows += f'<tr>{cells}</tr>'

    tbody = f"<tbody>{tbody_rows}</tbody>"

    table = f'<table style="border-collapse:collapse;font-size:12px;min-width:100%">{thead}{tbody}</table>'
    return f'<div style="overflow-x:auto;max-height:70vh;overflow-y:auto;border:1px solid #ddd;border-radius:4px">{table}</div>'


def generate():
    print("データを読み込み中...")
    df = load_data()
    brand_cols, existing_flags, exclude_pr = load_brand_settings()
    if not brand_cols:
        brand_cols = [(b,None) for b in TARGET_BRANDS]
        existing_flags = [True]*19 + [False]*(len(brand_cols)-19)
        exclude_pr = EXCLUDE_PR

    today = date.today()
    y, m = today.year, today.month
    me = month_end_ts(y, m); ms = month_start_ts(y, m)
    r1, r2, r3 = aggregate(df, me, ms, [b for b,_ in brand_cols], exclude_pr)

    sum_pr = sum(v["pr"] for v in r1.values())
    sum_all = sum(v["all"] for v in r1.values())

    # ── ブランド×業態表 ──
    brand_rows = []
    for brand, gyoutai in brand_cols:
        if gyoutai: keys=[k for k in r1 if k==f"{brand}|{gyoutai}"]
        else: keys=[k for k in r1 if k.startswith(f"{brand}|")]
        cnt_all=sum(r1[k]["all"] for k in keys); cnt_pr=sum(r1[k]["pr"] for k in keys)
        label = f"{brand}({gyoutai})" if gyoutai else brand
        brand_rows.append({"ブランド":label,"広報・IR用":cnt_pr,"全拠点(ALL)":cnt_all})
    brand_df = pd.DataFrame(brand_rows)
    total_row = pd.DataFrame([
        {"ブランド":"集計内訳合計","広報・IR用":sum_pr,"全拠点(ALL)":sum_all},
        {"ブランド":"海外(OrangeTwist)加算","広報・IR用":ORANGE_TWIST_COUNT,"全拠点(ALL)":"－"},
        {"ブランド":"最終報告数値","広報・IR用":sum_pr+ORANGE_TWIST_COUNT,"全拠点(ALL)":sum_all},
    ])
    brand_df = pd.concat([brand_df, total_row], ignore_index=True)

    # ── 国内/海外×法人表 ──
    dom, ovs = [], []
    for region in ["国内","海外"]:
        for k in sorted([k for k in r2 if k.startswith(region+"|")]):
            p=k.split("|",1); e={"法人名":p[1],"広報・IR用":r2[k]["pr"],"全拠点(ALL)":r2[k]["all"]}
            (dom if region=="国内" else ovs).append(e)
    dom_df=pd.DataFrame(dom); ovs_df=pd.DataFrame(ovs)

    # ── 法人別表 ──
    houjin_rows=[]; total_h=0
    for h in HOUJIN_ORDER:
        cnt=r3.get(h,{}).get("all",0); houjin_rows.append({"法人名":h,"全拠点(ALL)":cnt}); total_h+=cnt
    houjin_rows.append({"法人名":"合計","全拠点(ALL)":total_h})
    houjin_df=pd.DataFrame(houjin_rows)

    # ── 時系列推移（法人・ブランド別）テーブル ──
    print("時系列推移（法人・ブランド別）を集計中...")
    jikei_records = build_jikei_table_with_houjin(df, [b for b,_ in brand_cols], exclude_pr)
    jikei_table_html = jikei_to_html_table(jikei_records)

    # ── ブランド別棒グラフ ──
    plot_df = brand_df[brand_df["ブランド"].isin([f"{b}({g})" if g else b for b,g in brand_cols])].copy()
    fig2=px.bar(plot_df,x="全拠点(ALL)",y="ブランド",orientation="h",height=550,color_discrete_sequence=[C_BLUE])
    fig2.update_layout(yaxis=dict(autorange="reversed"),margin=dict(t=30))
    chart2_html = fig2.to_html(full_html=False, include_plotlyjs="cdn")

    # ── 開院・閉院・業態転換履歴 ──
    print("開院・閉院・業態転換履歴を集計中...")
    history, hist_years = build_history(df, [b for b,_ in brand_cols], exclude_pr)
    openclose_html = build_history_html(history, hist_years, mode="openclose", prefix="oc")
    convert_html   = build_history_html(history, hist_years, mode="convert",   prefix="cv")

    # ── HTML生成 ──
    report_date = f"{y}年{m}月末"
    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>クリニック数ダッシュボード {report_date}</title>
<style>
  body{{font-family:'Meiryo',sans-serif;margin:0;background:#f5f5f5;color:#333}}
  .header{{background:{C_HEADER};color:white;padding:20px 30px}}
  .header h1{{margin:0;font-size:22px}}
  .header p{{margin:5px 0 0;opacity:.7;font-size:13px}}
  .kpi{{display:flex;gap:16px;padding:20px 30px;background:white;border-bottom:1px solid #ddd}}
  .kpi-card{{flex:1;background:#f8f9fa;border-radius:8px;padding:16px;text-align:center;border-left:4px solid {C_BLUE}}}
  .kpi-card.orange{{border-left-color:{C_ORANGE}}}
  .kpi-label{{font-size:12px;color:#666;margin-bottom:4px}}
  .kpi-value{{font-size:28px;font-weight:bold;color:{C_HEADER}}}
  .tabs{{display:flex;background:white;border-bottom:2px solid #ddd;padding:0 30px}}
  .tab{{padding:12px 20px;cursor:pointer;border-bottom:3px solid transparent;font-size:14px;color:#666}}
  .tab.active{{border-bottom-color:{C_BLUE};color:{C_BLUE};font-weight:bold}}
  .content{{display:none;padding:24px 30px}}
  .content.active{{display:block}}
  .section-title{{background:{C_GREEN};padding:8px 14px;font-weight:bold;border-radius:4px;margin-bottom:12px}}
  .grid-2{{display:grid;grid-template-columns:1fr 1fr;gap:20px}}
  .grid-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px}}
  .box{{background:white;border-radius:8px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
  .updated{{text-align:right;font-size:12px;color:#999;padding:8px 30px}}
</style>
</head>
<body>
<div class="header">
  <h1>クリニック数ダッシュボード</h1>
  <p>{report_date} 時点</p>
</div>

<div class="kpi">
  <div class="kpi-card orange">
    <div class="kpi-label">IR・広報用（OrangeTwist含む）</div>
    <div class="kpi-value">{sum_pr+ORANGE_TWIST_COUNT:,} 院</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">全拠点(ALL)</div>
    <div class="kpi-value">{sum_all:,} 院</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">OrangeTwist（別管理）</div>
    <div class="kpi-value">{ORANGE_TWIST_COUNT} 院</div>
  </div>
</div>

<div class="tabs">
  <div class="tab active" onclick="showTab('brand')">📋 ブランド×業態</div>
  <div class="tab" onclick="showTab('region')">🌏 国内／海外×法人</div>
  <div class="tab" onclick="showTab('houjin')">🏢 法人別</div>
  <div class="tab" onclick="showTab('trend')">📈 時系列推移</div>
  <div class="tab" onclick="showTab('history')">🏥 開院・閉院履歴</div>
  <div class="tab" onclick="showTab('convert')">🔵 業態転換履歴</div>
</div>

<div id="brand" class="content active">
  <div class="box">
    <div class="section-title">{report_date} ブランド×業態</div>
    {df_to_html_table(brand_df, highlight_last=False, orange_last=True)}
  </div>
  <div class="box" style="margin-top:20px">{chart2_html}</div>
</div>

<div id="region" class="content">
  <div class="grid-2">
    <div class="box">
      <div class="section-title">国内</div>
      {df_to_html_table(dom_df) if not dom_df.empty else '<p>データなし</p>'}
    </div>
    <div class="box">
      <div class="section-title">海外</div>
      {df_to_html_table(ovs_df) if not ovs_df.empty else '<p>データなし</p>'}
    </div>
  </div>
</div>

<div id="houjin" class="content">
  <div class="box">
    <div class="section-title">{y}年{m}月初 法人別内訳</div>
    {df_to_html_table(houjin_df)}
  </div>
</div>

<div id="trend" class="content">
  <div class="box">
    <div class="section-title">時系列推移（法人・ブランド別）2021/01〜{y}/{m:02d}</div>
    {jikei_table_html}
  </div>
</div>

<div id="history" class="content">
  <h3 style="margin-top:0">開院・閉院 年別・月別履歴</h3>
  {openclose_html}
</div>

<div id="convert" class="content">
  <h3 style="margin-top:0">業態転換 年別・月別履歴</h3>
  {convert_html}
</div>

<div class="updated">最終更新: {today.strftime('%Y/%m/%d')}</div>

<script>
function showTab(id){{
  document.querySelectorAll('.content').forEach(el=>el.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(el=>el.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  event.target.classList.add('active');
}}
function toggleAcc(accId, arrId){{
  var el=document.getElementById(accId);
  var arr=document.getElementById(arrId);
  if(el.style.display==='none'){{el.style.display='block';arr.textContent='▼';}}
  else{{el.style.display='none';arr.textContent='▶';}}
}}
</script>
</body>
</html>"""

    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"HTMLレポートを生成しました: {OUTPUT_HTML}")
    return True

if __name__ == "__main__":
    generate()
    print("完了！")
