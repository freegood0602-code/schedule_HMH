#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
健檢科部 工作分配排班系統 v5
- 人員定義表：圖示化 Excel風格
- 站別順序表：可調整順序
- 班表：優先標記、下拉排除、自動計算可覆蓋
"""

from http.server import HTTPServer, BaseHTTPRequestHandler
import json, urllib.parse, io
from datetime import date, timedelta
from copy import deepcopy

BASE_PERIOD = 3
BASE_DATE   = (2026, 7, 27)
PERIOD_DAYS = 28

N16 = ["淑娥","筱穎","欣潔","儀安","宜家","欣儒","于姍","蕙瑜","姿菁","雯茵","菀瑜","珮君","孟穎","嘉娟","曉娟","怡瑄"]
N11 = ["淑娥","筱穎","欣潔","儀安","宜家","欣儒","于姍","蕙瑜","姿菁","雯茵","菀瑜"]
N26 = ["淑娥","筱穎","欣潔","儀安","宜家","欣儒","于姍","蕙瑜","姿菁","雯茵","菀瑜","珮君","孟穎","嘉娟","曉娟","怡瑄","敏柔","伃倢","紫喬","悅蓁","曉雯","雅萍","雅虹","雨辰","思瑋","婷茹"]

PRIORITY_LABEL = {
    "胃腸(主)":"★1","胃腸(副)":"★2",
    "胃腸鏡清洗(前)":"★3","胃腸鏡清洗(後)":"★4",
    "家醫一(前)":"★5","家醫一(後)":"★6",
    "流動1":"●1","腹超一":"●2","家醫二":"●3","眼科散瞳+OCT":"●4",
    "抽血一(前)":"●5a","抽血二(後)":"●5b",
    "抽血一(後)":"●6a","抽血二(前)":"●6b",
    "流動2":"●7","協助抽血+EKG":"●8","協助眼科+量血壓":"●9",
}

# 站別群組順序（用於人員定義表欄位分組）
STATION_GROUPS = [
    ("★ 重要優先",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)"]),
    ("● 護理輪值",["流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","抽血二(前)","抽血二(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("批價/事務",["批價櫃台","協助櫃檯","腹超二","帶流程-放射科","發餐主/抹片","ABI","送檢體/抹片","流程主/寄信"]),
    ("放射師",["X光","協助X光骨密","心超","頸超","乳超","婦超","甲超","下午骨密","櫃台"]),
]
ALL_STATIONS_FLAT = [s for _, grp in STATION_GROUPS for s in grp]

STATIONS_DEFAULT = [
    {"name":"胃腸(主)",       "group":"A優先","seq":["筱穎","儀安","于姍","欣儒","蕙瑜","宜家","姿菁","雯茵","淑娥","欣潔","菀瑜"],"cur":"儀安","type":"single"},
    {"name":"胃腸(副)",       "group":"A優先","seq":[],"cur":"","type":"auto","auto_src":"胃腸(主)","auto_offset":-1},
    {"name":"胃腸鏡清洗(前)", "group":"A優先","seq":["姿菁","雯茵","欣儒","欣潔","筱穎","儀安","于姍","蕙瑜","宜家","淑娥","菀瑜"],"cur":"欣潔","type":"single"},
    {"name":"胃腸鏡清洗(後)", "group":"A優先","seq":["儀安","于姍","蕙瑜","宜家","淑娥","姿菁","雯茵","欣潔","筱穎","欣儒","菀瑜"],"cur":"宜家","type":"single"},
    {"name":"家醫一(前)",     "group":"A優先","seq":[],"cur":"","type":"auto","auto_src":"胃腸鏡清洗(後)","auto_offset":0},
    {"name":"家醫一(後)",     "group":"A優先","seq":[],"cur":"","type":"auto","auto_src":"胃腸鏡清洗(前)","auto_offset":0},
    {"name":"流動1",          "group":"護理","seq":N11,"cur":"淑娥","type":"single"},
    {"name":"腹超一",         "group":"護理","seq":N16,"cur":"雯茵","type":"single"},
    {"name":"家醫二",         "group":"護理","seq":N16,"cur":"蕙瑜","type":"single"},
    {"name":"眼科散瞳+OCT",   "group":"護理","seq":N16,"cur":"曉娟","type":"single"},
    {"name":"抽血一(前)",     "group":"護理","seq":N16,"cur":"姿菁","type":"single"},
    {"name":"抽血一(後)",     "group":"護理","seq":N16,"cur":"怡瑄","type":"single"},
    {"name":"抽血二(前)",     "group":"護理","seq":[],"cur":"","type":"auto","auto_src":"抽血一(後)","auto_offset":0},
    {"name":"抽血二(後)",     "group":"護理","seq":[],"cur":"","type":"auto","auto_src":"抽血一(前)","auto_offset":0},
    {"name":"流動2",          "group":"護理","seq":N16,"cur":"欣儒","type":"single"},
    {"name":"協助抽血+EKG",   "group":"護理","seq":N16,"cur":"孟穎","type":"single"},
    {"name":"協助眼科+量血壓","group":"護理","seq":N26,"cur":"于姍","type":"single"},
    {"name":"批價櫃台",       "group":"批價","seq":["敏柔","珮君","嘉娟","孟穎"],"cur":"珮君","type":"single"},
    {"name":"協助櫃檯",       "group":"批價","seq":["孟穎","敏柔","珮君","嘉娟"],"cur":"嘉娟","type":"single"},
    {"name":"櫃台",           "group":"放射","seq":["曉雯","筱穎","宜家","雅萍"],"cur":"筱穎","type":"single"},
    {"name":"腹超二",         "group":"事務","seq":["伃倢","紫喬","悅蓁","敏柔"],"cur":"紫喬","type":"single"},
    {"name":"帶流程-放射科",  "group":"事務","seq":["伃倢","紫喬","悅蓁","敏柔"],"cur":"悅蓁","type":"single"},
    {"name":"發餐主/抹片",    "group":"事務","seq":["悅蓁","伃倢","紫喬","敏柔"],"cur":"敏柔","type":"single"},
    {"name":"ABI",            "group":"事務","seq":["伃倢","紫喬","悅蓁","敏柔"],"cur":"伃倢","type":"single"},
    {"name":"送檢體/抹片",    "group":"服務","seq":["祐熙"],"cur":"祐熙","type":"single"},
    {"name":"流程主/寄信",    "group":"事務","seq":["悅蓁","伃倢","紫喬"],"cur":"悅蓁","type":"single"},
    {"name":"X光",            "group":"放射","seq":["雅萍","曉雯","婷茹","雅虹","雨辰","思瑋"],"cur":"思瑋","type":"single"},
    {"name":"協助X光骨密",    "group":"放射","seq":[],"cur":"","type":"pending"},
    {"name":"心超",           "group":"放射","seq":["思瑋","雅萍","雅虹","雨辰","婷茹"],"cur":"思瑋","type":"skip_xray"},
    {"name":"頸超",           "group":"放射","seq":["雨辰","雅虹","曉雯","曉娟"],"cur":"雅虹","type":"skip_xray"},
    {"name":"乳超",           "group":"放射","seq":["曉娟","雨辰","雅萍","曉雯"],"cur":"雨辰","type":"skip_xray"},
    {"name":"婦超",           "group":"待排","seq":["雨辰","雅虹","思瑋"],"cur":"","type":"pending"},
    {"name":"甲超",           "group":"待排","seq":["雨辰","婷茹","思瑋"],"cur":"","type":"pending"},
    {"name":"下午骨密",       "group":"放射","seq":["曉雯","雅萍","雅虹","思瑋","雨辰"],"cur":"雨辰","type":"single"},
    {"name":"肺功能",         "group":"待排","seq":[],"cur":"","type":"pending"},
    {"name":"下午櫃台",       "group":"待排","seq":[],"cur":"","type":"pending"},
    {"name":"學習櫃台",       "group":"待排","seq":[],"cur":"","type":"pending"},
]

STAFF_DEFAULT = [
    ("淑娥","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("筱穎","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓","批價櫃台","協助櫃檯","櫃台"]),
    ("欣潔","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("儀安","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("宜家","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓","批價櫃台","協助櫃檯","櫃台"]),
    ("欣儒","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("于姍","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("蕙瑜","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("姿菁","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("雯茵","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("菀瑜","護理師",["胃腸(主)","胃腸(副)","胃腸鏡清洗(前)","胃腸鏡清洗(後)","家醫一(前)","家醫一(後)","流動1","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","流動2","協助抽血+EKG","協助眼科+量血壓"]),
    ("曉娟","護理師",["流動2","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","協助抽血+EKG","協助眼科+量血壓","頸超","乳超"]),
    ("怡瑄","護理師",["流動2","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","協助抽血+EKG","協助眼科+量血壓"]),
    ("珮君","護理師兼批價",["批價櫃台","協助櫃檯","流動2","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","協助抽血+EKG","協助眼科+量血壓"]),
    ("孟穎","護理師兼批價",["批價櫃台","協助櫃檯","流動2","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","協助抽血+EKG","協助眼科+量血壓"]),
    ("嘉娟","護理師兼批價",["批價櫃台","協助櫃檯","流動2","腹超一","家醫二","眼科散瞳+OCT","抽血一(前)","抽血一(後)","協助抽血+EKG","協助眼科+量血壓"]),
    ("敏柔","事務員兼批價",["批價櫃台","協助櫃檯","腹超二","帶流程-放射科","發餐主/抹片","ABI","協助眼科+量血壓"]),
    ("伃倢","事務員",["腹超二","帶流程-放射科","發餐主/抹片","ABI","流程主/寄信"]),
    ("紫喬","事務員",["腹超二","帶流程-放射科","發餐主/抹片","ABI","流程主/寄信"]),
    ("悅蓁","事務員",["腹超二","帶流程-放射科","發餐主/抹片","ABI","流程主/寄信"]),
    ("祐熙","服務員",["送檢體/抹片"]),
    ("曉雯","放射師",["X光","協助X光骨密","心超","頸超","乳超","下午骨密","協助眼科+量血壓","批價櫃台","協助櫃檯","櫃台"]),
    ("雅萍","放射師",["X光","協助X光骨密","心超","乳超","下午骨密","協助眼科+量血壓","批價櫃台","協助櫃檯","櫃台"]),
    ("雅虹","放射師",["X光","協助X光骨密","心超","頸超","婦超","下午骨密","協助眼科+量血壓"]),
    ("雨辰","放射師",["X光","協助X光骨密","心超","頸超","乳超","婦超","甲超","下午骨密","協助眼科+量血壓"]),
    ("思瑋","放射師",["X光","協助X光骨密","心超","婦超","甲超","下午骨密","協助眼科+量血壓"]),
    ("婷茹","放射師",["X光","協助X光骨密","心超","甲超","下午骨密","協助眼科+量血壓"]),
]

ALLOWED_DUAL = set()
for _p in [
    ("家醫一(前)","胃腸鏡清洗(後)"),("家醫一(後)","胃腸鏡清洗(前)"),
    ("家醫一(前)","胃腸鏡清洗(前)"),("家醫一(後)","胃腸鏡清洗(後)"),
    ("抽血一(前)","抽血二(後)"),("抽血一(後)","抽血二(前)"),
    ("胃腸(主)","胃腸(副)"),
    ("下午骨密","X光"),("下午骨密","心超"),("下午骨密","頸超"),
    ("下午骨密","乳超"),("下午骨密","婦超"),("下午骨密","甲超"),("下午骨密","櫃台"),
    ("婦超","甲超"),("婦超","心超"),("甲超","心超"),
    ("協助X光骨密","乳超"),("協助X光骨密","甲超"),("協助X光骨密","婦超"),("協助X光骨密","心超"),
    ("批價櫃台","協助櫃檯"),("批價櫃台","ABI"),
    ("協助抽血+EKG","協助櫃檯"),("協助抽血+EKG","批價櫃台"),
    ("帶流程-放射科","流程主/寄信"),("帶流程-放射科","發餐主/抹片"),
    ("帶流程-放射科","ABI"),("帶流程-放射科","腹超二"),
    ("發餐主/抹片","流程主/寄信"),("發餐主/抹片","ABI"),("發餐主/抹片","腹超二"),
    ("ABI","流程主/寄信"),("ABI","腹超二"),("腹超二","流程主/寄信"),
    ("櫃台","下午骨密"),
]:
    ALLOWED_DUAL.add(tuple(sorted(_p)))

PRIORITY_ORDER = [
    "胃腸(主)","胃腸(副)",
    "胃腸鏡清洗(前)","胃腸鏡清洗(後)",
    "家醫一(前)","家醫一(後)",
    "流動1","腹超一","家醫二","眼科散瞳+OCT",
    "抽血一(前)","抽血二(後)","抽血一(後)","抽血二(前)",
    "流動2","協助抽血+EKG","協助眼科+量血壓",
]

app_state = {
    "period": BASE_PERIOD,
    "stations": deepcopy(STATIONS_DEFAULT),
    "staff": [list(s) for s in STAFF_DEFAULT],
    "manual": {},
}

# ══ 排班邏輯 ══
def get_period_date(n):
    base = date(*BASE_DATE)
    s = base + timedelta(days=(n-1)*PERIOD_DAYS)
    e = s + timedelta(days=PERIOD_DAYS-1)
    def roc(d): return f"{d.year-1911:03d}/{d.month:02d}/{d.day:02d}"
    return f"{roc(s)} ～ {roc(e)}"

def get_cur_idx(seq, cur):
    try: return seq.index(cur)
    except: return 0

def calc_single(s, offset, xray=""):
    if s["type"] in ("auto","pending") or not s.get("seq") or not s.get("cur"):
        return None
    seq, cur = s["seq"], s["cur"]
    n = len(seq)
    idx = (get_cur_idx(seq, cur) + offset) % n
    if s["type"] == "skip_xray" and xray:
        tries = 0
        while seq[idx] == xray and tries < n:
            idx = (idx+1) % n; tries += 1
    return seq[idx]

def calc_schedule(state, offset=0):
    stations = state["stations"]
    manual = state.get("manual", {}) if offset == 0 else {}
    xray_s = next((s for s in stations if s["name"]=="X光"), None)
    xray = calc_single(xray_s, offset) if xray_s else ""
    result = {}
    for s in stations:
        if s["type"] == "pending": result[s["name"]] = "（待排）"
        elif s["type"] != "auto": result[s["name"]] = calc_single(s, offset, xray)
    for s in stations:
        if s["type"] == "auto":
            src = next((x for x in stations if x["name"]==s["auto_src"]), None)
            off = s.get("auto_offset",0)
            result[s["name"]] = calc_single(src, offset+off, xray) if src else None
    if offset == 0:
        for k,v in manual.items():
            if v and v.strip(): result[k] = v.strip()
    return result

def check_conflicts(persons):
    ps = {}
    for sname, p in persons.items():
        if p and p != "（待排）": ps.setdefault(p,[]).append(sname)
    conflicts = {}
    for sname, p in persons.items():
        if p and p != "（待排）":
            bad = [o for o in ps.get(p,[]) if o!=sname and tuple(sorted([sname,o])) not in ALLOWED_DUAL]
            if bad: conflicts[sname] = "⚠ 也排了"+"、".join(bad)
    return conflicts

def get_members(state, stn):
    return [s[0] for s in state["staff"] if stn in s[2]]

def get_used_before(persons, stn):
    used = set()
    for s in PRIORITY_ORDER:
        if s == stn: break
        p = persons.get(s)
        if p and p != "（待排）": used.add(p)
    return used

# ══ CSS ══
CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Microsoft JhengHei",Arial,sans-serif;background:#f0f2f5;color:#333;font-size:13px}
.hdr{background:linear-gradient(135deg,#1f4e79,#2e75b6);color:white;padding:11px 20px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.3)}
.hdr h1{font-size:16px;font-weight:700}
.nav a{color:rgba(255,255,255,.85);text-decoration:none;padding:5px 13px;border-radius:20px;font-size:12px;border:1px solid rgba(255,255,255,.3);margin-left:6px}
.nav a.on,.nav a:hover{background:rgba(255,255,255,.2)}
.wrap{max-width:1400px;margin:14px auto;padding:0 12px}
.card{background:white;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.1);overflow:hidden;margin-bottom:12px}
.card-hd{padding:10px 16px;display:flex;align-items:center;justify-content:space-between;border-bottom:2px solid #2e75b6;background:#f8f9fa}
.card-hd h2{font-size:14px;font-weight:700;color:#1f4e79}
.pbar{display:flex;align-items:center;gap:10px;padding:11px 16px;flex-wrap:wrap}
.pnum{font-size:28px;font-weight:700;color:#1565c0}
.pdate{font-size:14px;color:#c00000;font-weight:600}
.btn{padding:6px 14px;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;transition:.15s;font-family:inherit;text-decoration:none;display:inline-block;white-space:nowrap}
.b-blue{background:#1f4e79;color:white}.b-blue:hover{background:#16375a}
.b-green{background:#2e7d32;color:white}.b-green:hover{background:#1b5e20}
.b-red{background:#c62828;color:white}.b-red:hover{background:#8b0000}
.b-org{background:#e65100;color:white}.b-org:hover{background:#bf360c}
.b-gray{background:#6c757d;color:white}.b-gray:hover{background:#495057}
.b-sm{padding:3px 9px;font-size:11px}
table{width:100%;border-collapse:collapse}
th{background:#2e75b6;color:white;padding:7px 8px;font-size:12px;text-align:center;white-space:nowrap;position:sticky;top:0;z-index:10}
td{padding:6px 8px;border-bottom:1px solid #eee;vertical-align:middle}
.sn{font-weight:600;font-size:12px}
.mb{font-size:10px;color:#aaa;margin-top:1px}
.ok{color:#2e7d32;font-weight:700}
.cf{color:#c00000;font-weight:700}
.auto{color:#7b1fa2;font-weight:600}
.cf-msg{font-size:11px;color:#c00000;font-weight:600}
.nxt{font-size:12px;color:#e65100;font-weight:600}
.sb{padding:8px 14px;border-radius:6px;margin-bottom:10px;font-weight:600;font-size:13px}
.sb-ok{background:#e8f5e9;color:#2e7d32;border:1px solid #c8e6c9}
.sb-cf{background:#ffebee;color:#c00000;border:1px solid #ffcdd2}
.split{display:flex;gap:3px;align-items:flex-start;flex-wrap:wrap}
.sp-lbl{font-size:10px;color:#999;margin-bottom:1px}
select{padding:3px 5px;border:1px solid #ddd;border-radius:4px;font-size:12px;font-family:inherit;background:white;max-width:100px}
select:focus{outline:none;border-color:#2e75b6}
input[type=text]{border:1px solid #ddd;border-radius:4px;padding:3px 5px;font-size:12px;font-family:inherit}
input[type=text]:focus{outline:none;border-color:#2e75b6}
.or{color:#ccc;font-size:10px;margin:0 2px;align-self:center}
.tag{display:inline-block;padding:1px 5px;border-radius:3px;font-size:10px;font-weight:700;margin-right:3px}
.ts{background:#e53935;color:white}
.td{background:#1565c0;color:white}
.ta{background:#7b1fa2;color:white;font-size:9px}
.cf-row{background:#fff8f8!important}
/* 人員定義表 */
.def-wrap{overflow-x:auto}
.def-table{border-collapse:collapse;font-size:11px;white-space:nowrap}
.def-table th{background:#2e75b6;color:white;padding:5px 6px;font-size:10px;position:sticky;top:0;z-index:5;text-align:center}
.def-table th.grp-hd{font-size:11px;font-weight:700;text-align:center}
.def-table .th-stn{writing-mode:vertical-lr;transform:rotate(180deg);min-height:70px;padding:6px 4px;font-size:10px;font-weight:600;cursor:default}
.def-table td{padding:4px 6px;border:1px solid #e0e0e0;text-align:center;vertical-align:middle}
.def-table td.name-cell{font-weight:700;text-align:left;background:#f8f9fa;position:sticky;left:0;z-index:3;min-width:60px;font-size:12px}
.def-table td.title-cell{font-size:10px;color:#666;background:#f8f9fa;position:sticky;left:60px;z-index:3;min-width:80px}
.def-table .chk-yes{background:#e8f5e9;cursor:pointer;font-size:14px}
.def-table .chk-no{background:#fafafa;cursor:pointer;color:#ddd;font-size:14px}
.def-table .chk-yes:hover{background:#c8e6c9}
.def-table .chk-no:hover{background:#f0f4c3}
.def-table .cnt-cell{background:#e3f2fd;font-weight:700;color:#1565c0;font-size:11px}
.def-table .grp-sep{border-left:3px solid #1f4e79!important}
/* 站別順序表 */
.seq-table{border-collapse:collapse;font-size:12px}
.seq-table th{background:#1f4e79;color:white;padding:7px 10px;font-size:12px}
.seq-table td{padding:6px 10px;border-bottom:1px solid #eee;vertical-align:top}
.seq-pill{display:inline-flex;align-items:center;background:#e3f2fd;border:1px solid #90caf9;border-radius:16px;padding:2px 8px 2px 10px;margin:2px;font-size:12px;font-weight:600;color:#1565c0}
.seq-pill.cur{background:#e8f5e9;border-color:#81c784;color:#2e7d32;font-weight:700}
.seq-pill.nxt{background:#fff9c4;border-color:#f9a825;color:#e65100}
.seq-pill .num{font-size:9px;color:#90a4ae;margin-right:4px}
.seq-pill .mv{cursor:pointer;color:#90a4ae;font-size:12px;margin-left:4px;padding:0 2px}
.seq-pill .mv:hover{color:#1565c0}
.legend{display:flex;gap:10px;flex-wrap:wrap;font-size:11px;color:#666;padding:8px 12px;justify-content:center}
.form-label{font-size:12px;color:#555;font-weight:600;margin-bottom:3px}
.form-row{display:flex;gap:10px;align-items:flex-start;flex-wrap:wrap;margin-bottom:10px}
"""

def head(t="排班系統"):
    return f'<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{t}</title><style>{CSS}</style></head><body>'

def nav(act="s"):
    tabs = [("s","/","📋 班表"),("d","/staff_def","👥 人員定義表"),("q","/seq","📋 站別順序表")]
    links = "".join(f'<a href="{url}" class="{"on" if act==k else ""}">{label}</a>' for k,url,label in tabs)
    return f'<div class="hdr"><h1>🏥 健檢科部 工作分配排班系統 v5</h1><div class="nav">{links}</div></div>'

def priority_tag(stn):
    lbl = PRIORITY_LABEL.get(stn,"")
    if lbl.startswith("★"): return f'<span class="tag ts">{lbl}</span>'
    if lbl.startswith("●"): return f'<span class="tag td">{lbl}</span>'
    return ""

def make_select(name, val, options):
    opts = '<option value="">-- 選 --</option>'
    for o in options:
        opts += f'<option value="{o}" {"selected" if o==val else ""}>{o}</option>'
    return f'<select name="{name}">{opts}</select>'

# ══ 班表頁 ══
OVERVIEW = [
    ("櫃台","single","櫃台"),("批價櫃台","single","批價櫃台"),("協助櫃檯","single","協助櫃檯"),
    ("流動1","single","流動1"),("流動2","single","流動2"),
    ("抽血一","split","抽血一(前)","抽血一(後)"),
    ("抽血二/POR","auto_edit","抽血二(前)","抽血二(後)"),
    ("腹超一","single","腹超一"),("腹超二","single","腹超二"),
    ("家醫一","auto_edit","家醫一(前)","家醫一(後)"),
    ("家醫二","single","家醫二"),("眼科散瞳+OCT","single","眼科散瞳+OCT"),
    ("ABI","single","ABI"),("協助眼科+量血壓","single","協助眼科+量血壓"),
    ("協助抽血+EKG","single","協助抽血+EKG"),
    ("胃腸","split_auto","胃腸(主)","胃腸(副)"),
    ("胃腸鏡清洗","split","胃腸鏡清洗(前)","胃腸鏡清洗(後)"),
    ("X光","single","X光"),("協助X光骨密","pending","協助X光骨密"),
    ("心超","single","心超"),("頸超","single","頸超"),("乳超","single","乳超"),
    ("婦超","pending","婦超"),("甲超","pending","甲超"),("下午骨密","single","下午骨密"),
    ("帶流程-放射科","single","帶流程-放射科"),("發餐主/抹片","single","發餐主/抹片"),
    ("送檢體/抹片","single","送檢體/抹片"),("流程主/寄信","single","流程主/寄信"),
    ("肺功能","pending","肺功能"),("下午櫃台","pending","下午櫃台"),("學習櫃台","pending","學習櫃台"),
]
GROUP_BG={"A優先":"#e8f5e9","護理":"#e3f2fd","批價":"#fffde7","事務":"#f3e5f5","服務":"#f5f5f5","放射":"#fce4ec","待排":"#e8eaf6"}

def render_schedule(state, msg=""):
    period=state["period"]
    p0=calc_schedule(state,0); p1=calc_schedule(state,1)
    cf=check_conflicts(p0)
    cf_n=len(set(cf.keys()))
    sb_cls="sb-cf" if cf_n>0 else "sb-ok"
    sb_msg=f"⚠ 本期有 {cf_n} 個衝突（紅字），可手動輸入覆蓋，輸入後仍會顯示衝突提醒" if cf_n>0 else "✅ 本期無衝突，排班正常"
    msg_html=f'<div class="sb" style="background:#e3f2fd;color:#1565c0;border:1px solid #bbdefb">{msg}</div>' if msg else ""

    rows=""
    for ov in OVERVIEW:
        mode=ov[1]; k1=ov[2]; k2=ov[3] if len(ov)>3 else None
        s1=next((s for s in state["stations"] if s["name"]==k1),None)
        bg=GROUP_BG.get(s1["group"] if s1 else "","#fff")
        t1=priority_tag(k1); t2=priority_tag(k2) if k2 else ""

        if mode=="pending":
            mv=state["manual"].get(k1,"")
            mb="、".join(get_members(state,k1)) or "（無設定）"
            sel=make_select(f"m_{k1}",mv,get_members(state,k1))
            rows+=f"""<tr style="background:{bg}">
              <td><div class="sn">{t1}{ov[0]} <span class="tag ta">待排</span></div>
                  <div class="mb">可排：{mb}</div></td>
              <td><div class="split">{sel} <span class="or">或</span>
                  <input type="text" name="t_{k1}" value="{mv}" style="width:80px" placeholder="手動填入"></div></td>
              <td><span style="color:#999;font-size:10px">待排</span></td>
              <td></td></tr>"""; continue

        if mode=="single":
            v=p0.get(k1,"") or ""; vn=p1.get(k1,"") or ""; c=cf.get(k1,"")
            mb="、".join(get_members(state,k1))
            used=get_used_before(p0,k1)
            av=[m for m in get_members(state,k1) if m not in used or m==v]
            is_auto=s1 and s1["type"]=="auto"
            sel=make_select(f"m_{k1}",v,get_members(state,k1) if is_auto else av)
            in_cls="cf" if c else ("auto" if is_auto else "ok")
            txt=f'<input type="text" name="t_{k1}" value="" style="width:75px" placeholder="{"覆蓋" if is_auto else "手動"}">'
            auto_tag='<span class="tag ta">自動</span>' if is_auto else ""
            rows+=f"""<tr style="background:{bg}" class="{"cf-row" if c else ""}">
              <td><div class="sn">{t1}{ov[0]}{auto_tag}</div><div class="mb">{mb[:40]}</div></td>
              <td><div class="split"><span class="{in_cls}" style="min-width:35px">{v}</span>{sel}<span class="or">或</span>{txt}</div></td>
              <td class="nxt">{vn}</td><td class="cf-msg">{c}</td></tr>"""; continue

        if mode=="split":
            v1=p0.get(k1,"") or ""; v2=p0.get(k2,"") or ""
            n1=p1.get(k1,"") or ""; n2=p1.get(k2,"") or ""
            c1=cf.get(k1,""); c2=cf.get(k2,""); call=" / ".join([x for x in [c1,c2] if x])
            mb="、".join(get_members(state,k1))
            used1=get_used_before(p0,k1); used2=get_used_before(p0,k2)
            av1=[m for m in get_members(state,k1) if m not in used1 or m==v1]
            av2=[m for m in get_members(state,k2) if m not in (used2|{v1}) or m==v2]
            s1e=make_select(f"m_{k1}",v1,av1); s2e=make_select(f"m_{k2}",v2,av2)
            t1i=f'<input type="text" name="t_{k1}" value="" style="width:70px" placeholder="手動">'
            t2i=f'<input type="text" name="t_{k2}" value="" style="width:70px" placeholder="手動">'
            rows+=f"""<tr style="background:{bg}" class="{"cf-row" if call else ""}">
              <td><div class="sn">{t1}{ov[0]}</div><div class="mb">{mb[:40]}</div></td>
              <td><div class="split" style="gap:4px">
                <div><div class="sp-lbl">{t1}前兩周</div><div class="split">{s1e}<span class="or">或</span>{t1i}</div></div>
                <span style="color:#ccc;align-self:center;margin:0 2px">/</span>
                <div><div class="sp-lbl">{t2 or t1}後兩周</div><div class="split">{s2e}<span class="or">或</span>{t2i}</div></div>
              </div></td>
              <td class="nxt">{n1}/{n2}</td><td class="cf-msg">{call}</td></tr>"""; continue

        if mode=="auto_edit":
            v1=p0.get(k1,"") or ""; v2=p0.get(k2,"") or ""
            n1=p1.get(k1,"") or ""; n2=p1.get(k2,"") or ""
            c1=cf.get(k1,""); c2=cf.get(k2,""); call=" / ".join([x for x in [c1,c2] if x])
            m1=get_members(state,k1); m2=get_members(state,k2)
            note="胃腸洗後↔前" if "家醫" in ov[0] else "抽血一後↔前"
            s1e=make_select(f"m_{k1}",v1,m1); s2e=make_select(f"m_{k2}",v2,m2)
            t1i=f'<input type="text" name="t_{k1}" value="" style="width:65px" placeholder="覆蓋">'
            t2i=f'<input type="text" name="t_{k2}" value="" style="width:65px" placeholder="覆蓋">'
            rows+=f"""<tr style="background:{bg}" class="{"cf-row" if call else ""}">
              <td><div class="sn">{t1}{ov[0]} <span class="tag ta">自動={note}</span></div>
                  <div class="mb">請假時可覆蓋</div></td>
              <td><div class="split" style="gap:4px">
                <div><div class="sp-lbl">前兩周（自動）</div>
                  <div class="split"><span class="auto" style="min-width:40px">{v1}</span>{s1e}<span class="or">或</span>{t1i}</div></div>
                <span style="color:#ccc;align-self:center">/</span>
                <div><div class="sp-lbl">後兩周（自動）</div>
                  <div class="split"><span class="auto" style="min-width:40px">{v2}</span>{s2e}<span class="or">或</span>{t2i}</div></div>
              </div></td>
              <td><span style="color:#7b1fa2;font-size:12px;font-weight:600">{n1}/{n2}</span></td>
              <td class="cf-msg">{call}</td></tr>"""; continue

        if mode=="split_auto":
            v1=p0.get(k1,"") or ""; v2=p0.get(k2,"") or ""
            n1=p1.get(k1,"") or ""; n2=p1.get(k2,"") or ""
            c1=cf.get(k1,""); c2=cf.get(k2,""); call=" / ".join([x for x in [c1,c2] if x])
            m1=get_members(state,k1)
            used1=get_used_before(p0,k1)
            av1=[m for m in m1 if m not in used1 or m==v1]
            s1e=make_select(f"m_{k1}",v1,av1); s2e=make_select(f"m_{k2}",v2,m1)
            t1i=f'<input type="text" name="t_{k1}" value="" style="width:70px" placeholder="手動">'
            t2i=f'<input type="text" name="t_{k2}" value="" style="width:70px" placeholder="覆蓋">'
            rows+=f"""<tr style="background:{bg}" class="{"cf-row" if call else ""}">
              <td><div class="sn">{t1}{ov[0]}</div><div class="mb">{"、".join(m1)[:40]}</div></td>
              <td><div class="split" style="gap:4px">
                <div><div class="sp-lbl">{t1}主</div><div class="split">{s1e}<span class="or">或</span>{t1i}</div></div>
                <span style="color:#ccc;align-self:center">/</span>
                <div><div class="sp-lbl">{t2 or ""}副<span class="tag ta" style="margin-left:2px">自動</span></div>
                  <div class="split"><span class="auto" style="min-width:40px">{v2}</span>{s2e}<span class="or">或</span>{t2i}</div></div>
              </div></td>
              <td class="nxt">{n1}/<span style="color:#7b1fa2">{n2}</span></td>
              <td class="cf-msg">{call}</td></tr>"""; continue

    html=head()+nav()
    html+=f"""<div class="wrap">
<form method="POST" action="/save">
<input type="hidden" name="period" value="{period}">
<div class="card">
  <div class="pbar">
    <div><div style="font-size:10px;color:#888">目前期別</div><div class="pnum">{period}</div></div>
    <div><div style="font-size:10px;color:#888">期別日期</div><div class="pdate">{get_period_date(period)}</div></div>
    <div style="flex:1"></div>
    <a href="/prev?period={period}" class="btn b-blue">◀ 上一期</a>
    <a href="/next?period={period}" class="btn b-blue">▶ 下一期</a>
    <button type="submit" class="btn b-org">💾 儲存</button>
    <a href="/export" class="btn b-green">📥 匯出 Excel</a>
  </div>
</div>
{msg_html}
<div class="sb {sb_cls}">{sb_msg}</div>
<div class="card"><table>
<thead><tr>
  <th style="width:22%;text-align:left">站別（可排人員）</th>
  <th style="width:38%">本期負責人</th>
  <th style="width:14%">下期預排</th>
  <th style="text-align:left">衝突提醒</th>
</tr></thead>
<tbody>{rows}</tbody></table></div>
<div class="legend">
  <span style="background:#e53935;color:white;padding:2px 8px;border-radius:3px;font-size:11px">★ 重要優先</span>
  <span style="background:#1565c0;color:white;padding:2px 8px;border-radius:3px;font-size:11px">● 後段輪值</span>
  <span style="background:#7b1fa2;color:white;padding:2px 8px;border-radius:3px;font-size:11px">自動計算</span>
  <span>🟢 綠=正常　🔴 紅=衝突　下拉已排除前面已用人員</span>
</div>
</form></div></body></html>"""
    return html

# ══ 人員定義表（Excel風格）══
def render_staff_def(state, msg=""):
    staff=state["staff"]
    # 計算每站人數
    stn_count={sn: sum(1 for s in staff if sn in s[2]) for _,g in STATION_GROUPS for sn in g}

    # 標題列：群組 + 站別（直排）
    # 先算各群組寬度
    grp_cols=[(gname,stns) for gname,stns in STATION_GROUPS]

    # 建立HTML
    # 第一行：姓名/職稱 + 群組標題合併
    grp_header="".join(
        f'<th colspan="{len(stns)}" class="grp-hd" style="background:{"#e53935" if "★" in gname else "#1565c0" if "●" in gname else "#37474f"};border-left:3px solid white">{gname}</th>'
        for gname,stns in grp_cols
    )
    # 第二行：各站別直排
    stn_header=""
    first_of_grp={stns[0] for _,stns in grp_cols}
    for gname,stns in grp_cols:
        for i,sn in enumerate(stns):
            border="border-left:3px solid white;" if i==0 else ""
            lbl=PRIORITY_LABEL.get(sn,"")
            tag=f'<span style="display:block;font-size:9px;margin-bottom:2px">{"🔴" if lbl.startswith("★") else "🔵" if lbl.startswith("●") else ""}{lbl}</span>' if lbl else ""
            stn_header+=f'<th class="th-stn" style="{border}">{tag}{sn}</th>'

    # 人數列
    cnt_row='<td class="name-cell" style="font-size:10px;color:#888">人數</td><td class="title-cell" style="font-size:10px;color:#888"></td>'
    for gname,stns in grp_cols:
        for i,sn in enumerate(stns):
            border="border-left:3px solid #1f4e79;" if i==0 else ""
            cnt_row+=f'<td class="cnt-cell" style="{border}">{stn_count.get(sn,0)}</td>'

    # 人員列
    person_rows=""
    for pi,(name,title,can_do) in enumerate(staff):
        cells=""
        for gname,stns in grp_cols:
            for i,sn in enumerate(stns):
                border="border-left:3px solid #90caf9;" if i==0 else ""
                has=sn in can_do
                cls="chk-yes" if has else "chk-no"
                icon="✓" if has else "·"
                cells+=f'<td class="{cls}" style="{border}" onclick="toggleStaff({pi},\'{sn}\')" title="{"點擊取消" if has else "點擊新增"}">{icon}</td>'
        person_rows+=f"""<tr>
          <td class="name-cell">{name}</td>
          <td class="title-cell">{title}</td>
          {cells}
        </tr>"""

    msg_html=f'<div class="sb sb-ok">{msg}</div>' if msg else ""
    html=head("人員定義表")+nav("d")
    html+=f"""<div class="wrap">
{msg_html}
<div class="card">
  <div class="card-hd">
    <h2>👥 人員定義表（點擊✓格切換）</h2>
    <div style="display:flex;gap:8px">
      <a href="/add_staff" class="btn b-green">➕ 新增人員</a>
      <button onclick="saveStaffDef()" class="btn b-org">💾 儲存變更</button>
    </div>
  </div>
  <div class="def-wrap" style="padding:12px">
  <form id="staffForm" method="POST" action="/save_staff_def">
  <table class="def-table" id="defTable">
    <thead>
      <tr>
        <th rowspan="2" style="position:sticky;left:0;z-index:20;background:#1f4e79;min-width:60px">姓名</th>
        <th rowspan="2" style="position:sticky;left:60px;z-index:20;background:#1f4e79;min-width:80px">職稱</th>
        {grp_header}
      </tr>
      <tr>{stn_header}</tr>
    </thead>
    <tbody>
      <tr style="background:#e3f2fd">{cnt_row}</tr>
      {person_rows}
    </tbody>
  </table>
  <input type="hidden" id="staffData" name="staff_data" value="">
  </form>
  </div>
</div>
<div style="font-size:11px;color:#888;text-align:center;padding:8px">
  🟢 ✓=可做這個站別　· =不可做　點擊格子切換　新增/刪除人員請點右上角按鈕
</div>
</div>

<script>
// 人員資料（從Python傳入）
var staffData = {json.dumps([[s[0],s[1],s[2]] for s in staff], ensure_ascii=False)};

function toggleStaff(personIdx, stnName) {{
  var can = staffData[personIdx][2];
  var idx = can.indexOf(stnName);
  if (idx >= 0) can.splice(idx, 1);
  else can.push(stnName);
  // 更新格子顯示
  var rows = document.querySelectorAll('#defTable tbody tr');
  var row = rows[personIdx + 1]; // +1 for count row
  var cells = row.querySelectorAll('td.chk-yes, td.chk-no');
  // 找到對應格子
  var allStns = {json.dumps(ALL_STATIONS_FLAT, ensure_ascii=False)};
  var stnIdx = allStns.indexOf(stnName);
  if (stnIdx >= 0 && cells[stnIdx]) {{
    var cell = cells[stnIdx];
    var has = can.indexOf(stnName) >= 0;
    cell.className = has ? 'chk-yes' : 'chk-no';
    cell.innerHTML = has ? '✓' : '·';
    cell.title = has ? '點擊取消' : '點擊新增';
  }}
  updateCount();
}}

function updateCount() {{
  var allStns = {json.dumps(ALL_STATIONS_FLAT, ensure_ascii=False)};
  var countRow = document.querySelectorAll('#defTable tbody tr')[0];
  var cells = countRow.querySelectorAll('.cnt-cell');
  allStns.forEach(function(sn, idx) {{
    var cnt = staffData.filter(function(s) {{ return s[2].indexOf(sn) >= 0; }}).length;
    if (cells[idx]) cells[idx].textContent = cnt;
  }});
}}

function saveStaffDef() {{
  document.getElementById('staffData').value = JSON.stringify(staffData);
  document.getElementById('staffForm').submit();
}}
</script>
</body></html>"""
    return html

# ══ 站別順序表 ══
def render_seq(state, msg=""):
    stations=state["stations"]
    persons=calc_schedule(state,0)
    persons_nxt=calc_schedule(state,1)

    rows=""
    # 照優先順序顯示，然後是其他站別
    ordered_names=PRIORITY_ORDER + [s["name"] for s in stations if s["name"] not in PRIORITY_ORDER]
    ordered_stns=[next((s for s in stations if s["name"]==n),None) for n in ordered_names]
    ordered_stns=[s for s in ordered_stns if s]

    for s in ordered_stns:
        if s["type"]=="auto":
            note=f'自動 = {s.get("auto_src","")} {"上一期" if s.get("auto_offset",0)==-1 else "(對調)"}'
            cur_val=persons.get(s["name"],"") or "？"
            nxt_val=persons_nxt.get(s["name"],"") or "？"
            lbl=PRIORITY_LABEL.get(s["name"],"")
            tag=priority_tag(s["name"])
            rows+=f"""<tr>
              <td><div class="sn">{tag}{s["name"]}</div></td>
              <td><span style="color:#7b1fa2;font-size:11px">自動計算</span></td>
              <td><span style="color:#7b1fa2;font-weight:700">{cur_val}</span></td>
              <td><span style="color:#e65100">{nxt_val}</span></td>
              <td style="font-size:11px;color:#888">{note}</td>
              <td></td></tr>"""
            continue

        if s["type"]=="pending" and not s.get("seq"):
            lbl=PRIORITY_LABEL.get(s["name"],"")
            rows+=f"""<tr>
              <td><div class="sn">{priority_tag(s["name"])}{s["name"]}</div></td>
              <td><span style="background:#e8eaf6;color:#5c6bc0;padding:2px 8px;border-radius:10px;font-size:11px">待排（無固定順序）</span></td>
              <td colspan="4" style="color:#888;font-size:11px">有需要才排，手動填入</td></tr>"""
            continue

        seq=s.get("seq",[])
        cur=s.get("cur","")
        cur_val=persons.get(s["name"],"") or "？"
        nxt_val=persons_nxt.get(s["name"],"") or "？"

        # 產生順序藥丸
        cur_idx=get_cur_idx(seq,cur)
        pills=""
        for i,person in enumerate(seq):
            is_cur=(person==cur)
            is_nxt=(i==(cur_idx+1)%len(seq))
            pcls="cur" if is_cur else ("nxt" if is_nxt else "")
            label="▶本期" if is_cur else ("▷下期" if is_nxt else "")
            badge=f'<span style="font-size:9px;color:{"#2e7d32" if is_cur else "#e65100" if is_nxt else "#90a4ae"};margin-left:2px">{label}</span>'
            # 移動按鈕
            up=f'<a href="/seq_move?stn={urllib.parse.quote(s["name"])}&idx={i}&dir=up" class="mv" title="往前">↑</a>' if i>0 else ''
            dn=f'<a href="/seq_move?stn={urllib.parse.quote(s["name"])}&idx={i}&dir=dn" class="mv" title="往後">↓</a>' if i<len(seq)-1 else ''
            rm=f'<a href="/seq_move?stn={urllib.parse.quote(s["name"])}&idx={i}&dir=rm" class="mv" style="color:#e53935" title="移出順序">×</a>'
            pills+=f'<span class="seq-pill {pcls}"><span class="num">{i+1}</span>{person}{badge}{up}{dn}{rm}</span>'

        tag=priority_tag(s["name"])
        rows+=f"""<tr>
          <td style="min-width:130px"><div class="sn">{tag}{s["name"]}</div>
              <div class="mb">{len(seq)}人</div></td>
          <td style="padding:6px">{pills}</td>
          <td><span style="color:#2e7d32;font-weight:700">{cur_val}</span></td>
          <td><span style="color:#e65100;font-weight:600">{nxt_val}</span></td>
          <td style="font-size:11px;color:#888">{s.get("type","")}</td>
          <td>
            <form style="display:inline" method="POST" action="/seq_setcur">
              <input type="hidden" name="stn" value="{s['name']}">
              <select name="new_cur" style="font-size:11px;padding:2px 4px">
                {"".join(f"<option {'selected' if p==cur else ''} value='{p}'>{p}</option>" for p in seq)}
              </select>
              <button type="submit" class="btn b-blue b-sm" style="margin-left:4px">設為本期✓</button>
            </form>
          </td></tr>"""

    msg_html=f'<div class="sb sb-ok">{msg}</div>' if msg else ""
    html=head("站別順序表")+nav("q")
    html+=f"""<div class="wrap">
{msg_html}
<div class="card">
  <div class="card-hd">
    <h2>📋 站別順序表（可調整順序、設定本期✓）</h2>
    <div style="font-size:11px;color:#888">↑↓ 調整順序　× 移出　▶ 本期（綠色）　▷ 下期（橙色）</div>
  </div>
  <div style="overflow-x:auto">
  <table class="seq-table">
    <thead><tr>
      <th style="width:130px;text-align:left">站別</th>
      <th style="text-align:left">輪值順序（可調整）</th>
      <th style="width:80px">本期</th>
      <th style="width:80px">下期</th>
      <th style="width:80px">類型</th>
      <th style="width:200px">設定本期✓</th>
    </tr></thead>
    <tbody>{rows}</tbody>
  </table>
  </div>
</div>
</div></body></html>"""
    return html

# ══ 匯出 Excel ══
def export_excel(state):
    try:
        import openpyxl
        from openpyxl.styles import PatternFill,Font,Alignment,Border,Side
    except: return None,"請先安裝 openpyxl"
    period=state["period"]
    persons=calc_schedule(state,0); persons_nxt=calc_schedule(state,1)
    cf=check_conflicts(persons)
    def fill(h): return PatternFill("solid",fgColor=h)
    def fnt(bold=False,size=11,color="000000"): return Font(name="新細明體",bold=bold,size=size,color=color)
    def algn(h="left"): return Alignment(horizontal=h,vertical="center",wrap_text=True)
    thin=Side(style='thin',color='000000')
    def bd(): return Border(left=thin,right=thin,top=thin,bottom=thin)
    def gv(key):
        if key is None: return ""
        if isinstance(key,tuple):
            v1=persons.get(key[0],"") or ""; v2=persons.get(key[1],"") or ""
            return f"{'' if v1=='（待排）' else v1}/{'' if v2=='（待排）' else v2}"
        p=persons.get(key,""); return "" if p=="（待排）" else (p or "")
    def hcf(key):
        if key is None: return False
        if isinstance(key,tuple): return bool(cf.get(key[0],"") or cf.get(key[1],""))
        return bool(cf.get(key,""))
    wb=openpyxl.Workbook()
    BLUE="00B0F0";YELLOW="FFFF00";GREEN="00B050";GRAY="D9D9D9";PINK="FF66FF";WHITE="FFFFFF"
    ws1=wb.active; ws1.title="格式"
    ws1.column_dimensions["A"].width=24.25; ws1.column_dimensions["B"].width=18
    ws1.row_dimensions[1].height=18
    ws1["A1"].value="日期"; ws1["A1"].font=fnt(bold=True); ws1["A1"].alignment=algn()
    ws1["B1"].value=get_period_date(period); ws1["B1"].font=fnt(bold=True); ws1["B1"].alignment=algn("center")
    p1_rows=[
        ("櫃台",BLUE,"櫃台"),("批櫃",YELLOW,"批價櫃台"),("流動 1",YELLOW,"流動1"),("流動 2",YELLOW,"流動2"),
        ("抽血一",YELLOW,("抽血一(前)","抽血一(後)")),("抽血二/POR",YELLOW,("抽血二(前)","抽血二(後)")),
        ("腹超一",WHITE,"腹超一"),("腹超二",YELLOW,"腹超二"),
        ("家醫一",YELLOW,("家醫一(前)","家醫一(後)")),("家醫二",YELLOW,"家醫二"),
        ("眼科",WHITE,"眼科散瞳+OCT"),("ABI",YELLOW,"ABI"),
        ("協助眼科+量血壓",YELLOW,"協助眼科+量血壓"),("協助抽血",YELLOW,None),
        ("協助抽血+EKG",YELLOW,"協助抽血+EKG"),
        ("胃腸",BLUE,("胃腸(主)","胃腸(副)")),("胃腸洗",BLUE,("胃腸鏡清洗(前)","胃腸鏡清洗(後)")),
        ("X光",GREEN,"X光"),("協助X光、骨密",GREEN,"協助X光骨密"),
        ("心超",YELLOW,"心超"),("頸超",YELLOW,"頸超"),("乳超",YELLOW,"乳超"),
        ("婦超",YELLOW,"婦超"),("肺功能",YELLOW,"肺功能"),("甲超",YELLOW,"甲超"),
        ("協助櫃檯",YELLOW,"協助櫃檯"),("下午櫃檯",BLUE,"下午櫃台"),
        ("下午骨密",GRAY,"下午骨密"),("學習櫃台",YELLOW,"學習櫃台"),
        ("流程主/寄信",PINK,"流程主/寄信"),("發餐主/抹片",PINK,"發餐主/抹片"),
        ("送檢體/抹片",PINK,"送檢體/抹片"),
    ]
    for i,(lbl,color,key) in enumerate(p1_rows,2):
        ws1.row_dimensions[i].height=18
        c=ws1.cell(i,1,lbl); c.font=fnt(bold=True); c.fill=fill(color); c.alignment=algn(); c.border=bd()
        val=gv(key); hc=hcf(key)
        c=ws1.cell(i,2,val); c.font=fnt(bold=True,color="C00000" if hc else "000000")
        c.fill=fill("FFE0E0" if hc else ("FFF2CC" if val else "FFFFFF")); c.alignment=algn("center"); c.border=bd()
    ws2=wb.create_sheet("工作內容")
    ws2.column_dimensions["A"].width=21; ws2.column_dimensions["B"].width=18.5; ws2.column_dimensions["C"].width=18.5
    ws2.row_dimensions[1].height=18
    for col,h in enumerate(["項  目/診  室","診室準備","負責人"],1):
        c=ws2.cell(1,col,h); c.font=fnt(bold=True,color="000000"); c.fill=fill("D7E4BD"); c.alignment=algn("center"); c.border=bd()
    p2_rows=[
        ("批價櫃台","","批價櫃台"),("協助櫃檯","","協助櫃檯"),("櫃台","","櫃台"),
        ("帶流程-放射科檢查","吧檯","帶流程-放射科"),("送檢體","","送檢體/抹片"),
        ("發餐 / 10:00 抹片跟診","抹片室","發餐主/抹片"),
        ("抽血一         (前兩周/後兩周)","",("抽血一(前)","抽血一(後)")),
        ("抽血二/POR (前兩周/後兩周)","",("抽血二(前)","抽血二(後)")),
        ("腹超一","","腹超一"),("腹超二","","腹超二"),
        ("家醫一         (前兩周/後兩周)","EKG1檢查室",("家醫一(前)","家醫一(後)")),
        ("家醫二","EKG2檢查室","家醫二"),("眼科散瞳+OCT檢查","","眼科散瞳+OCT"),
        ("流動/協助眼科+量血壓","","協助眼科+量血壓"),
        ("流動/協助抽血+EKG","VIP室","協助抽血+EKG"),
        ("ABI","","ABI"),("放射科會診/ 流動","","流動2"),("流動","","流動1"),
        ("胃腸(主)/胃腸(副)","",("胃腸(主)","胃腸(副)")),
        ("胃腸鏡清洗 (前兩周/後兩周)","",("胃腸鏡清洗(前)","胃腸鏡清洗(後)")),
        ("X光/骨密","身高體重室","X光"),("刪除EKG","",None),
    ]
    for i,(lbl,room,key) in enumerate(p2_rows,2):
        ws2.row_dimensions[i].height=18
        c=ws2.cell(i,1,lbl); c.font=fnt(bold=True); c.fill=fill("FFFF00"); c.alignment=algn(); c.border=bd()
        c=ws2.cell(i,2,room); c.font=fnt(bold=True); c.fill=fill("FFFF00"); c.alignment=algn(); c.border=bd()
        val=gv(key); hc=hcf(key)
        c=ws2.cell(i,3,val); c.font=fnt(bold=True,color="C00000" if hc else "000000")
        c.fill=fill("FFE0E0" if hc else ("FFF2CC" if val else "FFFFFF")); c.alignment=algn("center"); c.border=bd()
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read(),None

# ══ HTTP Server ══
class Handler(BaseHTTPRequestHandler):
    def log_message(self,f,*a): pass
    def send_html(self,html):
        self.send_response(200); self.send_header("Content-Type","text/html; charset=utf-8"); self.end_headers()
        self.wfile.write(html.encode("utf-8"))
    def redirect(self,url):
        self.send_response(302); self.send_header("Location",url); self.end_headers()

    def do_GET(self):
        path=self.path.split("?")[0]
        params=dict(urllib.parse.parse_qsl(self.path.split("?")[1])) if "?" in self.path else {}
        if path=="/":
            self.send_html(render_schedule(app_state,"✅ 儲存成功！" if params.get("msg")=="saved" else ""))
        elif path=="/next":
            app_state["period"]=int(params.get("period",app_state["period"]))+1
            app_state["manual"]={}; self.redirect("/")
        elif path=="/prev":
            app_state["period"]=max(1,int(params.get("period",app_state["period"]))-1)
            app_state["manual"]={}; self.redirect("/")
        elif path=="/export":
            data,err=export_excel(app_state)
            if err: self.send_html(f"<h3>{err}</h3>")
            else:
                fname=f"健檢科部_排班_第{app_state['period']}期.xlsx"
                self.send_response(200)
                self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition",f"attachment; filename*=UTF-8''{urllib.parse.quote(fname)}")
                self.end_headers(); self.wfile.write(data)
        elif path=="/staff_def":
            self.send_html(render_staff_def(app_state,"✅ 儲存成功！" if params.get("msg")=="saved" else ""))
        elif path=="/seq":
            self.send_html(render_seq(app_state,"✅ 已更新！" if params.get("msg")=="saved" else ""))
        elif path=="/seq_move":
            stn=urllib.parse.unquote(params.get("stn",""))
            idx=int(params.get("idx",0)); direction=params.get("dir","up")
            s=next((x for x in app_state["stations"] if x["name"]==stn),None)
            if s and s.get("seq"):
                seq=s["seq"]
                if direction=="up" and idx>0:
                    seq[idx],seq[idx-1]=seq[idx-1],seq[idx]
                elif direction=="dn" and idx<len(seq)-1:
                    seq[idx],seq[idx+1]=seq[idx+1],seq[idx]
                elif direction=="rm":
                    person=seq.pop(idx)
                    if s.get("cur")==person: s["cur"]=seq[0] if seq else ""
            self.redirect("/seq?msg=saved")
        elif path=="/add_staff":
            self.send_html(render_edit_staff(app_state))
        elif path=="/edit_staff":
            self.send_html(render_edit_staff(app_state,int(params.get("idx",0))))
        elif path=="/del_staff":
            idx=int(params.get("idx",0))
            if 0<=idx<len(app_state["staff"]):
                name=app_state["staff"][idx][0]
                app_state["staff"].pop(idx)
                for s in app_state["stations"]:
                    if name in s.get("seq",[]):
                        s["seq"].remove(name)
                        if s.get("cur")==name: s["cur"]=s["seq"][0] if s["seq"] else ""
            self.redirect("/staff_def?msg=saved")
        else:
            self.send_response(404); self.end_headers()

    def do_POST(self):
        length=int(self.headers.get("Content-Length",0))
        body=self.rfile.read(length).decode("utf-8")
        params=urllib.parse.parse_qs(body,keep_blank_values=True)
        def get(k,d=""): return params.get(k,[d])[0]
        path=self.path.split("?")[0]
        qp=dict(urllib.parse.parse_qsl(self.path.split("?")[1])) if "?" in self.path else {}

        if path=="/save":
            app_state["period"]=int(get("period",str(app_state["period"])))
            new_manual={}
            for sn in [s["name"] for s in app_state["stations"]]:
                tv=get(f"t_{sn}","").strip(); sv=get(f"m_{sn}","").strip()
                v=tv if tv else sv
                if v and v!="-- 選 --": new_manual[sn]=v
            app_state["manual"]=new_manual; self.redirect("/?msg=saved")

        elif path=="/save_staff_def":
            raw=get("staff_data","")
            if raw:
                try:
                    new_staff=json.loads(raw)
                    # 更新staff
                    old_names={s[0] for s in app_state["staff"]}
                    new_names={s[0] for s in new_staff}
                    app_state["staff"]=[list(s) for s in new_staff]
                    # 同步站別順序表
                    removed=old_names-new_names
                    added=new_names-old_names
                    for s in app_state["stations"]:
                        for name in removed:
                            if name in s.get("seq",[]):
                                s["seq"].remove(name)
                                if s.get("cur")==name: s["cur"]=s["seq"][0] if s["seq"] else ""
                        # 新增人員加到可做的站別末尾
                        for ns in new_staff:
                            if ns[0] in added and s["name"] in ns[2] and ns[0] not in s.get("seq",[]):
                                s.setdefault("seq",[]).append(ns[0])
                                if not s.get("cur"): s["cur"]=ns[0]
                except: pass
            self.redirect("/staff_def?msg=saved")

        elif path=="/seq_setcur":
            stn=get("stn"); new_cur=get("new_cur")
            s=next((x for x in app_state["stations"] if x["name"]==stn),None)
            if s and new_cur in s.get("seq",[]): s["cur"]=new_cur
            self.redirect("/seq?msg=saved")

        elif path=="/save_staff":
            idx_str=qp.get("idx","")
            name=get("name").strip(); title=get("title").strip()
            can_do=params.get("can_do",[])
            if name and title:
                if idx_str:
                    idx=int(idx_str); old=app_state["staff"][idx][0]
                    if old!=name:
                        for s in app_state["stations"]:
                            if old in s.get("seq",[]):
                                i=s["seq"].index(old); s["seq"][i]=name
                                if s.get("cur")==old: s["cur"]=name
                    app_state["staff"][idx]=[name,title,can_do]
                else:
                    app_state["staff"].append([name,title,can_do])
                    for s in app_state["stations"]:
                        if s["name"] in can_do and name not in s.get("seq",[]):
                            s.setdefault("seq",[]).append(name)
                            if not s.get("cur"): s["cur"]=name
            self.redirect("/staff_def?msg=saved")
        else:
            self.redirect("/")

def render_edit_staff(state,idx=None):
    all_stn=[s["name"] for s in state["stations"]]
    if idx is not None and idx<len(state["staff"]):
        name,title,can_do=state["staff"][idx]; action=f"/save_staff?idx={idx}"; heading=f"編輯：{name}"
    else:
        name,title,can_do="","",""; can_do=[]; action="/save_staff"; heading="新增人員"
    checks=""
    for sn in all_stn:
        lbl=PRIORITY_LABEL.get(sn,"")
        tag=f'<span class="tag {"ts" if lbl.startswith("★") else "td" if lbl.startswith("●") else "ta"}">{lbl}</span>' if lbl else ""
        checks+=f'<label style="display:inline-flex;align-items:center;gap:3px;margin:3px 10px 3px 0;font-size:12px;cursor:pointer"><input type="checkbox" name="can_do" value="{sn}" {"checked" if sn in can_do else ""}> {tag}{sn}</label>'
    html=head(heading)+nav("d")
    html+=f"""<div class="wrap"><div class="card">
  <div class="card-hd"><h2>{"✏️" if idx is not None else "➕"} {heading}</h2></div>
  <div style="padding:16px"><form method="POST" action="{action}">
  <div class="form-row">
    <div><div class="form-label">姓名</div><input type="text" name="name" value="{name}" required style="width:140px"></div>
    <div><div class="form-label">職稱</div><input type="text" name="title" value="{title}" required style="width:160px"></div>
  </div>
  <div style="margin-top:12px"><div class="form-label">可做站別</div>
    <div style="margin-top:6px;padding:10px;background:#f8f9fa;border-radius:8px;line-height:2.2">{checks}</div>
  </div>
  <div style="margin-top:12px;display:flex;gap:8px">
    <button type="submit" class="btn b-green">💾 儲存</button>
    <a href="/staff_def" class="btn b-gray">取消</a>
  </div>
  </form></div>
</div></div></body></html>"""
    return html

if __name__=="__main__":
    port=8080
    server=HTTPServer(("0.0.0.0",port),Handler)
    print(f"✅ 排班系統 v5 啟動！")
    print(f"👉 瀏覽器開啟：http://localhost:{port}")
    print(f"   Ctrl+C 結束")
    try: server.serve_forever()
    except KeyboardInterrupt: print("\n🛑 已關閉")
