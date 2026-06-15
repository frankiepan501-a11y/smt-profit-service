# -*- coding: utf-8 -*-
"""速卖通毛利报表 半自动云服务 (Zeabur)。
领星 multiplatform 对接过期 → 改后台导出+映射。运营把 2 店资金明细(俄/非俄)+物流账单
传到飞书「速卖通月度数据上传台」→ cron 每月16号 → POST /recompute → 下载附件解析 →
映射领星 cg_price → 算毛利 → 灌全渠道总表(速卖通行) → 飞书汇报 Frankie。
口径: 销售=资金明细成交金额(CNY,半托管已折供货价); 退款=售中+售后; 平台费=佣金+交易服务费+
联盟佣金+金币营销+cashback+营销技术/智投/新商孵化−各退回; 采购=领星cg_price×数量(仅净成交>0行);
物流=物流账单按订单内数量分摊(仅自发货单)。毛利=成交−退款−平台费−采购−物流。
回款=已放款金额(放款金额列,真实收口径A,随放款进度增长)。
待结算行平台费暂0(放款滞后)→ 按已结算行费率估算(任意日期算都准,不必等16号放款);
月底幂等重跑会用真实放款值收口(误差归零)。可配合运营提前导数→提前cron。
env: FEISHU_APP_ID/FEISHU_APP_SECRET(聪哥1号) / LX_APP_ID/LX_SECRET(领星) / AUTH_TOKEN
"""
import os, io, json, time, hashlib, base64, datetime, tempfile, warnings
from collections import defaultdict
import requests, openpyxl
from Crypto.Cipher import AES
from fastapi import FastAPI, Request, HTTPException
warnings.filterwarnings("ignore")

FEISHU_APP_ID = os.environ["FEISHU_APP_ID"]
FEISHU_APP_SECRET = os.environ["FEISHU_APP_SECRET"]
LX_APP_ID = os.environ.get("LX_APP_ID", "ak_B1P0qz2mkImfS")
LX_SECRET = os.environ.get("LX_SECRET", "IMJm0f/dwDM7YYR+2FrlEQ==")
AUTH_TOKEN = os.environ.get("AUTH_TOKEN", "")
FEISHU = "https://open.feishu.cn/open-apis"
LXHOST = "https://openapi.lingxing.com"

FIN_APP = "P9awbhG9faFstxsO1KZc9b9Qnxb"     # 财务报表汇总
UP_TBL = "tbl5Hvrty3oqLdIF"                 # 速卖通月度数据上传台
OVERVIEW_TBL = "tbltFK8vwdcrlfBa"           # 全渠道销售总览
FRANKIE = "ou_629ce01f4bc31de078e10fcb038dbf78"
ZWJ = "ou_274ee5199a763b7ec97980cd54e3fecb"  # 赵伟俊(速卖通负责人)

# SKU_MAP: 速卖通后台SKU → 领星ERP编码(赵伟俊2026-06-15确认: FF01-4波纹=FF01A-04 cg¥90; 余3个映射对)
SKU_MAP = {"FF01-4": "FF01A-04", "FF01-7": "FF01B-01", "FF01-5": "FF01A-05", "KS42-2-2": "KS42-2"}
FEE = ["佣金", "交易服务费", "联盟佣金", "金币营销费", "cashback", "营销技术服务费", "营销智投服务费", "新商孵化服务费"]
BACK = [c + "退回" for c in FEE]
app = FastAPI()


def tok():
    r = requests.post(f"{FEISHU}/auth/v3/tenant_access_token/internal",
                      json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}, timeout=20)
    return r.json()["tenant_access_token"]


def num(x):
    try:
        return float(str(x).replace(",", "")) if x not in (None, "") else 0.0
    except Exception:
        return 0.0


def dg(s):
    return "".join(c for c in str(s) if c.isdigit())


def gv(f, k):
    v = f.get(k)
    if v is None:
        return ""
    if isinstance(v, list):
        return " ".join(str(x.get("text", x.get("value", x)) if isinstance(x, dict) else x) for x in v)
    if isinstance(v, dict):
        return str(v.get("text") or v.get("value") or "")
    return str(v)


def getall(T, app_token, tbl):
    items = []; pt = None
    while True:
        u = f"{FEISHU}/bitable/v1/apps/{app_token}/tables/{tbl}/records?page_size=500" + (f"&page_token={pt}" if pt else "")
        d = requests.get(u, headers={"Authorization": f"Bearer {T}"}, timeout=30).json().get("data", {})
        items += (d.get("items") or []); pt = d.get("page_token")
        if not d.get("has_more"): break
    return items


# ===== 领星 cg_price (AES 签名) =====
def _lx_sign(p):
    ks = sorted(k for k in p if p[k] not in ('', None))
    s = "&".join(f"{k}={p[k]}" for k in ks)
    md5 = hashlib.md5(s.encode()).hexdigest().upper()
    key = LX_APP_ID.encode()[:16].ljust(16, b'\0')
    c = AES.new(key, AES.MODE_ECB); pad = 16 - len(md5) % 16
    return base64.b64encode(c.encrypt(md5.encode() + bytes([pad]) * pad)).decode()


def lx_cg():
    body = "appId=" + LX_APP_ID + "&appSecret=" + requests.utils.quote(LX_SECRET, safe="")
    t = requests.post(f"{LXHOST}/api/auth-server/oauth/access-token", data=body,
                      headers={"Content-Type": "application/x-www-form-urlencoded"}, timeout=30).json()
    access = t["data"]["access_token"]

    def post(path, b):
        ts = int(time.time()); base = {"access_token": access, "app_key": LX_APP_ID, "timestamp": ts}; sp = dict(base)
        for k, v in b.items(): sp[k] = json.dumps(v) if isinstance(v, (list, dict)) else str(v)
        qs = "&".join(f"{k}={requests.utils.quote(str(v), safe='')}" for k, v in {**base, "sign": _lx_sign(sp)}.items())
        return requests.post(f"{LXHOST}{path}?{qs}", data=json.dumps(b),
                             headers={"Content-Type": "application/json"}, timeout=60).json()
    cg = {}; off = 0
    while True:
        d = (post("/erp/sc/routing/data/local_inventory/productList", {"offset": off, "length": 200}).get("data") or [])
        for it in d:
            if it.get("sku"): cg[it["sku"]] = num(it.get("cg_price"))
        if len(d) < 200: break
        off += 200; time.sleep(0.3)
    return cg


# ===== 下载上传台附件 =====
def download_files(T, month):
    """返回 {'funlab':[path..], 'linyuvo':[path..], 'logi':[path..]} (按文件名路由)。"""
    mk = dg(month); out = {"funlab": [], "linyuvo": [], "logi": []}
    rec_id = None
    for r in getall(T, FIN_APP, UP_TBL):
        f = r["fields"]
        if dg(gv(f, "月份")) != mk: continue
        rec_id = r["record_id"]
        for a in (f.get("数据文件") or []):
            ftok = a.get("file_token"); fname = a.get("name", "")
            if not ftok: continue
            try:
                resp = requests.get(f"{FEISHU}/drive/v1/medias/{ftok}/download",
                                    headers={"Authorization": f"Bearer {T}"}, timeout=90)
                p = os.path.join(tempfile.gettempdir(), f"{mk}_{ftok}.xlsx")
                open(p, "wb").write(resp.content)
                low = fname.lower()
                if "物流" in fname or "logi" in low:
                    out["logi"].append(p)
                elif "资金明细" in fname or "资金" in fname:
                    if "linyuvo" in low or "联游" in fname: out["linyuvo"].append(p)
                    elif "funlab" in low or "纷岚" in fname: out["funlab"].append(p)
            except Exception as e:
                print("download fail", fname, e)
    return out, rec_id


def parse_capital(path, store):
    """读资金明细 子表「订单商品明细」→ rows。"""
    recs = []
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sns = [s for s in wb.sheetnames if "商品明细" in s]
        if not sns: wb.close(); return recs
        ws = wb[sns[0]]
        data = list(ws.iter_rows(values_only=True)); wb.close()
        if not data: return recs
        hdr = [("" if c is None else str(c).strip()) for c in data[0]]
        idx = {h: i for i, h in enumerate(hdr)}
        g = lambda r, k: (r[idx[k]] if k in idx and idx[k] < len(r) else None)
        for r in data[1:]:
            on = g(r, "订单号")
            if on in (None, ""): continue
            recs.append({"store": store, "on": str(on), "sku": str(g(r, "sku编码") or "").strip(),
                         "qty": num(g(r, "商品数量")), "cj": num(g(r, "成交金额")),
                         "refund": num(g(r, "售中退款金额")) + num(g(r, "售后退款金额")),
                         "fee": sum(num(g(r, c)) for c in FEE) - sum(num(g(r, c)) for c in BACK),
                         "payout": num(g(r, "放款金额")),  # 已放款金额=真回款(实收口径)
                         "st": str(g(r, "结算状态") or "").strip()})
    except Exception as e:
        print("parse_capital", path, e)
    return recs


def parse_logi(paths):
    wl = defaultdict(float)
    for p in paths:
        try:
            wb = openpyxl.load_workbook(p, data_only=True); ws = wb.worksheets[0]
            data = list(ws.iter_rows(values_only=True)); wb.close()
            if not data: continue
            hdr = [("" if c is None else str(c).strip()) for c in data[0]]
            idx = {h: i for i, h in enumerate(hdr)}
            for r in data[1:]:
                on = r[idx["交易单号"]] if "交易单号" in idx and idx["交易单号"] < len(r) else None
                amt = num(r[idx["计费金额"]]) if "计费金额" in idx and idx["计费金额"] < len(r) else 0
                if on: wl[str(on)] += amt
        except Exception as e:
            print("parse_logi", p, e)
    return wl


def compute(files, cg):
    recs = []
    for p in files["funlab"]: recs += parse_capital(p, "FUNLAB")
    for p in files["linyuvo"]: recs += parse_capital(p, "LinYuvo")
    wl = parse_logi(files["logi"])
    erp = lambda s: SKU_MAP.get(s, s)
    oq = defaultdict(float)
    for x in recs: oq[x["on"]] += x["qty"]
    # 已结算发货行的平台费率 → 估算待结算行平台费(待结算行fee=0,放款后才结算→任意日期算都准, 不用等16号)
    s_fee = sum(x["fee"] for x in recs if x["st"] == "已结算" and (x["cj"] - x["refund"]) > 0.5)
    s_cj = sum(x["cj"] for x in recs if x["st"] == "已结算" and (x["cj"] - x["refund"]) > 0.5)
    fee_rate = (s_fee / s_cj) if s_cj else 0.0
    tot = [0.0] * 8  # qty,cj,refund,fee,cg,wl,maoli,payout(回款)
    miss = defaultdict(float); pend = 0; washed = 0; fee_est = 0.0; agg = defaultdict(lambda: [0.0] * 8)
    for x in recs:
        es = erp(x["sku"]); net = x["cj"] - x["refund"]; shipped = net > 0.5
        cgc = (cg.get(es, 0) * x["qty"]) if shipped else 0
        if shipped and es and es not in cg: miss[f"{x['sku']}->{es}"] += x["cj"]
        w = (wl.get(x["on"], 0) * (x["qty"] / oq[x["on"]] if oq[x["on"]] else 0)) if shipped else 0
        fee = x["fee"]
        if x["st"] != "已结算":
            pend += 1
            if shipped and fee == 0:           # 待结算行平台费暂0 → 按已结算费率估算
                fee = x["cj"] * fee_rate; fee_est += fee
        if not shipped: washed += 1
        ml = x["cj"] - x["refund"] - fee - cgc - w
        vals = [x["qty"], x["cj"], x["refund"], fee, cgc, w, ml, x["payout"]]
        a = agg[(x["store"], es)]
        for i, v in enumerate(vals): a[i] += v; tot[i] += v
    return {"rows": len(recs), "pending": pend, "washed": washed, "logi_orders": len(wl),
            "tot": tot, "miss": dict(miss), "agg": agg, "fee_rate": fee_rate, "fee_est": fee_est}


# ===== 灌全渠道总表 (速卖通 summary 行, 幂等) =====
def upsert_overview(T, month, r):
    tot = r["tot"]
    net_sales = tot[1] - tot[2]                # 净成交
    maoli = tot[6]; payout = tot[7]            # payout=放款金额=真回款(实收口径A)
    fields = {
        "月份": month, "渠道大类": "跨境电商", "平台": "速卖通",
        "店铺": "FUNLAB+LinYuvo(2店)",
        "销量": round(tot[0]),
        "销售额RMB": round(net_sales, 2),
        "平台费RMB": round(tot[3], 2),
        "采购成本RMB": round(tot[4], 2),
        "物流费RMB": round(tot[5], 2),
        "全额毛利RMB": round(maoli, 2),
        "毛利率": round(maoli / net_sales, 4) if net_sales else 0,
        "回款RMB": round(payout, 2),            # 真回款=已放款金额(随放款进度增长)
        "回款率": round(payout / net_sales, 4) if net_sales else 0,
    }
    found = None
    for rec in getall(T, FIN_APP, OVERVIEW_TBL):
        f = rec["fields"]
        if gv(f, "月份") == month and gv(f, "平台") == "速卖通":
            found = rec["record_id"]; break
    H = {"Authorization": f"Bearer {T}", "Content-Type": "application/json"}
    if found:
        requests.put(f"{FEISHU}/bitable/v1/apps/{FIN_APP}/tables/{OVERVIEW_TBL}/records/{found}",
                     headers=H, json={"fields": fields}, timeout=30)
        return found, "updated"
    else:
        rr = requests.post(f"{FEISHU}/bitable/v1/apps/{FIN_APP}/tables/{OVERVIEW_TBL}/records",
                           headers=H, json={"fields": fields}, timeout=30).json()
        return rr.get("data", {}).get("record", {}).get("record_id"), "created"


def send_msg(T, oid, text):
    requests.post(f"{FEISHU}/im/v1/messages?receive_id_type=open_id",
                  headers={"Authorization": f"Bearer {T}", "Content-Type": "application/json"},
                  json={"receive_id": oid, "msg_type": "text", "content": json.dumps({"text": text}, ensure_ascii=False)}, timeout=20)


def mark_done(T, rec_id, summary):
    if not rec_id: return
    requests.put(f"{FEISHU}/bitable/v1/apps/{FIN_APP}/tables/{UP_TBL}/records/{rec_id}",
                 headers={"Authorization": f"Bearer {T}", "Content-Type": "application/json"},
                 json={"fields": {"处理状态": "已计算", "计算结果摘要": summary[:480]}}, timeout=30)


def do_recompute(month):
    T = tok()
    files, rec_id = download_files(T, month)
    nfiles = len(files["funlab"]) + len(files["linyuvo"]) + len(files["logi"])
    if not files["funlab"] and not files["linyuvo"]:
        msg = f"🟡 [FIN·P2] 速卖通毛利 · {month}\n上传台未找到 {month} 的资金明细文件，赵伟俊还没传。请提醒上传后重跑。"
        send_msg(T, FRANKIE, msg)
        return {"ok": False, "month": month, "reason": "no_files", "found": nfiles}
    cg = lx_cg()
    r = compute(files, cg)
    rid, act = upsert_overview(T, month, r)
    tot = r["tot"]; net = tot[1] - tot[2]
    top = sorted(r["agg"].items(), key=lambda x: -x[1][6])[:5]
    toptxt = " / ".join(f"{sk}{ml[6]:.0f}" for (st, sk), ml in top if sk)
    lines = [
        f"🟡 [FIN·P2] 速卖通毛利报表 · {month}",
        f"总成交¥{tot[1]:.0f} | 净成交¥{net:.0f} | 退款¥{tot[2]:.0f} | 平台费¥{tot[3]:.0f} | 采购¥{tot[4]:.0f} | 物流¥{tot[5]:.0f}",
        f"净毛利 ¥{tot[6]:.0f} ({tot[6]/net*100 if net else 0:.1f}%净成交) | 回款(已放款)¥{tot[7]:.0f} · 已{act}总表(跨境/速卖通)",
        f"Top: {toptxt}",
    ]
    if r["pending"]: lines.append(f"ℹ️ {r['pending']}行待结算,平台费已按已结算费率{r['fee_rate']*100:.1f}%估算¥{r['fee_est']:.0f}(任意日期算都准,无需等月底)")
    if r["miss"]: lines.append(f"⚠️ 领星缺cg: {r['miss']}")
    summary = "\n".join(lines)
    send_msg(T, FRANKIE, summary)
    mark_done(T, rec_id, summary)
    return {"ok": True, "month": month, "net_profit": round(tot[6]), "net_sales": round(net),
            "pending": r["pending"], "miss": r["miss"], "overview": act, "files": nfiles}


def do_reminder(month):
    T = tok()
    txt = (f"📦 速卖通月度数据提醒 · {month}\n威哥，{month} 的速卖通数据该导了。请从两个店(FUNLAB+LinYuvo)卖家后台导:\n"
           f"① 资金明细(俄罗斯+非俄罗斯) ② 物流账单\n传到飞书「速卖通月度数据上传台」(月份填 {month})。\n"
           f"系统每月16号自动算上月毛利，建议15号前传完(此时多数订单已放款，费用才准)。")
    send_msg(T, ZWJ, txt)
    return {"ok": True, "month": month, "target": "赵伟俊"}


def last_month():
    last = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    return last.strftime("%Y-%m")


@app.get("/health")
def health(): return {"ok": True}


@app.post("/recompute")
async def recompute(request: Request, month: str = ""):
    if AUTH_TOKEN and request.headers.get("Authorization") != f"Bearer {AUTH_TOKEN}":
        raise HTTPException(401, "unauthorized")
    return do_recompute(month or last_month())


@app.post("/send-reminder")
async def send_reminder(request: Request, month: str = ""):
    if AUTH_TOKEN and request.headers.get("Authorization") != f"Bearer {AUTH_TOKEN}":
        raise HTTPException(401, "unauthorized")
    return do_reminder(month or last_month())
