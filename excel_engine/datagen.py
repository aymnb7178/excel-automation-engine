"""Generate MESSY SYNTHETIC sales files (xlsx + csv) with known injected problems.

Everything here is invented: names are random combinations from small lists, prices are random.
The returned ``truth`` dict records exactly what was injected so tests/demo can compare it with
what the engine detects.
"""

from __future__ import annotations

import csv
import random
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

CITIES = ["Riyadh", "Jeddah", "Makkah", "Dammam", "Madinah"]
CITY_AR = {"Riyadh": "الرياض", "Jeddah": "جدة", "Makkah": "مكة المكرمة", "Dammam": "الدمام", "Madinah": "المدينة المنورة"}
FIRST = [("Ahmed", "أحمد"), ("Mohammed", "محمد"), ("Fatima", "فاطمة"), ("Khalid", "خالد"),
         ("Sara", "سارة"), ("Omar", "عمر"), ("Noura", "نورة"), ("Yusuf", "يوسف")]
LAST = [("Alharbi", "الحربي"), ("Alotaibi", "العتيبي"), ("Alghamdi", "الغامدي"),
        ("Alzahrani", "الزهراني"), ("Alqahtani", "القحطاني"), ("Alshehri", "الشهري")]
PRODUCTS = {"Laptop": 3200, "Monitor": 850, "Keyboard": 120, "Mouse": 65,
            "Printer": 780, "Router": 240, "Webcam": 210, "Headset": 310}
STATUSES = ["Delivered", "Pending", "Cancelled", "Returned"]
STATUS_AR = {"Delivered": "تم التسليم", "Pending": "قيد الانتظار", "Cancelled": "ملغي", "Returned": "مرتجع"}
STATUS_LEGACY = {"Delivered": "done", "Pending": "pending", "Cancelled": "canceled", "Returned": "returned"}
BAD_DATES = ["31/02/2025", "13/13/2025", "TBD", "next week", "32/01/2025", "00/00/2025"]
_AR_DIGITS = str.maketrans("0123456789,.", "٠١٢٣٤٥٦٧٨٩٬٫")

F_NORTH, F_RIYADH, F_ARABIC, F_WEST, F_LEGACY = (
    "sales_north.xlsx", "sales_riyadh_branch.csv", "sales_arabic_headers.xlsx",
    "sales_west_two_sheets.xlsx", "sales_east_legacy.csv")


def _ar(s: str) -> str:
    return s.translate(_AR_DIGITS)


def generate_messy_inputs(out_dir: str | Path, seed: int = 42, n_orders: int = 500) -> dict:
    """Write 5 messy files into out_dir and return the ground-truth counts of what was injected."""
    rng = random.Random(seed)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    files = [F_NORTH, F_RIYADH, F_ARABIC, F_WEST, F_LEGACY]
    recs = []
    for i in range(n_orders):
        prod = rng.choice(list(PRODUCTS))
        recs.append({"order_id": f"ORD-{1001 + i}", "first": rng.randrange(len(FIRST)), "last": rng.randrange(len(LAST)),
                     "region": rng.choice(CITIES), "product": prod, "qty": rng.randint(1, 20),
                     "price": round(PRODUCTS[prod] * rng.uniform(0.9, 1.1), 2),
                     "date": date(2025, 1, 1) + timedelta(days=rng.randrange(365)),
                     "status": rng.choices(STATUSES, [0.7, 0.15, 0.1, 0.05])[0],
                     "file": rng.choices(files, [0.3, 0.2, 0.2, 0.15, 0.15])[0]})
    order = rng.sample(range(n_orders), n_orders)
    # Arabic-file records are never duplicated (their customer text is Arabic, so they could not match).
    pool = [i for i in order if recs[i]["file"] != F_ARABIC]
    dup_src, fuzzy_src, conflict_src = pool[:30], pool[30:42], pool[42:46]
    used = set(pool[:46])
    rest = [i for i in order if i not in used]
    plan = [("date_invalid", 6), ("date_missing", 6), ("date_year_typo", 4), ("qty_negative", 5),
            ("qty_text", 4), ("price_free", 3), ("price_zero", 3), ("customer_missing", 8), ("status_unknown", 5)]
    pos = 0
    for kind, k in plan:
        for i in rest[pos:pos + k]:
            r = recs[i]
            if kind == "date_invalid":
                r["date"] = rng.choice(BAD_DATES)
            elif kind == "date_missing":
                r["date"] = ""
            elif kind == "date_year_typo":
                r["date"] = r["date"].replace(year=2052)
            elif kind == "qty_negative":
                r["qty"] = -rng.randint(1, 5)
            elif kind == "qty_text":
                r["qty"] = rng.choice(["abc", "n/a"])
            elif kind == "price_free":
                r["price"] = "free"
            elif kind == "price_zero":
                r["price"] = 0
            elif kind == "customer_missing":
                r["no_customer"] = True
            elif kind == "status_unknown":
                r["status"] = "Shipped??"
        pos += k
    extras = []

    def other_file(src):
        return rng.choice([f for f in files if f not in (src, F_ARABIC)])

    for i in dup_src:
        extras.append({**recs[i], "file": other_file(recs[i]["file"])})
    for i in fuzzy_src:
        extras.append({**recs[i], "order_id": f"ORD-{9000 + len(extras)}", "typo": True, "file": other_file(recs[i]["file"])})
    for i in conflict_src:
        extras.append({**recs[i], "qty": recs[i]["qty"] + 1, "file": other_file(recs[i]["file"])})
    for k in range(3):  # brand-new rows that have no order id at all
        prod = rng.choice(list(PRODUCTS))
        extras.append({"order_id": None, "first": rng.randrange(len(FIRST)), "last": rng.randrange(len(LAST)),
                       "region": rng.choice(CITIES), "product": prod, "qty": rng.randint(1, 9),
                       "price": round(PRODUCTS[prod] * 1.0, 2), "date": date(2025, 6, 1) + timedelta(days=k),
                       "status": "Pending", "file": rng.choice([F_NORTH, F_RIYADH, F_LEGACY])})
    allrecs = recs + extras
    rng.shuffle(allrecs)

    def name(r, arabic):
        if r.get("no_customer"):
            return rng.choice(["", "N/A", "-"])
        f, l = FIRST[r["first"]], LAST[r["last"]]
        s = f"{f[1]} {l[1]}" if arabic else f"{f[0]} {l[0]}"
        if arabic and rng.random() < 0.4:
            s = s.replace("أ", "ا")  # alef variant
        if r.get("typo"):
            p = len(s) // 2
            p += s[p] == " "
            s = s[:p] + s[p + 1:]
        return s

    def messy_case(s: str, style: int) -> str:
        return [s, s.upper(), s.lower(), f"  {s} "][style]

    # ---- file 1: north (xlsx, banner rows, native types, extra Notes column) ----
    wb = Workbook(); ws = wb.active; ws.title = "Sales"
    ws.append(["Northern branch sales export (SYNTHETIC DEMO DATA)"]); ws.append([])
    ws.append(["Order ID", "Customer", "Region", "Product", "Qty", "Unit Price", "Order Date", "Status", "Notes"])
    # ---- file 4: west (two sheets, typo headers) ----
    wb4 = Workbook(); q1 = wb4.active; q1.title = "Q1"; q2 = wb4.create_sheet("Q2")
    for w in (q1, q2):
        w.append(["Ordr ID", "Custmer", "Regon", "Product", "Quantity", "Unit Prise", "Order Date", "Status"])
    # ---- file 3: arabic ----
    wb3 = Workbook(); w3 = wb3.active; w3.title = "المبيعات"
    w3.append(["رقم الطلب", "العميل", "المنطقة", "المنتج", "الكمية", "سعر الوحدة", "تاريخ الطلب", "الحالة"])
    f2 = [["order_no", "client name", "area", "item", "quantity", "price (SAR)", "date", "state"]]
    f5 = [["ID", "Customer Name", "City", "Product", "Qty", "Price", "Order Date", "Status"]]
    counts = {f: 0 for f in files}
    for r in allrecs:
        f, oid, st = r["file"], r["order_id"], r["status"]
        d, qty, price = r["date"], r["qty"], r["price"]
        counts[f] += 1
        if f == F_NORTH:
            ws.append([oid, messy_case(name(r, False), rng.randrange(4)), r["region"], r["product"],
                       qty, price, datetime(d.year, d.month, d.day) if isinstance(d, date) else d, st,
                       rng.choice(["", "", "rush", "gift"])])
        elif f == F_RIYADH:
            ds = d.strftime("%d/%m/%Y") if isinstance(d, date) else d
            ps = price if isinstance(price, str) else rng.choice([f"SAR {price:,.2f}", f"{price:.2f}"])
            f2.append([oid, messy_case(name(r, False), rng.randrange(4)), r["region"].upper() if rng.random() < .3 else r["region"].lower(),
                       r["product"].lower(), qty, ps, ds, st.lower()])
        elif f == F_ARABIC:
            ds = _ar(d.strftime("%d/%m/%Y")) if isinstance(d, date) else d
            ps = price if isinstance(price, str) else _ar(f"{price:,.2f} ر.س")
            qs = _ar(str(qty)) if isinstance(qty, int) else qty
            w3.append([_ar(oid) if oid else None, name(r, True), CITY_AR[r["region"]], r["product"], qs, ps, ds,
                       STATUS_AR.get(st, st)])
        elif f == F_WEST:
            ds = d.strftime("%d %b %Y") if isinstance(d, date) else d
            ps = price if isinstance(price, str) else f"{price:.2f} SAR"
            sheet = q1 if (isinstance(d, date) and d.month <= 6 and d.year == 2025) else q2
            sheet.append([oid, name(r, False), r["region"], r["product"], qty, ps, ds, st])
        else:
            ds = d.strftime(rng.choice(["%Y-%m-%d", "%Y/%m/%d"])) if isinstance(d, date) else d
            f5.append([oid, name(r, False), r["region"], r["product"], qty, price, ds, STATUS_LEGACY.get(st, st)])
    wb.save(out / F_NORTH); wb3.save(out / F_ARABIC); wb4.save(out / F_WEST)
    with open(out / F_RIYADH, "w", newline="", encoding="utf-8-sig") as fh:
        csv.writer(fh, delimiter=";").writerows(f2)
    with open(out / F_LEGACY, "w", newline="", encoding="utf-8-sig") as fh:
        csv.writer(fh).writerows(f5)
    return {"seed": seed, "files": files, "data_rows_written": len(allrecs), "rows_per_file": counts,
            "base_orders": n_orders,
            "injected": {"exact_duplicates": len(dup_src), "fuzzy_duplicates": len(fuzzy_src),
                         "conflicting_order_ids": len(conflict_src), "missing_order_id": 3,
                         **{k: n for k, n in plan}}}
