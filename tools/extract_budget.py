"""แปลง "รายงานประมาณการรายจ่าย" จากไฟล์ข้อบัญญัติงบประมาณ (PDF จากระบบ e-LAAS) เป็น JSON

ใช้งาน:
    python tools/extract_budget.py a_230926_103242.pdf data/budget_2570.json

ถ้าไฟล์ปลายทางลงท้าย .js จะเขียนเป็น `window.BUDGET_DATA = {...}` ให้หน้าเว็บโหลดผ่าน file:// ได้

หน้ารายงานในไฟล์เป็นแนวนอน (rotate 90) ข้อความมีชั้น text จริง จึงอ่านตำแหน่ง span ได้โดยตรง
หลังหมุน: bbox[0] = ตำแหน่งแถว, bbox[1] = ตำแหน่งคอลัมน์
"""
import json
import re
import sys

import pymupdf

# ตำแหน่งคอลัมน์ตัวเลข (bbox[1]) -> ชื่อฟิลด์
COLUMNS = {43: "y2570", 128: "diff_pct", 196: "y2569", 281: "y2568", 376: "y2567"}
NUM_RE = re.compile(r"^-?[\d,]+(\.\d+)?( %)?$")
HEADER_TEXTS = {"ปี 2567", "ปี 2568", "ปี 2569", "ปี 2570", "ประมาณการ", "รายการ", "ยอดต่าง (%)", "รายจ่ายจริง",
                "รายงานประมาณการรายจ่าย", "องค์การบริหารส่วนตำบลปุโรง"}
GROUP_NAMES = ("งบบุคลากร", "งบดำเนินงาน", "งบลงทุน", "งบเงินอุดหนุน", "งบรายจ่ายอื่น", "งบกลาง")


# ข้อมูลจากหน้าสแกน (หน้า 4-5, 7, 9-10) ไม่มีชั้นข้อความ จึงกรอกจากการอ่านภาพ (ภายหลังแทนด้วย OCR)
SCANNED_SUMMARY = {
    "title": "ข้อบัญญัติงบประมาณรายจ่าย ประจำปีงบประมาณ พ.ศ. 2570",
    "org": "องค์การบริหารส่วนตำบลปุโรง อำเภอกรงปินัง จังหวัดยะลา",
    "announced": "22 กันยายน 2569",
    "effective": "1 ตุลาคม 2569",
    "area_km2": 17.88,
    "population": {"total": 6041, "male": 3015, "female": 3026, "as_of": "7 กันยายน 2569"},
    "fiscal_status_2569": {"as_of": "18 สิงหาคม 2569", "bank_deposit": 35693987.05, "reserve": 75428377.49, "loans": 0},
    "revenue": [
        {"group": "รายได้จัดเก็บเอง", "name": "หมวดภาษีอากร", "y2568": 32905.90, "y2569": 33500, "y2570": 33500},
        {"group": "รายได้จัดเก็บเอง", "name": "หมวดค่าธรรมเนียม ค่าปรับ และใบอนุญาต", "y2568": 200442.00, "y2569": 42200, "y2570": 57200},
        {"group": "รายได้จัดเก็บเอง", "name": "หมวดรายได้จากทรัพย์สิน", "y2568": 238771.53, "y2569": 263700, "y2570": 263700},
        {"group": "รายได้จัดเก็บเอง", "name": "หมวดรายได้เบ็ดเตล็ด", "y2568": 24982.00, "y2569": 23500, "y2570": 26500},
        {"group": "รายได้ที่รัฐบาลเก็บแล้วจัดสรร", "name": "หมวดภาษีจัดสรร", "y2568": 21771931.87, "y2569": 21637100, "y2570": 23349100},
        {"group": "รายได้ที่รัฐบาลอุดหนุน", "name": "หมวดเงินอุดหนุน", "y2568": 19763513.30, "y2569": 21000000, "y2570": 21570000},
    ],
    # หน้า 10: รายจ่ายจำแนกตามงบ (ใช้ตรวจทานยอดที่แยกได้จากรายงาน)
    "expenditure_by_group": [
        {"name": "งบกลาง", "y2568": 6516732.00, "y2569": 7838000, "y2570": 7970400},
        {"name": "งบบุคลากร", "y2568": 14275602.34, "y2569": 16826200, "y2570": 18086436},
        {"name": "งบดำเนินงาน", "y2568": 11601013.36, "y2569": 12350100, "y2570": 12717364},
        {"name": "งบลงทุน", "y2568": 641280.00, "y2569": 3390900, "y2570": 3873800},
        {"name": "งบเงินอุดหนุน", "y2568": 2495000.00, "y2569": 2594800, "y2570": 2632000},
        {"name": "งบรายจ่ายอื่น", "y2568": 0, "y2569": 0, "y2570": 20000},
    ],
    "sections": [
        {"pages": "1", "name": "ปก", "kind": "image"},
        {"pages": "2", "name": "สารบัญ", "kind": "text"},
        {"pages": "3", "name": "ประกาศใช้ข้อบัญญัติ", "kind": "image"},
        {"pages": "4-5", "name": "คำแถลงงบประมาณ (สถานะการคลัง)", "kind": "image"},
        {"pages": "7-10", "name": "คำแถลงประกอบงบประมาณ (รายรับ/รายจ่าย)", "kind": "image"},
        {"pages": "12-17", "name": "บันทึกหลักการและเหตุผล / รายจ่ายตามแผนงาน", "kind": "image"},
        {"pages": "18-22", "name": "ข้อบัญญัติ (ลงนาม)", "kind": "image"},
        {"pages": "23-73", "name": "รายงานประมาณการรายจ่าย", "kind": "text"},
        {"pages": "74-259", "name": "รายงานรายละเอียดประมาณการรายจ่าย", "kind": "text"},
        {"pages": "260-389", "name": "ภาคผนวก (เอกสารสแกน)", "kind": "image"},
    ],
}


def to_number(text):
    return float(text.replace(",", "").replace("%", "").strip())


def nearest_column(pos):
    col = min(COLUMNS, key=lambda c: abs(c - pos))
    return COLUMNS[col] if abs(col - pos) < 30 else None


def page_rows(page):
    """คืนรายการแถว [(ตำแหน่ง, label, bold, {ค่าตัวเลข})] เรียงจากบนลงล่าง"""
    nums, texts = [], []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                t = span["text"].strip()
                if not t:
                    continue
                x0, y0, x1, _ = span["bbox"]
                center = (x0 + x1) / 2
                if y0 < 450 and NUM_RE.match(t):
                    field = nearest_column(y0)
                    if field:
                        nums.append((center, field, to_number(t)))
                elif y0 >= 450 and x0 < 700 and t not in HEADER_TEXTS and not t.startswith(("ประจำปีงบประมาณ", "อำเภอ", "วันที่พิมพ์", "หน้าที่")):
                    texts.append((center, t, "Bold" in span["font"]))

    # รวมตัวเลขที่อยู่แถวเดียวกัน
    rows = []
    for center, field, value in sorted(nums):
        if rows and abs(rows[-1]["pos"] - center) < 6:
            rows[-1]["values"][field] = value
        else:
            rows.append({"pos": center, "values": {field: value}, "label": [], "bold": False})

    # จับข้อความเข้ากับแถวตัวเลขที่ใกล้ที่สุด (ป้ายหลายบรรทัดอยู่ห่าง ±8) ที่เหลือคือหัวข้อ
    headings = []
    for center, text, bold in sorted(texts):
        row = min(rows, key=lambda r: abs(r["pos"] - center), default=None)
        if row and abs(row["pos"] - center) <= 18:
            row["label"].append((center, text))
            row["bold"] = row["bold"] or bold
        elif headings and center - headings[-1]["pos"] <= 18:  # หัวข้อยาวที่ขึ้นบรรทัดใหม่
            headings[-1]["label"].append((center, text))
        else:
            headings.append({"pos": center, "label": [(center, text)], "values": None, "bold": bold})

    out = []
    for r in sorted(rows + headings, key=lambda r: r["pos"]):
        label = " ".join(t for _, t in sorted(r["label"])).strip()
        out.append((r["pos"], label, r["bold"], r["values"]))
    return out


def extract(pdf_path):
    doc = pymupdf.open(pdf_path)
    # หน้าแรกของรายงานมีชื่อรายงาน + "หน้าที่ 1/N" หน้าถัดไปมีแค่เลขหน้า จึงนับต่อไป N หน้า
    report_pages = []
    for i, p in enumerate(doc):
        text = p.get_text()
        m = re.search(r"หน้าที่ 1/(\d+)", text)
        if m and "รายงานประมาณการรายจ่าย" in text and "รายละเอียด" not in text:
            report_pages = list(range(i, i + int(m.group(1))))
            break

    plans, items = [], []
    plan = work = group = category = None
    for i in report_pages:
        for _, label, bold, values in page_rows(doc[i]):
            if not label:
                continue
            if values is None:  # หัวข้อ
                if label.startswith("แผนงาน"):
                    plan = {"name": label, "works": [], "total": None}
                    plans.append(plan)
                    work = group = category = None
                elif label.startswith("งาน") or (label == "งบกลาง" and work is None):
                    work = {"name": label, "total": None}
                    plan["works"].append(work)
                    group = category = None
                elif label in GROUP_NAMES:
                    group, category = label, None
                else:
                    category = label
                continue

            if label.startswith("รวม"):
                if plan and label == "รวม" + plan["name"]:
                    plan["total"] = values
                elif work and label == "รวม" + work["name"]:
                    work["total"] = values
                continue

            items.append({
                "page": i + 1,
                "plan": plan["name"] if plan else None,
                "work": work["name"] if work else None,
                "group": group,
                "category": category,
                "name": label,
                **{k: values.get(k, 0.0) for k in COLUMNS.values() if k != "diff_pct"},
            })
    return doc, report_pages, plans, items


def main():
    pdf_path, out_path = sys.argv[1], sys.argv[2]
    doc, report_pages, plans, items = extract(pdf_path)

    by_group = {}
    for it in items:
        by_group[it["group"]] = by_group.get(it["group"], 0) + it["y2570"]

    result = {
        "summary": SCANNED_SUMMARY,
        "source": {
            "file": pdf_path.replace("\\", "/").split("/")[-1],
            "pages": len(doc),
            "text_pages": sum(1 for p in doc if len(p.get_text().strip()) > 100),
            "report_pages": [report_pages[0] + 1, report_pages[-1] + 1] if report_pages else None,
        },
        "plans": [
            {
                "name": p["name"],
                "y2570": (p["total"] or {}).get("y2570"),
                "y2569": (p["total"] or {}).get("y2569"),
                "y2568": (p["total"] or {}).get("y2568"),
                "y2567": (p["total"] or {}).get("y2567"),
                "works": [{"name": w["name"], "y2570": (w["total"] or {}).get("y2570")} for w in p["works"]],
            }
            for p in plans
        ],
        "groups": by_group,
        "items": items,
    }
    payload = json.dumps(result, ensure_ascii=False, separators=(",", ":"))
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"window.BUDGET_DATA = {payload};\n" if out_path.endswith(".js") else payload)

    total = sum(p["y2570"] or 0 for p in result["plans"])
    print(f"report pages {result['source']['report_pages']}, plans {len(plans)}, items {len(items)}, total 2570 = {total:,.2f}")
    for p in result["plans"]:
        print(f"  {p['name']}: {p['y2570']}")


if __name__ == "__main__":
    main()
