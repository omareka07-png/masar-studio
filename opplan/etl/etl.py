#!/usr/bin/env python3
"""
ETL لقالب «منظومة التخطيط التشغيلي والموازنة السنوية»
Extract  : يقرأ الأوراق DATA / كشاف الأهداف والمؤشرات / كشاف المشاريع / مشروع
Transform: ينظّف النصوص، يطبّع الرموز، يحذف القيم الشكلية (.............)،
           يستنتج نوع وحدة المؤشر، ويبني العلاقات (منظور ← هدف ← مؤشر ← مشروع ← خطوة)
Load     : يُخرج master.json و projects.json و quality_report.json
           (نفس البنية التي يقرأها النظام التفاعلي وتُحمّل في قاعدة البيانات)

الاستخدام:
    python3 -I etl.py <ملف.xlsx> <مجلد_الإخراج>
"""
import json, re, sys, unicodedata
from pathlib import Path
import openpyxl

PLACEHOLDER = re.compile(r"^[.\s…]+$")

PERSPECTIVES = {  # أسماء مستنتجة من بادئة رمز الهدف — قابلة للتعديل في النظام
    "س": "الأثر والمستفيدون",
    "ص": "أصحاب المصلحة",
    "م": "المالي",
    "ل": "العمليات الداخلية",
    "ت": "التعلم والنمو",
}


def clean(v):
    """نص نظيف: إزالة الفاصلة العليا البادئة والمسافات المكررة والطرفية وتوحيد يونيكود."""
    if v is None:
        return None
    if not isinstance(v, str):
        return v
    s = unicodedata.normalize("NFC", v).lstrip("'")
    s = re.sub(r"[\t ]+", " ", s)
    s = re.sub(r" {2,}", " ", s).strip()
    if not s or PLACEHOLDER.match(s):
        return None
    return s


def norm_code(code):
    """'س 1 - 1 ' → 'س1-1'"""
    return re.sub(r"\s+", "", code or "")


def unit_of(kpi_name):
    if kpi_name.startswith("نسبة"):
        return "نسبة مئوية"
    if kpi_name.startswith("عدد"):
        return "عدد"
    if "إيراد" in kpi_name:
        return "ريال"
    if kpi_name.startswith("مستوى"):
        return "مستوى"
    return "أخرى"


def main(src, out):
    out = Path(out); out.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(src, data_only=False)
    issues = []

    def issue(sev, sheet, cell, msg):
        issues.append({"severity": sev, "sheet": sheet, "cell": cell, "message": msg})

    # ---------- الأهداف والمؤشرات (المصدر الرئيس: كشاف الأهداف والمؤشرات) ----------
    ws = wb["كشاف الأهداف والمؤشرات "]
    objectives, kpis, seen_obj = [], [], {}
    for r in range(3, ws.max_row + 1):
        oc, on, kc, kn = (ws.cell(r, c).value for c in (4, 5, 6, 7))
        oc, on, kn = norm_code(clean(oc)), clean(on), clean(kn)
        kc_raw = kc
        kc = norm_code(clean(kc))
        if not (oc and on and kn):
            continue
        if kc_raw and kc_raw != kc_raw.strip():
            issue("منخفض", ws.title, f"F{r}", f"مسافة زائدة في رمز المؤشر «{kc_raw}»")
        if oc not in seen_obj:
            p = oc[0]
            seen_obj[oc] = {"id": oc, "code": oc, "name": on, "perspective": p}
            objectives.append(seen_obj[oc])
        elif seen_obj[oc]["name"] != on:
            issue("مرتفع", ws.title, f"E{r}", f"الرمز {oc} مرتبط بأكثر من اسم هدف")
        kpis.append({"id": kc, "code": kc, "name": kn, "objectiveId": oc, "unit": unit_of(kn)})

    # تحقق تقاطعي مع ورقة DATA (عمود F)
    data = wb["DATA"]
    data_kpis = [clean(data.cell(r, 6).value) for r in range(6, 52)]
    data_kpis = [k for k in data_kpis if k]
    if len(data_kpis) != len(kpis):
        issue("مرتفع", "DATA", "F6:F51", f"عدد المؤشرات في DATA ({len(data_kpis)}) لا يطابق الكشاف ({len(kpis)})")
    dup_cols = all(data.cell(r, 3).value == data.cell(r, 5).value for r in range(6, 52))
    if dup_cols:
        issue("منخفض", "DATA", "C:E", "عمودا الهدف C و E متطابقان (تكرار غير مطبّع)")

    # ---------- الإدارات والأقسام والأولويات ----------
    departments = []
    for r in range(6, 18):
        raw = data.cell(r, 8).value
        name = clean(raw)
        if not name:
            continue
        if isinstance(raw, str) and (raw.startswith("'") or raw != raw.strip() or "  " in raw):
            issue("منخفض", "DATA", f"H{r}", f"تنسيق نص الإدارة «{raw}» (فاصلة/مسافات)")
        departments.append({"id": f"D{len(departments)+1:02d}", "name": name})
    sections = [clean(data.cell(r, 11).value) for r in range(6, 36)]
    if all(s and re.match(r"^إدارة\d+ القسم\d+$", s) for s in sections if s):
        issue("متوسط", "DATA", "J6:K35", "جدول الأقسام قيم افتراضية (إدارة1 القسم1…) غير مرتبطة بالإدارات الفعلية — لم تُرحّل")
    priorities = [{"id": 1, "name": "أولوية عالية"}, {"id": 2, "name": "أولوية متوسطة"}, {"id": 3, "name": "أولوية منخفضة"}]

    # ---------- كشاف المشاريع ----------
    reg = wb["كشاف المشاريع"]
    months = [clean(reg.cell(4, c).value) for c in range(8, 20)]
    projects = []
    placeholder_rows = 0
    for r in range(5, 21):
        name = clean(reg.cell(r, 3).value)
        if name and name.startswith("="):
            # الصف الأول مربوط بورقة «مشروع»
            name = clean(wb["مشروع"]["C2"].value)
        if not name or name == "الخطة":
            placeholder_rows += 1
            continue
        monthly = [float(reg.cell(r, c).value or 0) for c in range(8, 20)]
        projects.append({
            "name": name, "description": clean(reg.cell(r, 4).value) or "",
            "manager": clean(reg.cell(r, 5).value) or "",
            "priority": int(reg.cell(r, 2).value or 2),
            "progress": float(reg.cell(r, 6).value or 0) * (100 if (reg.cell(r, 6).value or 0) <= 1 else 1),
            "budget": float(reg.cell(r, 7).value or 0) if not str(reg.cell(r, 7).value).startswith("=") else sum(monthly),
            "monthly": monthly, "steps": [],
        })
    issue("معلومة", "كشاف المشاريع", "A5:T20", f"{placeholder_rows} صفاً شكلياً (.............) تم تجاهلها — لا توجد مشاريع فعلية في الملف")
    for r in range(7, 20):
        if isinstance(reg.cell(r, 1).value, (int, float)):
            issue("منخفض", "كشاف المشاريع", f"A{r}", "رقم تسلسلي مكتوب يدوياً بدل المعادلة — يختل عند الفرز")
            break
    if clean(reg["T4"].value) == "إجمالي الموزانة":
        issue("منخفض", "كشاف المشاريع", "T4", "خطأ إملائي في العنوان: «الموزانة» ← «الموازنة»")

    # ---------- بطاقة المشروع (ورقة مشروع) ----------
    card = wb["مشروع"]
    if card["M10"].value and getattr(card["M10"].value, "year", 2026) != 2026:
        issue("مرتفع", "مشروع", "M8:M10", f"تقويم الأسابيع يبدأ في {card['M10'].value:%Y-%m-%d} (عام {card['M10'].value.year}) والملف لخطة 2026")
    if card["BS10"].value == "=BK10+7":
        issue("مرتفع", "مشروع", "BS10", "المعادلة =BK10+7 يجب أن تكون =BR10+7 (تاريخ الأسبوع الأخير خاطئ)")
    if card["BS11"].value == 52:
        issue("متوسط", "مشروع", "BS11", "رقم الأسبوع 52 مكرر بعد الأسبوع 58 — التقويم يتجاوز 52 أسبوعاً")
    if card["A20"].value and "A13" in str(card["A20"].value):
        issue("متوسط", "مشروع", "A20", "الترقيم يقفز من A13 إلى A20 (الصفوف 14-19 بلا تكلفة أو أسابيع)")
    issue("متوسط", "مشروع", "J12:J30", "نسبة الإنجاز = مدة التنفيذ ÷ مدة الخطة: تقيس الزمن المستهلك لا الإنجاز الفعلي")
    issue("متوسط", "مشروع", "—", "بطاقة مشروع واحدة لكل ورقة: لا يتوسع لعشرات المشاريع ولا يدعم التجميع")

    broken = [n for n, d in wb.defined_names.items() if "#REF!" in (d.attr_text or "") or "#N/A" in (d.attr_text or "")]
    external = [n for n, d in wb.defined_names.items() if re.search(r"\[\d+\]", d.attr_text or "")]
    issue("مرتفع", "Workbook", "Names", f"{len(broken)} اسماً معرّفاً معطوباً (#REF!/#N/A) تُغذي تنسيق مخطط جانت فلا يظهر التظليل")
    issue("متوسط", "Workbook", "External links", f"{len(external)} اسماً يشير إلى مصنفات خارجية غير موجودة ([2]…[9])")

    master = {
        "perspectives": [{"id": k, "name": v, "inferred": True} for k, v in PERSPECTIVES.items()],
        "objectives": objectives, "kpis": kpis, "departments": departments,
        "priorities": priorities, "months": months,
    }
    (out / "master.json").write_text(json.dumps(master, ensure_ascii=False, indent=1), encoding="utf-8")
    (out / "projects.json").write_text(json.dumps(projects, ensure_ascii=False, indent=1), encoding="utf-8")
    report = {
        "source": Path(src).name,
        "counts": {"sheets": len(wb.sheetnames), "perspectives": len(PERSPECTIVES), "objectives": len(objectives),
                   "kpis": len(kpis), "departments": len(departments), "projects": len(projects),
                   "placeholder_rows": placeholder_rows, "broken_names": len(broken), "external_names": len(external)},
        "issues": issues,
    }
    (out / "quality_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(report["counts"], ensure_ascii=False))
    for i in issues:
        print(f"[{i['severity']}] {i['sheet']}!{i['cell']}: {i['message']}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
