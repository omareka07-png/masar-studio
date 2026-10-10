#!/usr/bin/env python3
"""يبني صفحة النظام من المصدر src.html بحقن البيانات المرجعية الناتجة عن ETL.
الإخراج:
  - ../index.html          نسخة مستقلة تعمل على GitHub Pages أو أي خادم ملفات (وضع محلي)
  - opplan.html            نسخة Artifact (تُنشر على claude.ai وتستخدم قاعدة البيانات المشتركة)
"""
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
DATA = HERE.parent / "data"
src = (HERE / "src.html").read_text(encoding="utf-8")
master = json.loads((DATA / "master.json").read_text(encoding="utf-8"))
issues = json.loads((DATA / "quality_report.json").read_text(encoding="utf-8"))["issues"]
page = (src.replace("/*MASTER*/null", json.dumps(master, ensure_ascii=False))
           .replace("/*ISSUES*/[]", json.dumps(issues, ensure_ascii=False)))
(HERE / "opplan.html").write_text(page, encoding="utf-8")

shell_head = ('<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8">'
              '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
              '<style>:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}'
              'body{margin:0;font:14px system-ui,sans-serif}img{max-width:100%}[hidden]{display:none!important}</style>'
              '</head><body>')
(HERE.parent / "index.html").write_text(shell_head + page + "</body></html>", encoding="utf-8")
print("built", len(page), "bytes")
