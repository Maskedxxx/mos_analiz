#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Прогон всех 14 multi-rule типов через ПРОДОВЫЙ engine.run().

Для каждого типа: берём один реальный документ из logs_result/*/original/
и запускаем полный audit. Проверяем что результат корректный (не краш).
"""
import os
import sys
import tempfile
import time
from pathlib import Path

os.environ["NO_PROXY"] = "172.16.10.35,localhost,127.0.0.1"

PROJECT = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(PROJECT))

from audit_engine.engine import AuditEngine

FIXTURES = {
    "prikaz_ppu":          "logs_result/prikaz_ppu/session_20260404_171205/original/3.10._Приказ о ППУ СпецТехРесурс НТ.docx",
    "prikaz_pa":           "logs_result/prikaz_pa/session_20260408_103202/original/2.2._Приказ о внедрении ПА с приложениями ООО ЮТС.docx",
    "prikaz_comp_ppu":     "logs_result/prikaz_comp_ppu/session_20260404_172642/original/3.10._Приказ о конкурсах ППУ СпецТехРесурс НТ.docx",
    "prikaz_vyhod":        "logs_result/prikaz_vyhod/session_20260410_072356/original/3.6._Приказ об проведении выхода ООО Гофромир.docx",
    "prikaz_tirazh":       "logs_result/prikaz_tirazh/session_20260329_182646/original/0_5_Приказ_о_переходе_на_этап_Тиражирования_с_приложениями_ООО_АПТОС.docx",
    "prikaz_ic_potoka_el": "uploads/53680497/1.3._Приказ о создании ИЦ потока в эл.виде с прилож._МТЭР ЦТС.docx",
    "prikaz_ic_el":        "uploads/6dbbddcf/1.5._Приказ о создании ИЦ с прилож. МТЭР ЦТС эл. вид.docx",
    "prikaz_ic":           "uploads/f6499d23/1.5._Приказ о создании ИЦ с прилож. МТЭР ЦТС эл. вид.docx",
    "prikaz_ic_potoka":    "uploads/a8555cc4/1.3._Приказ о создании ИЦ потока в эл.виде с прилож._МТЭР ЦТС.docx",
    "prikaz_otvetstvennyh": "logs_result/prikaz_otvetstvennyh/session_20260406_111253/original/0.2._Приказ об отв. _ООО Пром Энерго в.4.docx",
    "prikaz_formirovanie_po": "logs_result/prikaz_formirovanie_po/session_20260410_071429/original/3.5._Приказ о формировании ПО _ООО Гофромир_.docx",
    "polozhenie_comp_ppu": "logs_result/polozhenie_comp_ppu/session_20260407_105903/original/3.10._Положение о конкурсах проектов и ППУ (1).docx",
    "polozhenie_ppu":      "logs_result/polozhenie_ppu/session_20260402_111335/original/Положение о ППУ.docx",
    "polozhenie_po":       "logs_result/polozhenie_po/session_20260410_071844/original/3.5._Положение о ПО _ООО Гофромир_.docx",
    "akt_nachala":         "logs_result/akt_nachala/session_20260330_122942/original/0.1_Акт начала мероприятий ООО С.П. Гелпик.docx",
    "otchet_rezultatov":   "logs_result/otchet_rezultatov/session_20260402_063814/original/Отчет о результатах вскрытия резервов ООО ОПТИМА ИМПОРТ утв...pdf",
    "presentation_eu":     "logs_result/presentation_eu/session_20260403_082452/original/2.5._Создание ЭУ наименование предприятия.pptx",
    "cheklist_eu":         "logs_result/cheklist_eu/session_20260403_085256/original/2.5._Чек лист выбора ЭУ.docx",
}

SUMMARY = []

for doc_type, rel_path in FIXTURES.items():
    file_path = PROJECT / rel_path
    print(f"\n{'='*60}")
    print(f"▶ {doc_type}: {file_path.name}")
    print(f"{'='*60}")

    if not file_path.exists():
        SUMMARY.append((doc_type, "SKIP", "фикстура не найдена"))
        print(f"⚠ Фикстура не найдена")
        continue

    try:
        t0 = time.time()
        with tempfile.TemporaryDirectory() as tmpdir:
            engine = AuditEngine(doc_type)
            result = engine.run(str(file_path), session_dir=tmpdir)
            elapsed = time.time() - t0
            violations_count = len(result.violations) if result.violations else 0
            SUMMARY.append((doc_type, "OK", f"{violations_count} violations, {elapsed:.1f}s"))
            print(f"✓ {violations_count} violations за {elapsed:.1f}s")
    except Exception as e:
        SUMMARY.append((doc_type, "FAIL", f"{type(e).__name__}: {str(e)[:100]}"))
        print(f"✗ {type(e).__name__}: {e}")

print(f"\n\n{'='*60}")
print("ИТОГ:")
print(f"{'='*60}")
for doc_type, status, details in SUMMARY:
    mark = "✓" if status == "OK" else ("⚠" if status == "SKIP" else "✗")
    print(f"{mark} {doc_type:<25} {status:<5} {details}")

ok_count = sum(1 for _, s, _ in SUMMARY if s == "OK")
fail_count = sum(1 for _, s, _ in SUMMARY if s == "FAIL")
print(f"\nВсего: {ok_count} OK, {fail_count} FAIL, {len(SUMMARY)-ok_count-fail_count} SKIP")
sys.exit(0 if fail_count == 0 else 1)
