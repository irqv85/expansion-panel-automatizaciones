"""
Lista las reuniones organizadas por "Sofia" en el calendario de Outlook de
el responsable comercial, con inicio entre hoy 00:00 y manana 23:59 (hora local).

Solo lee el calendario, no decide nada de negocio (eso lo hace quien llame a
este script: matching de empresa, contra vTiger, generacion de deck, etc.).

Salida: JSON por stdout, lista de objetos:
  {"subject": "...", "start": "2026-08-19T10:00:00", "organizer": "..."}

Ejecutar con: python listar_reuniones_sofia.py
"""
import json
import sys
from datetime import datetime, timedelta

import win32com.client as win32

SOFIA_MARKER = "sofia"
olFolderCalendar = 9


def to_py_datetime(com_date):
    return datetime(
        com_date.year, com_date.month, com_date.day,
        com_date.hour, com_date.minute, com_date.second,
    )


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    now = datetime.now()
    window_start = datetime(now.year, now.month, now.day)
    window_end = window_start + timedelta(days=2)  # hoy + manana, exclusivo al final

    outlook = win32.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    calendar = ns.GetDefaultFolder(olFolderCalendar)
    items = calendar.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True

    encontradas = []
    for item in items:
        try:
            item_start = to_py_datetime(item.Start)
        except Exception:
            continue
        if item_start < window_start:
            continue
        if item_start >= window_end:
            break
        organizer_text = (getattr(item, "Organizer", "") or "").lower()
        if SOFIA_MARKER not in organizer_text:
            continue
        encontradas.append({
            "subject": getattr(item, "Subject", "") or "",
            "start": item_start.isoformat(),
            "organizer": getattr(item, "Organizer", "") or "",
        })

    json.dump(encontradas, sys.stdout, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        json.dump({"error": str(exc)}, sys.stdout)
        sys.exit(1)
