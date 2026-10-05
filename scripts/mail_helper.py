"""Envio de correo con adjuntos via Outlook de escritorio (Windows).

Historia de este archivo: originalmente hablaba con Outlook por COM; al
migrar a Ubuntu se reescribio sobre `thunderbird -compose`, que no puede
enviar por si solo y dejaba el clic final en "Enviar" a quien lo abriera.
De vuelta en Windows recuperamos Outlook COM, asi que el envio vuelve a ser
completo y sin intervencion, que es lo que necesitan las rutinas
desatendidas.

Se conservan dos modos, porque los dos flujos existentes los usan:

  - send_mail(...)      envia de una vez (equivalente al comportamiento
                        anterior a Ubuntu, y lo que usan las rutinas).
  - display_mail(...)   abre la ventana de redaccion ya rellena y deja el
                        clic en "Enviar" al usuario, que es el
                        comportamiento que se acordo durante la etapa de
                        Thunderbird para envios supervisados.

`send_via_thunderbird` se mantiene como alias de display_mail para no
romper a quien todavia la llame; ver la nota en esa funcion.

Requiere: Outlook de escritorio instalado, con sesion iniciada, y pywin32
(`pip install pywin32`).
"""
import sys
import warnings
from pathlib import Path

IS_WINDOWS = sys.platform == "win32"

# Constantes de Outlook (olMailItem, olFormatHTML). Se escriben a mano para
# no depender de la generacion de la cache de tipos de win32com, que falla
# en instalaciones limpias.
OL_MAIL_ITEM = 0
OL_FORMAT_HTML = 2
OL_FORMAT_PLAIN = 1


def _outlook():
    """Devuelve la aplicacion Outlook, o explica con claridad que falta.

    Los dos modos de fallo son distintos y conviene distinguirlos: que no
    este pywin32 (problema de entorno, se arregla con pip) y que no este
    Outlook o no haya perfil configurado (problema de la maquina).
    """
    if not IS_WINDOWS:
        raise RuntimeError(
            "El envio por Outlook COM solo funciona en Windows. "
            f"Plataforma detectada: {sys.platform}."
        )
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "Falta pywin32, necesario para hablar con Outlook. "
            "Instalalo con: pip install pywin32"
        ) from exc
    try:
        return win32.Dispatch("Outlook.Application")
    except Exception as exc:
        raise RuntimeError(
            "No se pudo abrir Outlook. Verifica que Outlook de escritorio "
            "este instalado y con sesion iniciada en este usuario."
        ) from exc


def _addr_list(value):
    if isinstance(value, (list, tuple, set)):
        return list(value)
    if value is None:
        return []
    return [value]


def _joined(value):
    """Outlook espera los destinatarios separados por punto y coma."""
    return "; ".join(str(v) for v in _addr_list(value) if v)


def _build_item(to, subject, body, attachment=None, cc=None, html=False):
    outlook = _outlook()
    mail = outlook.CreateItem(OL_MAIL_ITEM)
    mail.To = _joined(to)
    if cc:
        mail.CC = _joined(cc)
    mail.Subject = str(subject)

    if html:
        mail.BodyFormat = OL_FORMAT_HTML
        mail.HTMLBody = str(body)
    else:
        mail.BodyFormat = OL_FORMAT_PLAIN
        mail.Body = str(body)

    if attachment:
        for ruta in _addr_list(attachment):
            ruta = Path(ruta).resolve()
            if not ruta.exists():
                raise FileNotFoundError(f"No existe el adjunto: {ruta}")
            # Outlook COM exige la ruta absoluta como string de Windows.
            mail.Attachments.Add(str(ruta))

    return mail


def send_mail(to, subject, body, attachment=None, cc=None, html=False, dry_run=False):
    """Envia el correo de una vez, sin intervencion del usuario.

    Devuelve el asunto enviado (util para logs). En dry_run no toca Outlook.
    """
    if dry_run:
        adjuntos = ", ".join(Path(a).name for a in _addr_list(attachment)) or "sin adjunto"
        print(f"[DRY-RUN] To={_joined(to)} CC={_joined(cc)} "
              f"Subject='{subject}' Adjunto={adjuntos}")
        return subject

    mail = _build_item(to, subject, body, attachment, cc, html)
    mail.Send()
    return subject


def display_mail(to, subject, body, attachment=None, cc=None, html=False, dry_run=False):
    """Abre la ventana de redaccion ya rellena y deja el clic en "Enviar".

    Para envios supervisados, donde se quiere revisar antes de mandar.
    """
    if dry_run:
        adjuntos = ", ".join(Path(a).name for a in _addr_list(attachment)) or "sin adjunto"
        print(f"[DRY-RUN] (ventana) To={_joined(to)} CC={_joined(cc)} "
              f"Subject='{subject}' Adjunto={adjuntos}")
        return subject

    mail = _build_item(to, subject, body, attachment, cc, html)
    mail.Display(False)  # False = no modal, no bloquea el script
    return subject


def send_via_thunderbird(to, subject, body, attachment=None, cc=None, dry_run=False):
    """Alias obsoleto de display_mail, de la epoca de Ubuntu/Thunderbird.

    Se conserva para no romper llamadas existentes. Mantiene el
    comportamiento de entonces (abre la ventana, no envia solo) para que
    ningun flujo empiece a mandar correos sin supervision por el simple
    hecho de haber migrado de sistema operativo. Los flujos que deban
    enviar sin intervencion tienen que llamar a send_mail de forma
    explicita.
    """
    warnings.warn(
        "send_via_thunderbird esta obsoleto; usa display_mail (misma "
        "conducta) o send_mail (envio directo).",
        DeprecationWarning,
        stacklevel=2,
    )
    return display_mail(to=to, subject=subject, body=body,
                        attachment=attachment, cc=cc, dry_run=dry_run)


def build_ics(subject, start, end, location="", organizer_email="", attendee_email="", out_path=None):
    """Genera un archivo .ics minimo para adjuntar una invitacion de reunion
    (uso: reenvio de invitacion de Sofia, ver Sofia - Schedule.py).

    start/end: datetime (naive, se asume hora local).
    """
    def _fmt(dt):
        return dt.strftime("%Y%m%dT%H%M%S")

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//GB Advisors//Sofia Schedule//ES",
        "BEGIN:VEVENT",
        f"SUMMARY:{subject}",
        f"DTSTART:{_fmt(start)}",
        f"DTEND:{_fmt(end)}",
    ]
    if location:
        lines.append(f"LOCATION:{location}")
    if organizer_email:
        lines.append(f"ORGANIZER:mailto:{organizer_email}")
    if attendee_email:
        lines.append(f"ATTENDEE:mailto:{attendee_email}")
    lines += ["END:VEVENT", "END:VCALENDAR"]

    ics_text = "\r\n".join(lines) + "\r\n"

    if out_path is None:
        import os
        import tempfile
        fd, out_path = tempfile.mkstemp(suffix=".ics", prefix="invitacion_")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(ics_text)
    else:
        out_path = Path(out_path)
        out_path.write_text(ics_text, encoding="utf-8")

    return Path(out_path)
