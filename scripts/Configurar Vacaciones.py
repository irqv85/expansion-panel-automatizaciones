"""
Ventana simple para marcar que vendedores estan de vacaciones. A quien se
marque aqui no se le enviara nada (ni por turno de rotacion ni aunque
aparezca como asistente en una reunion de Sofia) hasta que se le quite la
marca. Usa el mismo archivo vacaciones.json que "Sofia - Schedule.py".

Ejecutar con: python "Configurar Vacaciones.py"  (o el .bat correspondiente)
"""
import importlib.util
import tkinter as tk
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
SOFIA_SCRIPT = SCRIPT_DIR / "Sofia - Schedule.py"

_spec = importlib.util.spec_from_file_location("sofia_schedule", SOFIA_SCRIPT)
sofia = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(sofia)


def main():
    root = tk.Tk()
    root.title("Vacaciones - Vendedores")
    root.resizable(False, False)

    tk.Label(
        root,
        text="Marca quien esta de vacaciones.\nNo recibira ningun envio automatico mientras este marcado.",
        justify="left", padx=15, pady=10,
    ).pack()

    current_vacation = sofia.load_vacation_emails()
    vars_by_email = {}

    frame = tk.Frame(root, padx=15)
    frame.pack(anchor="w")
    for seller in sofia.SELLERS:
        var = tk.BooleanVar(value=seller["email"].lower() in current_vacation)
        vars_by_email[seller["email"].lower()] = var
        tk.Checkbutton(frame, text=seller["name"], variable=var, anchor="w").pack(fill="x")

    status_label = tk.Label(root, text="", fg="#2e7d32", pady=5)
    status_label.pack()

    def guardar():
        emails = {email for email, var in vars_by_email.items() if var.get()}
        sofia.save_vacation_emails(emails)
        if emails:
            nombres = [s["name"] for s in sofia.SELLERS if s["email"].lower() in emails]
            status_label.config(text="Guardado. De vacaciones: " + ", ".join(nombres))
        else:
            status_label.config(text="Guardado. Nadie esta de vacaciones.")

    btn_frame = tk.Frame(root, pady=12)
    btn_frame.pack()
    tk.Button(btn_frame, text="Guardar", width=12, command=guardar).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Cerrar", width=12, command=root.destroy).pack(side="left", padx=5)

    root.mainloop()


if __name__ == "__main__":
    main()
