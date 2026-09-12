import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from app.sto.p305_2026 import P3052026Profile
from app.document.checker import check_document
from app.document.formatting import format_document


def run() -> None:
    root = tk.Tk()
    root.title("СТО Formatter — P_305.2026")
    root.geometry("900x650")
    profile = P3052026Profile()
    state = {"path": None}

    tk.Label(root, text="СТО Formatter", font=("Arial", 24, "bold")).pack(pady=(18, 4))
    tk.Label(root, text="Проверка и оформление DOCX по профилю СТО P_305.2026").pack(pady=(0, 14))

    top = tk.Frame(root)
    top.pack(fill="x", padx=24)
    path_var = tk.StringVar(value="Файл не выбран")
    tk.Label(top, textvariable=path_var, anchor="w").pack(side="left", fill="x", expand=True)

    def choose() -> None:
        path = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])
        if path:
            state["path"] = path
            path_var.set(path)
            output.delete("1.0", "end")

    tk.Button(top, text="Выбрать DOCX", command=choose, width=18).pack(side="right")

    output = tk.Text(root, wrap="word", font=("Consolas", 10))
    output.pack(fill="both", expand=True, padx=24, pady=18)

    def show(findings) -> None:
        output.delete("1.0", "end")
        for finding in findings:
            mark = "OK" if finding.severity.value == "ok" else "!" if finding.severity.value == "warning" else "ERROR"
            fixed = " [ИСПРАВЛЕНО]" if finding.fixed else ""
            location = f" ({finding.location})" if finding.location else ""
            output.insert("end", f"[{mark}] {finding.rule_id}{location}: {finding.message}{fixed}\n")

    def require_file() -> str | None:
        if not state["path"]:
            messagebox.showwarning("СТО", "Сначала выберите DOCX.")
            return None
        return state["path"]

    def check() -> None:
        path = require_file()
        if path:
            show(check_document(path, profile.config, profile.structural_elements))

    def format_and_save() -> None:
        path = require_file()
        if not path:
            return
        source = Path(path)
        target = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=f"{source.stem}_STO_P305_2026.docx",
            filetypes=[("Word documents", "*.docx")],
        )
        if not target:
            return
        try:
            findings = format_document(path, target, profile.config, profile.structural_elements)
            show(findings)
            messagebox.showinfo("Готово", f"Создан новый файл:\n{target}\n\nИсходный файл не изменён.")
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    buttons = tk.Frame(root)
    buttons.pack(pady=(0, 18))
    tk.Button(buttons, text="Проверить", command=check, width=18).pack(side="left", padx=8)
    tk.Button(buttons, text="Оформить и сохранить", command=format_and_save, width=24).pack(side="left", padx=8)
    root.mainloop()
