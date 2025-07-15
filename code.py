import os
import re
import time
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32print
import subprocess

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def list_printers():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    return [printer[2] for printer in win32print.EnumPrinters(flags)]

def wait_for_print_job_to_finish(printer_name):
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        while True:
            jobs = win32print.EnumJobs(hPrinter, 0, -1, 1)
            if not jobs:
                break
            time.sleep(1)
    finally:
        win32print.ClosePrinter(hPrinter)

def print_pdf(filepath, printer_name):
    try:
        sumatra_path = resource_path("SumatraPDF.exe")
        if not os.path.exists(sumatra_path):
            raise FileNotFoundError("SumatraPDF.exe est introuvable dans le dossier du script.")
        subprocess.run([
            sumatra_path,
            "-print-to", printer_name,
            "-silent",
            filepath
        ], creationflags=subprocess.CREATE_NO_WINDOW)
        try:
            wait_for_print_job_to_finish(printer_name)
        except Exception:
            pass
    except Exception as e:
        messagebox.showerror("Erreur impression", f"Erreur lors de l'impression de {filepath} avec SumatraPDF :\n{e}")

def extract_number(filename):
    match = re.search(r"(\d+)", filename)
    return int(match.group(1)) if match else float("inf")

def process_folder(folder_path, printer_name, copies):
    try:
        folders_to_process = [
            "01_DUAO", "02_Manuel(s) de Montage", "03_Fiche(s) Technique", "04_Certificat(s)",
            "07_Rapport(s) de Réception", "08_Manuel(s) d'Utilisation EPI",
            "09_Certificat(s) Monteur(s)", "10_Garantie Décennale"
        ]
        for _ in range(copies):
            for folder in folders_to_process:
                current_path = os.path.join(folder_path, folder)
                if not os.path.exists(current_path):
                    continue
                files = sorted(
                    [f for f in os.listdir(current_path) if f.lower().endswith(".pdf")],
                    key=extract_number
                )
                if folder == "01_DUAO":
                    found_duao = False
                    for file in files:
                        if "DUAO" in file.upper():
                            full_path = os.path.join(current_path, file)
                            print_pdf(full_path, printer_name)
                            found_duao = True
                            break
                    if not found_duao:
                        messagebox.showerror("Fichier DUAO manquant", "Aucun fichier contenant 'DUAO' trouvé dans 01_DUAO. Opération stoppée.")
                        return
                else:
                    for file in files:
                        full_path = os.path.join(current_path, file)
                        print_pdf(full_path, printer_name)

        if copies == 1:
            messagebox.showinfo("Impression terminée", "1 copie imprimée avec succès.")
        else:
            messagebox.showinfo("Impression terminée", f"{copies} copies imprimées avec succès.")
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue :\n{e}")

def open_printer_preferences(printer_name):
    try:
        if not printer_name:
            messagebox.showwarning("Imprimante manquante", "Merci de sélectionner une imprimante d'abord.")
            return
        subprocess.Popen(f'rundll32 printui.dll,PrintUIEntry /p /n "{printer_name}"')
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible d’ouvrir les préférences : {e}")

def launch_gui():
    def select_folder():
        folder = filedialog.askdirectory()
        folder_var.set(folder)

    def run_process():
        folder = folder_var.get()
        printer = printer_var.get()
        try:
            copies = int(copies_var.get())
        except ValueError:
            messagebox.showwarning("Valeur invalide", "Le nombre de copies doit être un entier.")
            return
        if not folder or not printer:
            messagebox.showwarning("Champs manquants", "Merci de sélectionner un dossier et une imprimante.")
            return
        process_folder(folder, printer, copies)

    def show_credits():
        credits = tk.Toplevel(root)
        credits.title("Crédit")
        credits.geometry("300x100")
        tk.Label(credits, text="Logiciel d'impression As-Built", pady=5).pack()
        link = tk.Label(credits, text="Réalisé par Swax", fg="blue", cursor="hand2")
        link.pack()
        link.bind("<Button-1>", lambda e: subprocess.run(["start", "https://swax.cc/"], shell=True))

    def open_preferences():
        printer = printer_var.get()
        open_printer_preferences(printer)

    root = tk.Tk()
    root.title("As-Built Printer")
    root.iconbitmap(resource_path("iconeasbuiltprinter.ico"))
    root.geometry("500x420")

    folder_var = tk.StringVar()
    printer_var = tk.StringVar()
    copies_var = tk.StringVar(value="1")

    tk.Label(root, text="Sélectionne le dossier As-Built :", pady=5).pack()
    tk.Entry(root, textvariable=folder_var, width=60).pack()
    tk.Button(root, text="Parcourir...", command=select_folder).pack(pady=5)

    tk.Label(root, text="Sélectionne une imprimante :", pady=5).pack()
    printer_combo = ttk.Combobox(root, textvariable=printer_var, values=list_printers(), width=50)
    printer_combo.pack()
    printer_combo.set("Microsoft Print to PDF")

    options_frame = tk.Frame(root)
    options_frame.pack(pady=10)
    tk.Label(options_frame, text="Nombre de copies :").grid(row=0, column=0, padx=5)
    tk.Entry(options_frame, textvariable=copies_var, width=5).grid(row=0, column=1, padx=5)

    tk.Button(root, text="Imprimer", command=run_process, bg="green", fg="white").pack(pady=10)

    button_frame = tk.Frame(root)
    button_frame.pack()
    tk.Button(button_frame, text="Préférences d'impression", command=open_preferences).pack(side="left", padx=10)
    tk.Button(button_frame, text="Crédit", command=show_credits).pack(side="left", padx=10)

    root.mainloop()

launch_gui()
