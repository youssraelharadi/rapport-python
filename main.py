import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
import csv
import os
from datetime import datetime
from openpyxl import Workbook

# Chemins des fichiers CSV
CSV_FILE = "rapports.csv"
CSV_MENSUEL = "rapports_mensuels.csv"

# Créer les fichiers CSV s'ils n'existent pas
if not os.path.exists(CSV_FILE):
    with open(CSV_FILE, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Chemin", "Date"])

if not os.path.exists(CSV_MENSUEL):
    with open(CSV_MENSUEL, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Mois", "Chemin", "Date"])

# Ajouter un fichier ou un dossier
def ajouter_fichier():
    fichiers = filedialog.askopenfilenames(title="Sélectionner des fichiers")
    dossiers = filedialog.askdirectory(title="Sélectionner un dossier")

    if not fichiers and not dossiers:
        messagebox.showerror("Erreur", "Aucun fichier ou dossier sélectionné.")
        return

    for fichier in fichiers:
        enregistrer_fichier(fichier)

    if dossiers:
        enregistrer_fichier(dossiers)

    afficher_rapports()

# Enregistrer dans le fichier CSV principal
def enregistrer_fichier(chemin):
    with open(CSV_FILE, mode="a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow([chemin, datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

# Afficher les fichiers/dossiers dans le tableau principal
def afficher_rapports():
    for row in table.get_children():
        table.delete(row)

    with open(CSV_FILE, mode="r", encoding="utf-8") as file:
        reader = csv.reader(file)
        next(reader)
        for ligne in reader:
            table.insert("", tk.END, values=ligne)

# Ajouter un rapport mensuel
def ajouter_rapport_mensuel(mois, chemin):
    with open(CSV_MENSUEL, mode="a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow([mois, chemin, datetime.now().strftime('%Y-%m-%d %H:%M:%S')])

    afficher_rapports_mensuels()

# Afficher les rapports mensuels dans le tableau
def afficher_rapports_mensuels():
    for row in table_mensuel.get_children():
        table_mensuel.delete(row)

    with open(CSV_MENSUEL, mode="r", encoding="utf-8") as file:
        reader = csv.reader(file)
        next(reader)
        for ligne in reader:
            table_mensuel.insert("", tk.END, values=ligne)

# Générer un rapport mensuel
def generer_rapport_mensuel():
    mois = datetime.now().strftime('%Y-%m')
    fichiers = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("Fichiers CSV", "*.csv")],
        title="Enregistrer le rapport mensuel"
    )

    if not fichiers:
        return

    with open(CSV_FILE, mode="r", encoding="utf-8") as file:
        reader = csv.reader(file)
        lignes = list(reader)

    with open(fichiers, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerows(lignes)

    for ligne in lignes:
        ajouter_rapport_mensuel(mois, ligne[0])

    messagebox.showinfo("Succès", "Rapport mensuel généré et ajouté.")

# Générer un rapport annuel
def generer_rapport_annee():
    fichiers = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("Fichiers CSV", "*.csv")],
        title="Enregistrer le rapport annuel"
    )

    if not fichiers:
        return

    with open(CSV_MENSUEL, mode="r", encoding="utf-8") as file:
        reader = csv.reader(file)
        lignes = list(reader)

    with open(fichiers, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerows(lignes)

    messagebox.showinfo("Succès", "Rapport annuel généré avec succès.")

# Télécharger le rapport annuel en Excel
def telecharger_rapport_excel():
    fichiers = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Fichiers Excel", "*.xlsx")],
        title="Enregistrer le rapport annuel en Excel"
    )

    if not fichiers:
        return

    with open(CSV_MENSUEL, mode="r", encoding="utf-8") as file:
        reader = csv.reader(file)
        lignes = list(reader)

    wb = Workbook()
    ws = wb.active
    ws.title = "Rapports Année"

    if lignes:
        ws.append(lignes[0])

    for ligne in lignes[1:]:
        ws.append(ligne)

    try:
        wb.save(fichiers)
        messagebox.showinfo("Succès", "Rapport annuel exporté en Excel avec succès.")
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")

# Interface principale
app = tk.Tk()
app.title("Gestion des Fichiers et Dossiers")

frame_ajout = tk.Frame(app, padx=10, pady=10)
frame_ajout.pack(pady=10)

btn_ajouter = tk.Button(frame_ajout, text="Ajouter Fichier/Dossier", command=ajouter_fichier)
btn_ajouter.grid(row=0, column=1, pady=10, sticky="e")

frame_table = tk.Frame(app, padx=10, pady=10)
frame_table.pack()

table = ttk.Treeview(frame_table, columns=("Chemin", "Date"), show="headings", height=10)
table.heading("Chemin", text="Chemin")
table.heading("Date", text="Date d'ajout")
table.column("Chemin", width=400)
table.column("Date", width=150)
table.pack()

frame_table_mensuel = tk.Frame(app, padx=10, pady=10)
frame_table_mensuel.pack(pady=10)

table_mensuel = ttk.Treeview(frame_table_mensuel, columns=("Mois", "Chemin", "Date"), show="headings", height=10)
table_mensuel.heading("Mois", text="Mois")
table_mensuel.heading("Chemin", text="Chemin")
table_mensuel.heading("Date", text="Date d'ajout")
table_mensuel.column("Mois", width=100)
table_mensuel.column("Chemin", width=300)
table_mensuel.column("Date", width=150)
table_mensuel.pack()

btn_generer_mensuel = tk.Button(app, text="Générer Rapport Mensuel", command=generer_rapport_mensuel)
btn_generer_mensuel.pack(pady=10)

btn_generer_annee = tk.Button(app, text="Générer Rapport Annuel", command=generer_rapport_annee)
btn_generer_annee.pack(pady=10)

btn_telecharger_excel = tk.Button(app, text="Télécharger Rapport Annuel (Excel)", command=telecharger_rapport_excel)
btn_telecharger_excel.pack(pady=10)

afficher_rapports()
afficher_rapports_mensuels()

app.mainloop()