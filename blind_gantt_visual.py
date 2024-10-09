## title: BLIND GANTT Application ##  
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import Patch
from datetime import datetime, timedelta
import holidays
import os
import locale
import logging
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Déterminer le chemin de base en fonction de l'exécution (exécutable ou script)
if getattr(sys, 'frozen', False):
    # Si le script est exécuté en tant qu'exécutable PyInstaller
    base_path = sys._MEIPASS
else:
    # Si le script est exécuté normalement (fichier .py)
    base_path = os.path.dirname(__file__)

# Configurer le logging pour enregistrer les erreurs dans un fichier log.txt
log_file_path = os.path.join(os.getcwd(), 'log.txt')
with open(log_file_path, 'a') as log_file:
    log_file.write("\n\n" + "-" * 20 + f" Exécution le {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} " + "-" * 20 + "\n\n")

logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Créer l'application Tkinter
root = tk.Tk()
root.title("Application Diagramme de Gantt Accessible")
root.geometry("1200x700")
root.option_add('*Font', 'Arial 12')  # Définir une police plus grande par défaut
root.configure(bg='white')

# Créer un DataFrame vide
columns = ['Task', 'Start', 'Duration', 'Predecessor']
df = pd.DataFrame(columns=columns)

# Fonction pour calculer la date de fin en fonction des jours ouvrés
def calculate_end_date(start_date, duration_hours):
    working_hours_per_day = 8
    weekend_days = [5, 6]  # Samedi, Dimanche
    current_year = datetime.now().year
    try:
        french_holidays = holidays.country_holidays('FR', years=[current_year])
    except Exception as e:
        logging.error(f"Erreur lors de l'obtention des jours fériés : {e}")
        french_holidays = []

    remaining_hours = duration_hours
    current_date = pd.to_datetime(start_date)
    while remaining_hours > 0:
        if current_date.weekday() not in weekend_days and current_date not in french_holidays:
            remaining_hours -= working_hours_per_day
        current_date += timedelta(days=1)
    return current_date

# Fonction pour mettre à jour le diagramme de Gantt
def update_gantt_chart():
    global df
    try:
        # Mise à jour du DataFrame à partir des champs de saisie
        data = []
        for frame in task_frames:
            task = frame['task'].get()
            start = frame['start'].get()
            duration = frame['duration'].get()
            predecessor = frame['predecessor'].get()
            data.append([task, start, duration, predecessor])
        df = pd.DataFrame(data, columns=columns)

        # Convertir les colonnes appropriées
        df['Start'] = pd.to_datetime(df['Start'], dayfirst=True, errors='coerce')
        df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce')

        # Calculer la date de fin prévue pour chaque tâche
        df['End'] = df.apply(lambda row: calculate_end_date(row['Start'], row['Duration']), axis=1)

        # Ajuster les dates de début et de fin en fonction des prédécesseurs
        for index, row in df.iterrows():
            if row['Predecessor'] != '-':
                predecessor = df[df['Task'] == row['Predecessor']]
                if not predecessor.empty:
                    predecessor_end = predecessor['End'].values[0]
                    if row['Start'] < predecessor_end:
                        df.at[index, 'Start'] = predecessor_end
                        df.at[index, 'End'] = calculate_end_date(predecessor_end, row['Duration'])

        # Trier les tâches et jalons en fonction de la date de début
        milestones = df[df['Duration'] == 0].sort_values(by='Start')
        tasks = df[df['Duration'] > 0].sort_values(by='Start')
        df = pd.concat([milestones, tasks]).reset_index(drop=True)

        # Créer le diagramme de Gantt
        fig, ax = plt.subplots(figsize=(8, 5))
        status_colors = {
            'Terminé': 'blue',
            'En retard': 'red',
            'Démarré': 'orange',
            'Pas encore démarré': 'green'
        }
        y_position = range(len(df))
        for index, row in df.iterrows():
            color = status_colors.get(row.get('Status', 'Pas encore démarré'), 'green')
            y_index = len(df) - 1 - index  # Afficher de haut en bas
            if row['Duration'] == 0:  # Jalon
                ax.plot(row['Start'], y_index, marker='o', color=color, markersize=10)
                ax.text(row['Start'], y_index, f"{row['Task']}", va='center', ha='left', fontsize=14, color='black')
            else:  # Tâche
                ax.barh(y_index, (row['End'] - row['Start']).days, left=row['Start'], color=color, edgecolor='black')
                ax.text(row['Start'] + (row['End'] - row['Start']) / 2, y_index, f"{row['Task']}", va='center', ha='center', fontsize=14, color='white')

        ax.xaxis.set_major_locator(mdates.WeekdayLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%Y'))
        plt.xticks(rotation=45)
        ax.set_yticks([])
        legend_elements = [Patch(facecolor=color, edgecolor='black', label=status) for status, color in status_colors.items()]
        ax.legend(handles=legend_elements, loc='upper right')
        ax.set_xlabel('Dates')
        ax.set_title('Diagramme de Gantt avec Prédécesseurs et Jours Ouvrés')
        plt.tight_layout()

        # Afficher le diagramme dans la fenêtre Tkinter
        for widget in gantt_frame.winfo_children():
            widget.destroy()
        canvas = FigureCanvasTkAgg(fig, master=gantt_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la mise à jour du diagramme de Gantt : {e}")
        logging.error(f"Erreur lors de la mise à jour du diagramme de Gantt : {e}")

# Créer les widgets de l'application
datagrid_frame = tk.Frame(root, bg='white')
datagrid_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
gantt_frame = tk.Frame(root, bg='white')
gantt_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

# Ajouter des champs pour saisir les données des tâches et jalons
task_frames = []

# Fonction pour ajouter une tâche ou un jalon
def add_task():
    frame = tk.Frame(datagrid_frame, bg='white')
    frame.pack(fill=tk.X, pady=5)

    task_label = ttk.Label(frame, text="Tâche/Jalon:", background='white')
    task_label.pack(side=tk.LEFT, padx=5)
    task_entry = ttk.Entry(frame)
    task_entry.pack(side=tk.LEFT, padx=5)
    task_entry.focus_set()

    start_label = ttk.Label(frame, text="Date de début (jj/mm/aaaa):", background='white')
    start_label.pack(side=tk.LEFT, padx=5)
    start_entry = ttk.Entry(frame)
    start_entry.pack(side=tk.LEFT, padx=5)

    duration_label = ttk.Label(frame, text="Durée (heures):", background='white')
    duration_label.pack(side=tk.LEFT, padx=5)
    duration_entry = ttk.Entry(frame)
    duration_entry.pack(side=tk.LEFT, padx=5)

    predecessor_label = ttk.Label(frame, text="Prédécesseur:", background='white')
    predecessor_label.pack(side=tk.LEFT, padx=5)
    predecessor_entry = ttk.Entry(frame)
    predecessor_entry.pack(side=tk.LEFT, padx=5)

    task_frames.append({
        'frame': frame,
        'task': task_entry,
        'start': start_entry,
        'duration': duration_entry,
        'predecessor': predecessor_entry
    })

# Bouton pour ajouter une nouvelle tâche/jalon
add_task_button = ttk.Button(datagrid_frame, text="Ajouter une tâche/jalon", command=add_task)
add_task_button.pack(pady=10)
add_task_button.focus_set()  # Mettre le focus initial sur le bouton Ajouter une tâche

# Créer le menu
menu = tk.Menu(root)
root.config(menu=menu)

# Menu Fichier
file_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="Fichier", menu=file_menu)
file_menu.add_command(label="Enregistrer l'image", command=lambda: plt.savefig('gantt_chart.png'))
file_menu.add_command(label="Sauvegarder les données", command=lambda: df.to_excel('taches_et_jalons_sauvegarde.xlsx', index=False))
file_menu.add_separator()
file_menu.add_command(label="Quitter", command=root.quit)

# Bouton pour mettre à jour le diagramme de Gantt
update_button = ttk.Button(root, text="Mettre à jour le Diagramme de Gantt", command=update_gantt_chart)
update_button.pack(side=tk.BOTTOM, pady=10)

# Lancer l'application
root.mainloop()