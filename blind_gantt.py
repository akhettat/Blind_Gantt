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

# Charger les données des tâches et jalons depuis un fichier Excel
excel_file_path = 'taches_et_jalons.xlsx'  # Remplacer par le chemin du fichier Excel

try:
    # Lire les données depuis le fichier Excel
    # Utiliser le moteur openpyxl pour éviter les problèmes de dépendances
    df = pd.read_excel(excel_file_path, engine='openpyxl')
except FileNotFoundError as e:
    logging.error(f"Erreur : Le fichier {excel_file_path} est introuvable. {e}")
    exit(1)
except ValueError as e:
    logging.error(f"Erreur : Problème lors de la lecture du fichier Excel. Vérifiez le format du fichier. {e}")
    exit(1)

# Vérifier les colonnes nécessaires
required_columns = ['Task', 'Start', 'Duration', 'Predecessor']
for col in required_columns:
    if col not in df.columns:
        logging.error(f"Erreur : La colonne '{col}' est manquante dans le fichier Excel.")
        exit(1)

# Filtrer les lignes qui sont entièrement remplies et qui ont plus que seulement le numéro de tâche
df = df.dropna(how='any', subset=['Task', 'Start', 'Duration'])
df = df[df[['Task', 'Start', 'Duration']].notna().all(axis=1)]

# Supprimer les lignes où seules les colonnes 'Task ID' sont remplies
df = df[df[['Task', 'Start', 'Duration', 'Predecessor']].notna().any(axis=1)]

# Remplacer les valeurs de prédécesseur manquantes par '-'
df['Predecessor'] = df['Predecessor'].replace('-', None).fillna('-')

# Afficher le nombre de tâches et jalons valides trouvés
total_valid_tasks = len(df)
print(f"Nombre de tâches et jalons valides trouvés : {total_valid_tasks}")

# Paramètres
working_hours_per_day = 8
weekend_days = [5, 6]  # Samedi, Dimanche
current_year = datetime.now().year

# Essayer d'obtenir les jours fériés pour la France
try:
    # Obtenir le code de pays à partir de la locale du système
    locale.setlocale(locale.LC_TIME, '')
    system_country_code = locale.getlocale()[0].split('_')[1]
    try:
        french_holidays = holidays.country_holidays(system_country_code, years=[current_year])
    except Exception as e:
        logging.error(f"Erreur lors de l'obtention des jours fériés pour le pays '{system_country_code}' : {e}")
        french_holidays = []
except Exception as e:
    logging.error(f"Erreur lors de l'obtention de la locale du système : {e}")
    french_holidays = []

# Convertir les colonnes de dates en format datetime
try:
    df['Start'] = pd.to_datetime(df['Start'], dayfirst=True, errors='coerce')
except Exception as e:
    logging.error(f"Erreur lors de la conversion des dates : {e}")
    exit(1)

# Vérifier les valeurs manquantes ou incorrectes
if df['Start'].isna().any():
    logging.error("Erreur : Certaines dates de début ne sont pas valides dans le fichier Excel.")
    exit(1)

# Vérifier les durées
if df['Duration'].isna().any() or not pd.api.types.is_numeric_dtype(df['Duration']):
    logging.error("Erreur : Certaines durées sont manquantes ou non numériques dans le fichier Excel.")
    exit(1)

# Fonction pour calculer la date de fin en fonction des jours ouvrés
def calculate_end_date(start_date, duration_hours):
    remaining_hours = duration_hours
    current_date = pd.to_datetime(start_date)
    while remaining_hours > 0:
        if current_date.weekday() not in weekend_days and current_date not in french_holidays:
            remaining_hours -= working_hours_per_day
        current_date += timedelta(days=1)
    return current_date

# Calculer la date de fin prévue pour chaque tâche
df['End'] = df.apply(lambda row: calculate_end_date(row['Start'], row['Duration']), axis=1)

# Ajuster les dates de début et de fin en fonction des prédécesseurs
for index, row in df.iterrows():
    if row['Predecessor'] != '-':
        # Chercher le prédécesseur par son ID
        predecessor = df[df['Task'] == row['Predecessor']]
        if not predecessor.empty:
            predecessor_end = predecessor['End'].values[0]
            if row['Start'] < predecessor_end:
                df.at[index, 'Start'] = predecessor_end
                df.at[index, 'End'] = calculate_end_date(predecessor_end, row['Duration'])

# Mettre à jour les dates de début dans le fichier Excel
df.to_excel(excel_file_path, index=False)

# Trier les tâches et jalons en fonction de la date de début
milestones = df[df['Duration'] == 0].sort_values(by='Start')
tasks = df[df['Duration'] > 0].sort_values(by='Start')
df = pd.concat([milestones, tasks]).reset_index(drop=True)

# Définir la date d'aujourd'hui pour évaluer l'état des tâches
today = pd.to_datetime(datetime.today().strftime('%d/%m/%Y'), dayfirst=True)

# Ajouter une colonne pour indiquer l'état de la tâche
def task_status(row):
    if row['End'] < today:
        return 'Terminé' if row['Duration'] == 0 else 'En retard'
    elif row['Start'] <= today < row['End']:
        return 'Démarré'
    else:
        return 'Pas encore démarré'

df['Status'] = df.apply(task_status, axis=1)

# Création du diagramme de Gantt avec Matplotlib
fig, ax = plt.subplots(figsize=(12, 8))

# Définir des couleurs pour chaque état
status_colors = {
    'Terminé': 'blue',
    'En retard': 'red',
    'Démarré': 'orange',
    'Pas encore démarré': 'green'
}

# Tracer les tâches et jalons
y_position = range(len(df))
for index, row in df.iterrows():
    color = status_colors[row['Status']]
    y_index = len(df) - 1 - index  # Afficher de haut en bas
    if row['Duration'] == 0:  # Jalon
        ax.plot(row['Start'], y_index, marker='o', color=color, markersize=10)
        ax.text(row['Start'], y_index, f"{row['Task']}", va='center', ha='left', fontsize=16, color='black')
    else:  # Tâche
        ax.barh(y_index, (row['End'] - row['Start']).days, left=row['Start'], color=color, edgecolor='black')
        ax.text(row['Start'] + (row['End'] - row['Start']) / 2, y_index, f"{row['Task']}", va='center', ha='center', fontsize=16, color='white')

# Ajouter des lignes pour représenter les liens entre les tâches
for index, row in df.iterrows():
    if row['Predecessor'] != '-':
        predecessor = df[df['Task'] == row['Predecessor']]
        if not predecessor.empty:
            predecessor_index = len(df) - 1 - predecessor.index[0]
            y_index = len(df) - 1 - index
            ax.plot([predecessor['End'].values[0], row['Start']], [predecessor_index, y_index], 'k--', lw=1)

# Formatage des dates sur l'axe x
ax.xaxis.set_major_locator(mdates.WeekdayLocator())
ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%Y'))
plt.xticks(rotation=45)

# Supprimer les étiquettes de l'axe y
ax.set_yticks([])

# Ajouter une légende
legend_elements = [Patch(facecolor=color, edgecolor='black', label=status) for status, color in status_colors.items()]
ax.legend(handles=legend_elements, loc='upper right')

# Ajouter des labels
ax.set_xlabel('Dates')
ax.set_title('Diagramme de Gantt avec Prédécesseurs et Jours Ouvrés')

plt.tight_layout()

# Sauvegarder l'image dans le même répertoire que le fichier Excel
output_image_path = os.path.join(os.path.dirname(excel_file_path), 'gantt_chart.png')
plt.savefig(output_image_path)
print(f"Diagramme de Gantt sauvegardé sous : {output_image_path}")

plt.show()

# Si aucune exception n'a été détectée, enregistrer "Execution successful" dans le fichier log
if not logging.getLogger().hasHandlers():
    with open(log_file_path, 'a') as log_file:  
        log_file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - INFO - Execution successful\n")