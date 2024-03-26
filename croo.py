import pandas as pd

# Définition du chemin du fichier CSV
fichier_csv = r'C:\Users\steph\Downloads\portal-reports.csv'

# Définition des chemins des fichiers de sortie
fichier_ventes = 'appels-VENTE-manqués.xlsx'
fichier_sav = 'appels-SAV-manqués.xlsx'
fichier_stats_ventes = 'stats-ventes.xlsx'

# Liste des noms à exclure
noms_exclus = ["David Territo", "Boualem Djebara", "Franca Territo", "Marketing"]

# Fonction pour convertir les secondes en heures et minutes
def convert_seconds_to_hours_minutes(seconds):
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    return f"{hours}h{minutes}m"

# Chargement du fichier CSV
df = pd.read_csv(fichier_csv)

# Filtrage initial des données
df_filtre = df[(df["action"].isin(["missed", "voicemail"])) &
               (df["direction"] != "internal") &
               (~df["extension_name"].isin(noms_exclus))]

# Séparation des données VENTE et SAV
df_vente = df_filtre[df_filtre["extension_name"].str.contains("Sales", case=False, na=False)]
df_sav = df_filtre[~df_filtre["extension_name"].str.contains("Sales", case=False, na=False)]

# Sauvegarde des données filtrées dans des fichiers Excel
df_vente.to_excel(fichier_ventes, index=False)
df_sav.to_excel(fichier_sav, index=False)

# Calcul des statistiques pour les postes "Sales"
stats = []

for poste in df['extension_name'].dropna().unique():
    if "Sales" in poste:
        poste_df = df[df['extension_name'] == poste]
        entrants_avec_reponse = poste_df[(poste_df['direction'] == 'in') & (poste_df['action'] == 'hanged')]
        sortants_avec_reponse = poste_df[(poste_df['direction'] == 'out') & (poste_df['action'] == 'hanged')]
        messages_recus = poste_df[(poste_df['direction'] == 'in') & (poste_df['action'] == 'voicemail')]
        
        stats.append({
            "Poste": poste,
            "Appels Entrants Avec Réponse": entrants_avec_reponse.shape[0],
            "Temps Total Appels Entrants": convert_seconds_to_hours_minutes(entrants_avec_reponse['duration'].sum()),
            "Appels Sortants Avec Réponse": sortants_avec_reponse.shape[0],
            "Temps Total Appels Sortants": convert_seconds_to_hours_minutes(sortants_avec_reponse['duration'].sum()),
            "Messages Reçus": messages_recus.shape[0],
            "Temps Total Messages": convert_seconds_to_hours_minutes(messages_recus['duration'].sum()),
        })

# Création du DataFrame pour les statistiques
df_stats = pd.DataFrame(stats)

# Sauvegarde du DataFrame dans un fichier Excel
df_stats.to_excel(fichier_stats_ventes, index=False)

print("Statistiques générées avec succès.")
