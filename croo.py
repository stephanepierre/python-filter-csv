import pandas as pd

# Définition du chemin du fichier CSV
fichier_csv = r'C:\Users\steph\Downloads\portal-reports.csv'

# Définir le chemin de sortie pour le fichier Excel
fichier_excel_origine = r'C:\Users\steph\Downloads\portal-reports.xlsx'

# Définition des chemins des fichiers de sortie
fichier_ventes = r'C:\Users\steph\Downloads\appels-VENTE-manqués.xlsx'
fichier_sav = r'C:\Users\steph\Downloads\appels-SAV-manqués.xlsx'
fichier_stats_ventes = r'C:\Users\steph\Downloads\stats-ventes.xlsx'

# Liste des noms à exclure
noms_exclus = ["David Territo", "Boualem Djebara", "Franca Territo", "Marketing"]

# Fonction pour convertir les secondes en heures et minutes
def convert_seconds_to_hours_minutes(seconds):
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    return f"{hours}h{minutes}m"

# Chargement du fichier CSV
df = pd.read_csv(fichier_csv)

# Suppression des colonnes 'call_start_time' et 'call_id' avant la sauvegarde
df_modifie = df.drop(columns=['call_start_time', 'call_id'])

# Sauvegarde du DataFrame modifié en fichier Excel
df_modifie.to_excel(fichier_excel_origine, index=False)

# Filtrage initial des données
df_filtre = df[(df["action"].isin(["missed", "voicemail"])) &
               (df["direction"] != "internal") &
               (~df["extension_name"].isin(noms_exclus))]

# Identifier les appels manqués qui n'ont pas été rappelés
df_manques_non_rappeles = df_filtre.copy()
df_manques_non_rappeles['rappel'] = df_manques_non_rappeles.apply(
    lambda row: df[(df['from_number'] == row['to_number']) &
                   (df['direction'] == 'out') &
                   (pd.to_datetime(df['action_time']) > pd.to_datetime(row['action_time']))].empty,
    axis=1
)

# Séparation des données VENTE et SAV pour les appels non rappelés
df_vente = df_manques_non_rappeles[(df_manques_non_rappeles["extension_name"].str.contains("Sales", case=False, na=False)) & (df_manques_non_rappeles['rappel'])]
df_sav = df_manques_non_rappeles[~(df_manques_non_rappeles["extension_name"].str.contains("Sales", case=False, na=False)) & (df_manques_non_rappeles['rappel'])]

# Affichage du nombre d'appels manqués pour les catégories VENTE et SAV
nombre_appels_manques_vente = len(df_vente)
nombre_appels_manques_sav = len(df_sav)

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
print(f"Nombre d'appels manqués VENTE: {nombre_appels_manques_vente}")
print(f"Nombre d'appels manqués SAV: {nombre_appels_manques_sav}")
