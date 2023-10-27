import pandas as pd

# Charger le fichier CSV dans un dataframe
df = pd.read_csv(r'C:\Users\steph\Downloads\portal-reports.csv')


# Filtrer les appels selon vos conditions
filtered_calls = df[
    (df["action"].isin(["missed", "voicemail"])) &
    (df["direction"] != "internal") &
    (~df["extension_name"].isin(["David Territo", "Boualem Djebara", "Franca Territo", "Marketing"]))
]

# Filtrer les appels rappelés
def was_called_back(row):
    calls_back = df[
        (df["from_number"] == row["from_number"]) &
        (df["call_start_time"] > row["call_start_time"]) &
        (df["action"] == "hanged") &
        (df["duration"] > 50)
    ]
    return not calls_back.empty


# Appliquer la fonction pour obtenir une liste des appels qui n'ont pas été rappelés
not_called_back = filtered_calls[~filtered_calls.apply(was_called_back, axis=1)]

# Séparation des appels manqués en fonction de la colonne 'extension_name'
sales_calls = not_called_back[not_called_back["extension_name"].str.contains("Sales", case=False, na=False)]
other_calls = not_called_back[~not_called_back["extension_name"].str.contains("Sales", case=False, na=False)]

# Enregistrer dans des fichiers Excel distincts
sales_calls.to_excel('appels-vente-manqués.xlsx', index=False)
other_calls.to_excel('sav-appels-manqués.xlsx', index=False)







