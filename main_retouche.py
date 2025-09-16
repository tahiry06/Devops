import os
import pandas as pd
import pyodbc

#Paramètres
dossier_excel = r"D:\Asa vacance\RETOUCHE IND 2025" #chaine brute r""

# Période à analyser
date_debut = pd.to_datetime(input("👉 Entrez la date de début (jj/mm/aaaa) : "), format="%d/%m/%Y")
date_fin = pd.to_datetime(input("👉 Entrez la date de fin (jj/mm/aaaa) : "), format="%d/%m/%Y")

# Connexion SQL Server
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=informatik8;"
    "DATABASE=Requete_prime;"
    "Trusted_Connection=yes;" #login utiliser automatiquement, pas besoin de donner UID (utilisateur) ni PWD (mot de passe).
)    #SQL Server utilisera directement ton compte Windows actuel pour vérifier les droits d’accès.
cursor = conn.cursor()

#Création de la table si absente
cursor.execute("""
IF OBJECT_ID('dbo.ImportRetouche', 'U') IS NULL
CREATE TABLE dbo.ImportRetouche (
    CH NVARCHAR(50),
    MATRICULE NVARCHAR(50),
    SommeRetouche FLOAT,
    DateDebut DATE,
    DateFin DATE,
    NomFichier NVARCHAR(255)
)
""")
conn.commit()

#Trouver le dernier fichier
fichiers = [f for f in os.listdir(dossier_excel) if f.endswith((".xlsx", ".xls"))]
if not fichiers:
    print("❌ Aucun fichier trouvé")
    exit()

fichiers.sort(key=lambda x: os.path.getmtime(os.path.join(dossier_excel, x)))
dernier_fichier = fichiers[-1]
chemin_fichier = os.path.join(dossier_excel, dernier_fichier)
print(f"📌 Dernier fichier sélectionné : {dernier_fichier}")

#Lecture du fichier
df = pd.read_excel(chemin_fichier, sheet_name="RETOUCHE SEPT", header=4)  # ligne 5 = header=4

# Colonnes fixes
colonnes_a_importer = ["CH", "MATRICULE", "RETOUCHE"]

# Colonnes de dates = toutes les autres
colonnes_dates = [c for c in df.columns if c not in colonnes_a_importer]

# Debug
#print("📌 Colonnes détectées :", df.columns.tolist())
#print("📌 Colonnes dates détectées :", colonnes_dates)

# Transformation colonnes dates en format long
df_long = df.melt(
    id_vars=colonnes_a_importer,
    value_vars=colonnes_dates,
    var_name="Date",
    value_name="Valeur"
)

# Conversion des dates
df_long["Date"] = pd.to_datetime(df_long["Date"], format="%d/%m/%Y", errors="coerce")

# Filtrage par période
df_filtre = df_long[
    (df_long["Date"] >= date_debut) &
    (df_long["Date"] <= date_fin)
]

# Agrégation : somme des retouches par CH + MATRICULE
df_group = df_filtre.groupby(["CH", "MATRICULE"], as_index=False)["Valeur"].sum()
df_group.rename(columns={"Valeur": "SommeRetouche"}, inplace=True)

# Ajouter infos supplémentaires
df_group["DateDebut"] = date_debut
df_group["DateFin"] = date_fin
df_group["NomFichier"] = dernier_fichier

#Insertion SQL
if df_group.empty:
    print("⚠️ Aucun enregistrement trouvé → rien à insérer dans SQL.")
else:
    valeurs = df_group.where(pd.notnull(df_group), None).values.tolist()
    placeholders = ", ".join(["?"] * len(df_group.columns))
    colonnes_str = ", ".join([f"[{c}]" for c in df_group.columns])

    cursor.executemany(
        f"INSERT INTO dbo.ImportRetouche ({colonnes_str}) VALUES ({placeholders})",
        valeurs
    )
    conn.commit()
    print(f"✅ {len(df_group)} lignes insérées dans 'ImportRetouche'.")

cursor.close()
conn.close()
print("🎯 Import terminé.")
