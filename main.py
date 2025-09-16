import pandas as pd
import pyodbc

# --- Charger le fichier Excel ---
fichier_excel = r"D:\Asa vacance\EN 16 CTRL INDICATEUR QUALITE RETOUCHE.xlsx"
df = pd.read_excel(fichier_excel, sheet_name="RECAP_EPSILON", header=2)

# --- Garder les colonnes utiles ---
df = df[["DATE", "CHAINE", "Décision"]]

# --- Nettoyage des colonnes ---
df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce").dt.date
df["CHAINE"] = df["CHAINE"].astype(str).replace("nan", None)
df["Décision"] = df["Décision"].astype(str).replace("nan", None)

# --- Supprimer les lignes où CHAINE ou DATE sont vides ---
df = df[
    df["CHAINE"].notna() & (df["CHAINE"].astype(str).str.strip() != "") &
    df["DATE"].notna()
]

# --- Connexion SQL Server ---
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=informatik8;"
    "DATABASE=Requete_prime;"
    "Trusted_Connection=yes;"
)
cursor = conn.cursor()

# --- Création de la table principale si elle n'existe pas ---
cursor.execute("""
IF OBJECT_ID('dbo.qualitechaine', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.qualitechaine (
        [DATE] DATE,
        [CHAINE] VARCHAR(100),
        [Décision] VARCHAR(100)
    )
END
""")
conn.commit()

# --- Insertion des données dans la table principale ---
for _, row in df.iterrows():
    cursor.execute(
        "INSERT INTO dbo.qualitechaine ([DATE], [CHAINE], [Décision]) VALUES (?, ?, ?)",
        row["DATE"], row["CHAINE"], row["Décision"]
    )
conn.commit()

# --- Regroupement par CHAINE pour compter le nombre de Décision ---
summary = df.groupby("CHAINE")["Décision"].count().reset_index()
summary = summary.rename(columns={"Décision": "NbDécision"})

# --- Création de la table résumé global ---
cursor.execute("""
IF OBJECT_ID('dbo.qualitechaine_grouper', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.qualitechaine_grouper (
        [CHAINE] VARCHAR(100),
        [NbDécision] INT
    )
END
""")
conn.commit()

# --- Insertion du résumé global ---
for _, row in summary.iterrows():
    cursor.execute(
        "INSERT INTO dbo.qualitechaine_grouper ([CHAINE], [NbDécision]) VALUES (?, ?)",
        row["CHAINE"], row["NbDécision"]
    )
conn.commit()

# --- Regroupement par CHAINE et par MOIS ---
df["Mois"] = df["DATE"].apply(lambda x: x.strftime("%Y-%m"))  # format AAAA-MM
summary_month = df.groupby(["CHAINE", "Mois"])["Décision"].count().reset_index()
summary_month = summary_month.rename(columns={"Décision": "NbDécision"})

# --- Création de la table résumé par mois ---
cursor.execute("""
IF OBJECT_ID('dbo.qualitechaine_mois', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.qualitechaine_mois (
        [CHAINE] VARCHAR(100),
        [Mois] VARCHAR(7),
        [NbDécision] INT
    )
END
""")
conn.commit()

# --- Insertion du résumé par mois ---
for _, row in summary_month.iterrows():
    cursor.execute(
        "INSERT INTO dbo.qualitechaine_mois ([CHAINE], [Mois], [NbDécision]) VALUES (?, ?, ?)",
        row["CHAINE"], row["Mois"], row["NbDécision"]
    )
conn.commit()

# --- Fermeture de la connexion ---
cursor.close()
conn.close()

print("Import terminé ✅ Données complètes et résumés par CHAINE et par mois insérés")
