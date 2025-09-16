import pandas as pd
import pyodbc
import os

# ---------------------------
# PARAMÈTRES
# ---------------------------
dossier_excel = r"D:\Asa vacance\RETOUCHE IND 2025"

# Connexion SQL Server
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=informatik8;"
    "DATABASE=Requete_prime;"
    "Trusted_Connection=yes;"
)
cursor = conn.cursor()

# ---------------------------
# CRÉATION DE LA TABLE SI ABSENTE
# ---------------------------
cursor.execute("""
IF OBJECT_ID('dbo.retouche', 'U') IS NULL
CREATE TABLE dbo.retouche (
    CH NVARCHAR(255),
    MATRICULE NVARCHAR(255),
    RETOUCHE NVARCHAR(255),
    NomFichier NVARCHAR(255),
    NomFeuille NVARCHAR(255)
)
""")
conn.commit()

# ---------------------------
# PARAMÈTRES D'IMPORT
# ---------------------------
table = "dbo.retouche"
colonnes_a_importer = ["CH", " MATRICULE", "RETOUCHE"]  # espace conservé

# ---------------------------
# LECTURE DE TOUS LES FICHIERS DU RÉPERTOIRE
# ---------------------------
for fichier in os.listdir(dossier_excel):
    if fichier.endswith((".xlsx", ".xls")) and fichier.upper().startswith("RETOUCHE IND"):
        chemin_fichier = os.path.join(dossier_excel, fichier)
        xls = pd.ExcelFile(chemin_fichier)

        for nom_feuille in xls.sheet_names:
            if nom_feuille.upper().startswith("RETOUCHE"):
                try:
                    # Vérifier le nombre de lignes
                    df_preview = pd.read_excel(chemin_fichier, sheet_name=nom_feuille, header=None)
                    if len(df_preview) < 5:
                        continue

                    # Lire à partir de la 5ème ligne (header=4)
                    df = pd.read_excel(chemin_fichier, sheet_name=nom_feuille, header=4)

                    # Normaliser les noms de colonnes pour éviter les erreurs
                    df.columns = [col if col == " MATRICULE" else str(col).strip() for col in df.columns]

                    # Vérifier que toutes les colonnes attendues existent
                    if not all(col in df.columns for col in colonnes_a_importer):
                        print(f"⚠️ Colonnes manquantes dans {fichier} - {nom_feuille}")
                        continue

                    # Sélectionner les colonnes exactes à importer
                    df = df[colonnes_a_importer].fillna("")

                    # Colonnes SQL + placeholders
                    colonnes_sql = ", ".join([c.strip() for c in colonnes_a_importer] + ["NomFichier", "NomFeuille"])
                    placeholders = ", ".join(["?"] * (len(colonnes_a_importer) + 2))
                    valeurs = [row.tolist() + [fichier, nom_feuille] for _, row in df.iterrows()]

                    # Insertion en masse
                    cursor.executemany(
                        f"INSERT INTO {table} ({colonnes_sql}) VALUES ({placeholders})",
                        valeurs
                    )
                    conn.commit()

                    print(f"✅ Importé : {fichier} - {nom_feuille} → {table} ({len(df)} lignes)")

                except Exception as e:
                    print(f"⚠️ Erreur avec {fichier} - {nom_feuille} : {e}")

cursor.close()
conn.close()
print("🎯 Import terminé dans la table dbo.retouche.")