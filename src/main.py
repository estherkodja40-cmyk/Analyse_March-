import csv
import json
import os
import openpyxl 
import re
from difflib import SequenceMatcher

DOSSIER_DONNEES = "donnees"
# Chemin du dossier principal du projet
DOSSIER_PROJET = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Chemin du dossier contenant les données
DOSSIER_DONNEES = os.path.join(DOSSIER_PROJET, "donnees")
print("Dossier des données :", DOSSIER_DONNEES)

if os.path.exists(DOSSIER_DONNEES):
    print("Le dossier des données a été trouvé !")
else:
    print("Erreur : dossier des données introuvable.")
    # Récupération de la liste des fichiers
fichiers = os.listdir(DOSSIER_DONNEES)

print("\nFichiers disponibles :")

for fichier in fichiers:
    print(fichier)

def charger_fichier(chemin_fichier):

    # Récupérer l'extension du fichier
    extension = os.path.splitext(chemin_fichier)[1].lower()

    # ==================================
    # 1. LECTURE DES FICHIERS CSV
    # ==================================

    if extension == ".csv":

        with open(chemin_fichier, "r", encoding="utf-8-sig", newline="") as fichier:

            # Lire un échantillon pour détecter le séparateur
            echantillon = fichier.read(2048)
            fichier.seek(0)

            dialecte = csv.Sniffer().sniff(
                echantillon,
                delimiters=",;"
            )

            lecteur = csv.DictReader(fichier, dialect=dialecte)

            donnees = list(lecteur)

        return donnees

    # ==================================
    # 2. LECTURE DES FICHIERS JSON
    # ==================================

    elif extension == ".json":

        with open(chemin_fichier, "r", encoding="utf-8") as fichier:

            donnees = json.load(fichier)

        return donnees

    # ==================================
    # 3. LECTURE DES FICHIERS EXCEL
    # ==================================

    elif extension == ".xlsx":

        classeur = openpyxl.load_workbook(
            chemin_fichier,
            read_only=True,
            data_only=True
        )

        feuille = classeur.active

        # Récupérer toutes les lignes
        lignes = list(feuille.iter_rows(values_only=True))

        # Les véritables en-têtes se trouvent à la 3e ligne
        entetes = lignes[2]

        donnees = []

        # Les produits commencent à la 4e ligne
        for ligne in lignes[3:]:

            # Ignorer les lignes complètement vides
            if all(valeur is None for valeur in ligne):
                continue

            # Transformer chaque ligne en dictionnaire
            produit = dict(zip(entetes, ligne))

            donnees.append(produit)

        classeur.close()

        return donnees

    # ==================================
    # 4. FORMAT NON PRIS EN CHARGE
    # ==================================

    else:

        print("Format de fichier non pris en charge :", extension)

        return None
    
    


def normaliser_nom(nom):

    if nom is None:
        return ""

    nom = str(nom)

    nom = nom.lower()
    nom = nom.strip()
    nom = " ".join(nom.split())

    return nom
print("\n===== TEST DE NORMALISATION =====")

nom_test = "  Riz PARFUMÉ   5 KG  "

nom_normalise = normaliser_nom(nom_test)

print("Nom original :", nom_test)
print("Nom normalisé :", nom_normalise)
print("\n===== APERÇU DES DONNÉES DES BOUTIQUES =====")

fichiers_boutiques = [
    "marche_lagune.csv",
    "baobab_shop.json",
    "comptoir_ivoire.xlsx"
]

for fichier in fichiers_boutiques:

    chemin = os.path.join(DOSSIER_DONNEES, fichier)

    donnees = charger_fichier(chemin)

    print(f"\n--- {fichier} ---")

    print("Type de données :", type(donnees))

    if isinstance(donnees, list):
        print("Nombre de produits :", len(donnees))
        print("Premiers produits :", donnees[:2])

    else:
        print("Contenu :", donnees)

        print("\n========== STRUCTURE DES DONNÉES ==========")

for fichier in fichiers_boutiques:

    chemin = os.path.join(DOSSIER_DONNEES, fichier)

    donnees = charger_fichier(chemin)

    print("\nBoutique :", fichier)

    if isinstance(donnees, list):

        print("Nombre d'enregistrements :", len(donnees))

        if len(donnees) > 0:
            print("Colonnes disponibles :", donnees[0].keys())
            print("Premier enregistrement :", donnees[0])

    elif isinstance(donnees, dict):

        print("Clés disponibles :", donnees.keys())
        print("Contenu :", donnees)

        chemin_excel = os.path.join(
    DOSSIER_DONNEES,
    "comptoir_ivoire.xlsx"
)

classeur = openpyxl.load_workbook(
    chemin_excel,
    read_only=True,
    data_only=True
)

feuille = classeur.active

# Récupérer les lignes du fichier Excel
# Récupérer toutes les lignes du fichier
lignes = list(feuille.iter_rows(values_only=True))

# Récupérer les en-têtes situés à la 3e ligne
entetes = lignes[2]

donnees_excel = []

# Parcourir les produits à partir de la 4e ligne
for ligne in lignes[3:]:

    if all(valeur is None for valeur in ligne):
        continue

    produit = dict(zip(entetes, ligne))
    donnees_excel.append(produit)

print("\nNombre de produits :", len(donnees_excel))
print("Premier produit :", donnees_excel[0])

classeur.close()

# Récupérer les véritables en-têtes (3e ligne)
entetes = lignes[2]

donnees_excel = []

# Parcourir les produits à partir de la 4e ligne
for ligne in lignes[3:]:

    if all(valeur is None for valeur in ligne):
        continue

    produit = dict(zip(entetes, ligne))
    donnees_excel.append(produit)

classeur.close()

print("\nProduits récupérés :", len(donnees_excel))
print("Premier produit :", donnees_excel[0])
#pour le fichier csv
chemin_csv = os.path.join(DOSSIER_DONNEES, "marche_lagune.csv")

with open(chemin_csv, "r", encoding="utf-8-sig", newline="") as fichier:
    lecteur = csv.DictReader(fichier, delimiter=";")

    print("\n===== MARCHE LAGUNE =====")
    print("Colonnes :", lecteur.fieldnames)
    print("Premier produit :", next(lecteur, None))

#pour le json
chemin_json = os.path.join(DOSSIER_DONNEES, "baobab_shop.json")

with open(chemin_json, "r", encoding="utf-8") as fichier:
    donnees_json = json.load(fichier)

print("\n===== BAOBAB SHOP =====")
print("Type de données :", type(donnees_json))

if isinstance(donnees_json, list):
    print("Nombre de produits :", len(donnees_json))
    print("Premier produit :", donnees_json[0])

elif isinstance(donnees_json, dict):
    print("Clés disponibles :", donnees_json.keys())
    print("Contenu :", donnees_json)

#l'exploitation de catalogue_reference.csv.
chemin_catalogue = os.path.join(
    DOSSIER_DONNEES,
    "catalogue_reference.csv"
)

donnees_catalogue = charger_fichier(chemin_catalogue)

print("\n===== CATALOGUE DE RÉFÉRENCE =====")
print("Nombre de produits :", len(donnees_catalogue))
print("Premier produit :", donnees_catalogue[0])

#créer un index de référence
# Création de l'index des produits
index_catalogue = {}

for produit in donnees_catalogue:
    identifiant = produit["produit_id"]
    index_catalogue[identifiant] = produit

print("\n===== INDEX DU CATALOGUE =====")
print(index_catalogue["P01"])

#préparer un deuxième index
# Index des noms de produits normalisés
index_noms = {}

for produit in donnees_catalogue:
    nom_normalise = normaliser_nom(produit["nom"])
    identifiant = produit["produit_id"]

    index_noms[nom_normalise] = identifiant

print("\n===== INDEX DES NOMS =====")
print(index_noms.get("samsung galaxy a15 128 go"))

#identifier les champs de chaque boutique
print("\n===== STRUCTURE DES BOUTIQUES =====")

for fichier in fichiers_boutiques:
    chemin = os.path.join(DOSSIER_DONNEES, fichier)
    donnees = charger_fichier(chemin)

    print(f"\n{fichier}")

    if isinstance(donnees, list) and len(donnees) > 0:
        print("Champs :", donnees[0].keys())
    elif isinstance(donnees, dict):
        print("Clés :", donnees.keys())

#récupérer les produits des trois boutiques
# Charger les données des trois boutiques

donnees_lagune = charger_fichier(
    os.path.join(DOSSIER_DONNEES, "marche_lagune.csv")
)

donnees_baobab = charger_fichier(
    os.path.join(DOSSIER_DONNEES, "baobab_shop.json")
)

donnees_comptoir = charger_fichier(
    os.path.join(DOSSIER_DONNEES, "comptoir_ivoire.xlsx")
)


# Récupérer les produits de Baobab Shop
produits_baobab = donnees_baobab["items"]


print("\n===== RÉCAPITULATIF DES BOUTIQUES =====")

print("Marché Lagune :", len(donnees_lagune), "produits")
print("Baobab Shop :", len(produits_baobab), "produits")
print("Comptoir Ivoire :", len(donnees_comptoir), "produits")


print("\n===== PREMIER PRODUIT DE CHAQUE BOUTIQUE =====")

print("Marché Lagune :", donnees_lagune[0])
print("Baobab Shop :", produits_baobab[0])
print("Comptoir Ivoire :", donnees_comptoir[0])
print("\n===== CATALOGUE DE RÉFÉRENCE =====")
for produit in donnees_catalogue:
    print(produit)

print("\n===== NOMS DES PRODUITS DU CATALOGUE =====")

for produit in donnees_catalogue:
    print(produit["produit_id"], ":", produit["nom"])

def preparer_nom_comparaison(nom):

    if nom is None:
        return ""

    nom = str(nom).lower()

    # Uniformiser les unités de stockage
    nom = nom.replace("gb", "go")

    # Supprimer la ponctuation
    nom = re.sub(r"[^a-z0-9à-ÿ]+", " ", nom)

    # Séparer les mots
    mots = nom.split()

    # Trier les mots par ordre alphabétique
    mots.sort()

    return " ".join(mots)

print("\n===== TEST DE NORMALISATION AVANCÉE =====")

print(preparer_nom_comparaison("SAMSUNG GALAXY A15 (128GB)"))
print(preparer_nom_comparaison("Galaxy A15 128GB - Samsung"))
print(preparer_nom_comparaison("Samsun Galaxy A15 128Go"))