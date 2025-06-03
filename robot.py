import os
import shutil
import pandas as pd
import schedule
import time
import logging
import ast

# Assume fonctions.py is in the same directory
from fonctions import split_section_analytique, df_journaux, ajouter_type_piece5, completer_date_echeance,merge_custom, merge_Tiers2, merge_Tiers3, merge_journaux, merge_IFRS
from fonctions import retriev_code_PNL_from_df, renommer_colonnes, preparation2, convertir_sage100_en_x31, qualifier_et_controler3, pipeline, nettoyer_dataframe, transformer_ecritures_stock
from fonctions import extraire_infos_depuis_nom, charger_journal_pour_societe, df_journaux


#Données statiques
df_tiers = pd.read_excel("datas\MappingTousLesTiers.xlsx", sheet_name="ALL++")
df_IFRS = pd.read_excel("datas/Mapping IFRS LCR.xlsx")
pnl = pd.read_excel("datas/Mapping_PNL_Vrai.xlsb", engine='pyxlsb')  # <-- important

# ---------------------- CONFIGURATION ----------------------
#DOSSIER_SOURCE = r"\\172.16.10.75\rpa\SAGE_100"
DOSSIER_SOURCE = "datas\data_ecritures\TestRobot"
#DOSSIER_DEPOT = r"\\172.16.10.75\rpa\SAGE_X3"
DOSSIER_DEPOT = "datas\Enregistrements\TestRobot"

# ---------------------- LOGGING ----------------------
logging.basicConfig(filename="log_processus_local.log", level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# ---------------------- PLACEHOLDER FONCTIONS ----------------------


# ---------------------- TRAITEMENT MULTIPLE ----------------------
def traiter_tous_les_fichiers_excel(dossier, df_tiers, df_IFRS, pnl, dossier_sortie="datas/Enregistrements"):
    dossier_traite = os.path.join(dossier, "FICHIERS_TRAITES")
    os.makedirs(dossier_traite, exist_ok=True)

    if not os.path.exists(dossier):
        logging.error(f"Le chemin source est introuvable : {dossier}")
        return

    fichiers_excel = [f for f in os.listdir(dossier) if f.endswith('.xlsx') and os.path.isfile(os.path.join(dossier, f))]
    if not fichiers_excel:
        logging.info("Aucun fichier Excel trouvé dans le dossier Sage 100.")
        return
    else:
        logging.info(f"Il y a {len(fichiers_excel)} fichier(s) Excel dans le dossier Sage 100")

    versions_tracker = {}

    for fichier in fichiers_excel:
        chemin_fichier_source = os.path.join(dossier, fichier)
        logging.info(f"Début du traitement du fichier : {fichier}")

        try:
            societe, mois = extraire_infos_depuis_nom(fichier)
            if not societe:
                logging.warning(f"Société non trouvée dans le nom du fichier {fichier}")
                continue

            cle_version = f"{societe}_{mois}"
            versions_tracker[cle_version] = versions_tracker.get(cle_version, 0) + 1
            version = f"F{versions_tracker[cle_version]}"

            journal = charger_journal_pour_societe(societe)

            df_prep, df_prepAN, df_x3, df_filtered, df_tiers_filtered, df_excluded, df_tiers_filtered, df_qualite, fichier_sortie = pipeline(
                chemin_fichier_source,
                df_tiers,
                journal,
                df_IFRS,
                pnl,
                societe,
                mois,
                version,
                DS=dossier_sortie
            )

            logging.info(f"Transformation au format X3 du fichier {chemin_fichier_source} terminée !")
            logging.info(f"Fichier transformé et déposé ici : {fichier_sortie}")

            # ✅ Déplacement du fichier traité
            chemin_fichier_traite = os.path.join(dossier_traite, fichier)
            shutil.move(chemin_fichier_source, chemin_fichier_traite)
            logging.info(f"Fichier déplacé vers : {chemin_fichier_traite}")

        except Exception as e:
            logging.error(f"Erreur lors du traitement de {fichier}: {e}")


# ---------------------- JOB PLANIFIE ----------------------
def job_complet_local():
    logging.info("Début du traitement de tous les fichiers Excel")
    traiter_tous_les_fichiers_excel(DOSSIER_SOURCE, df_tiers,df_IFRS, pnl, DOSSIER_DEPOT)
    logging.info("Fin du traitement global")
    #compare_akanea_sage100(factures, TEXT)

# ---------------------- PLANIFICATION ----------------------
schedule.every(60).seconds.do(job_complet_local)  # Pour test rapide

logging.info("Lancement du planificateur local...")

while True:
    schedule.run_pending()
    time.sleep(1)


