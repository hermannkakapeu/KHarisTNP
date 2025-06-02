import pandas as pd
import re
from datetime import datetime

def est_date(val):
    try:
        datetime.strptime(val, "%d%m%y")
        return True
    except:
        return False

def est_montant(val):
    return bool(re.match(r"^-?\d+([.,]\d+)?$", val))

records = []
current_record = {}
bloc = []

with open("Akanea\\factures.Txt", "r", encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        if not line:
            continue
        if line.startswith("#MECG"):
            if bloc:
                records.append(current_record)
            bloc = []
            current_record = {"bloc_type": "MECG"}
        else:
            bloc.append(line)

            # Analyse dynamique de chaque ligne
            if line == "VTE" or line == "ACH":
                current_record["type_mouvement"] = line
            elif line.startswith("FAB"):
                if "id_1" not in current_record:
                    current_record["id_1"] = line
                else:
                    current_record["id_2"] = line
            elif est_date(line):
                if "date_debut" not in current_record:
                    current_record["date_debut"] = line
                elif "date_fin" not in current_record:
                    current_record["date_fin"] = line
                else:
                    current_record["date_facture"] = line
            elif line.isdigit() and len(line) == 8:
                if "compte1" not in current_record:
                    current_record["compte1"] = line
                elif "compte2" not in current_record:
                    current_record["compte2"] = line
                else:
                    current_record["compte3"] = line
            elif est_montant(line):
                current_record["montant"] = line.replace(',', '.')
            elif line.isdigit() and int(line) in [0, 1]:
                if "type_ligne" not in current_record:
                    current_record["type_ligne"] = line
                else:
                    current_record["inconnu"] = line
            elif len(line) > 20 or "FACT" in line:
                current_record["libelle"] = line

# Ajouter le dernier bloc
if current_record:
    records.append(current_record)

df = pd.DataFrame(records)

# Conversion de dates et montants
for col in ["date_debut", "date_fin", "date_facture"]:
    df[col] = pd.to_datetime(df[col], format="%d%m%y", errors="coerce")

df["montant"] = pd.to_numeric(df["montant"], errors="coerce")

df.to_excel("Akanea\\factures2.xlsx", index=False)
