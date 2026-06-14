#!/usr/bin/env python3
"""
Script principal pour le traitement des SMS
- Concaténation des fichiers Brutes avec dédoublonnage
- Création de fichiers par MSISDN du TP
- Routing des SMS dans les onglets correspondants
"""

import os
import sys
import shutil
from datetime import datetime

# Vérification des dépendances
try:
    import pandas as pd
    from openpyxl import Workbook, load_workbook
except ImportError as e:
    print(f"ERREUR: Dépendance manquante - {e}")
    print("Installation requise: pip install pandas openpyxl")
    sys.exit(1)

# Configuration
BRUTES_DIR = "Brutes"
CONCAT_DIR = "Concat"
ARCHIVE_DIR = "Brutes_archive"
TP_FILE = "TP/TP_test.xlsx"


def ensure_directories():
    """Créer les répertoires nécessaires"""
    os.makedirs(BRUTES_DIR, exist_ok=True)
    os.makedirs(CONCAT_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_DIR, exist_ok=True)


def load_file(filepath):
    """Charger un fichier CSV ou Excel"""
    try:
        if filepath.endswith('.csv'):
            return pd.read_csv(filepath)
        elif filepath.endswith('.xlsx') or filepath.endswith('.xls'):
            return pd.read_excel(filepath)
        else:
            print(f"AVERTISSEMENT: Format non supporté - {filepath}")
            return None
    except Exception as e:
        print(f"ERREUR: Échec du chargement de {filepath} - {e}")
        return None


def load_all_brutes():
    """Charger tous les fichiers du dossier Brutes"""
    frames = []
    
    if not os.path.exists(BRUTES_DIR):
        print(f"ERREUR: Dossier {BRUTES_DIR} introuvable")
        return pd.DataFrame()
    
    files = os.listdir(BRUTES_DIR)
    if not files:
        print(f"AVERTISSEMENT: Dossier {BRUTES_DIR} vide")
        return pd.DataFrame()
    
    for filename in files:
        filepath = os.path.join(BRUTES_DIR, filename)
        if os.path.isfile(filepath):
            df = load_file(filepath)
            if df is not None:
                frames.append(df)
    
    if not frames:
        print("ERREUR: Aucun fichier valide trouvé dans Brutes")
        return pd.DataFrame()
    
    return pd.concat(frames, ignore_index=True)


def build_concat():
    """Créer le fichier concaténé avec dédoublonnage"""
    try:
        df = load_all_brutes()
        
        if df.empty:
            print("ERREUR: DataFrame vide après chargement")
            return None
        
        # Dédoublonnage
        initial_count = len(df)
        df = df.drop_duplicates()
        final_count = len(df)
        
        if initial_count != final_count:
            print(f"Dédoublonnage: {initial_count - final_count} doublons supprimés")
        
        # Génération du nom de fichier avec timestamp
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = os.path.join(CONCAT_DIR, f"SMS_Concat_{ts}.xlsx")
        
        # Sauvegarde
        df.to_excel(out_file, index=False)
        print(f"SUCCESS: Fichier concaténé créé - {out_file}")
        
        return out_file
        
    except Exception as e:
        print(f"ERREUR: Échec de la création du fichier concaténé - {e}")
        return None


def archive_brutes():
    """Déplacer les fichiers Brutes vers l'archive"""
    try:
        if not os.path.exists(BRUTES_DIR):
            print(f"AVERTISSEMENT: Dossier {BRUTES_DIR} introuvable, rien à archiver")
            return True
        
        files = os.listdir(BRUTES_DIR)
        if not files:
            print(f"AVERTISSEMENT: Dossier {BRUTES_DIR} vide, rien à archiver")
            return True
        
        for filename in files:
            src = os.path.join(BRUTES_DIR, filename)
            dst = os.path.join(ARCHIVE_DIR, filename)
            
            # Gérer les conflits de noms
            counter = 1
            while os.path.exists(dst):
                name, ext = os.path.splitext(filename)
                dst = os.path.join(ARCHIVE_DIR, f"{name}_{counter}{ext}")
                counter += 1
            
            shutil.move(src, dst)
            print(f"Archivé: {filename} -> {dst}")
        
        print("SUCCESS: Tous les fichiers Brutes archivés")
        return True
        
    except Exception as e:
        print(f"ERREUR: Échec de l'archivage - {e}")
        return False


def load_tp_file(tp_path=TP_FILE):
    """Charger le fichier TP"""
    try:
        if not os.path.exists(tp_path):
            print(f"ERREUR: Fichier TP introuvable - {tp_path}")
            return None
        
        tp_df = pd.read_excel(tp_path, header=0)
        
        # Vérifier les colonnes requises
        # Le fichier peut avoir des colonnes numérotées (0, 1, 2, ...) ou nommées
        if 0 in tp_df.columns:
            # Renommer les colonnes basées sur shared strings
            column_mapping = {
                0: 'Identité',
                1: 'Nationalité', 
                2: 'MSISDN',
                3: 'IMSI',
                4: 'IMEI',
                5: 'Niveau',
                6: 'Lien'
            }
            tp_df = tp_df.rename(columns=column_mapping)
        
        required_cols = ['Identité', 'MSISDN']
        for col in required_cols:
            if col not in tp_df.columns:
                print(f"ERREUR: Colonne manquante dans TP - {col}")
                print(f"Colonnes disponibles: {tp_df.columns.tolist()}")
                return None
        
        # Nettoyer les données: convertir MSISDN en string et supprimer les lignes vides
        tp_df['MSISDN'] = tp_df['MSISDN'].astype(str)
        tp_df = tp_df.dropna(subset=['MSISDN'])
        
        print(f"SUCCESS: Fichier TP chargé ({len(tp_df)} lignes)")
        return tp_df
        
    except Exception as e:
        print(f"ERREUR: Échec du chargement du fichier TP - {e}")
        return None


def create_tp_sheets(tp_df, output_file):
    """Créer un fichier Excel avec un onglet par MSISDN du TP"""
    try:
        wb = Workbook()
        
        # Supprimer la feuille par défaut
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])
        
        sheet_map = {}
        
        for _, row in tp_df.iterrows():
            sheet_name = f"{row['Identité']} - {row['MSISDN']}"
            
            # Excel limite à 31 caractères
            sheet_name = sheet_name[:31]
            
            # Gérer les doublons de noms d'onglets
            original_name = sheet_name
            counter = 1
            while sheet_name in wb.sheetnames:
                sheet_name = f"{original_name[:25]}_{counter}"
                counter += 1
            
            ws = wb.create_sheet(title=sheet_name)
            
            # Écrire l'en-tête
            headers = ["ActionDate", "MSISDN APT", "MSISDN APE", "Preview", "Traduction"]
            ws.append(headers)
            
            sheet_map[row["MSISDN"]] = sheet_name
        
        wb.save(output_file)
        print(f"SUCCESS: Fichier avec onglets TP créé - {output_file}")
        return sheet_map
        
    except Exception as e:
        print(f"ERREUR: Échec de la création des onglets TP - {e}")
        return None


def route_sms(concat_file, tp_df, output_file):
    """Router les SMS dans les onglets correspondants"""
    try:
        # Charger le fichier concaténé
        df = pd.read_excel(concat_file)
        
        # Vérifier les colonnes requises
        required_cols = ['ActionDate', 'MSISDN APT', 'MSISDN APE', 'Preview', 'Traduction']
        for col in required_cols:
            if col not in df.columns:
                print(f"ERREUR: Colonne manquante dans concat - {col}")
                return False
        
        # Charger le workbook de sortie
        wb = load_workbook(output_file)
        
        # Créer le mapping MSISDN -> nom d'onglet
        msisdn_to_sheet = {}
        for _, row in tp_df.iterrows():
            sheet_name = f"{row['Identité']} - {row['MSISDN']}"[:31]
            # Trouver le nom réel (peut avoir été modifié pour les doublons)
            for real_name in wb.sheetnames:
                if real_name.startswith(sheet_name):
                    msisdn_to_sheet[row["MSISDN"]] = real_name
                    break
        
        print(f"Mapping MSISDN -> Onglets: {len(msisdn_to_sheet)} entrées")
        
        # Parser chaque ligne et copier dans les onglets
        for _, row in df.iterrows():
            apt = str(row["MSISDN APT"])
            ape = str(row["MSISDN APE"])
            
            row_data = [
                row["ActionDate"],
                row["MSISDN APT"],
                row["MSISDN APE"],
                row["Preview"],
                row["Traduction"]
            ]
            
            # Écrire dans APT si présent dans TP
            if apt in msisdn_to_sheet:
                ws = wb[msisdn_to_sheet[apt]]
                ws.append(row_data)
            
            # Écrire dans APE si différent et présent dans TP
            if ape in msisdn_to_sheet and ape != apt:
                ws = wb[msisdn_to_sheet[ape]]
                ws.append(row_data)
        
        wb.save(output_file)
        print(f"SUCCESS: Routing SMS terminé - {output_file}")
        return True
        
    except Exception as e:
        print(f"ERREUR: Échec du routing SMS - {e}")
        return False


def get_latest_concat():
    """Récupérer le dernier fichier concaténé créé"""
    try:
        if not os.path.exists(CONCAT_DIR):
            return None
        
        files = [f for f in os.listdir(CONCAT_DIR) if f.startswith("SMS_Concat_") and f.endswith(".xlsx")]
        
        if not files:
            return None
        
        # Trier par date de modification
        files.sort(key=lambda x: os.path.getmtime(os.path.join(CONCAT_DIR, x)), reverse=True)
        
        return os.path.join(CONCAT_DIR, files[0])
        
    except Exception as e:
        print(f"ERREUR: Échec de la récupération du dernier concat - {e}")
        return None


def run_pipeline():
    """Exécuter le pipeline complet"""
    print("=" * 60)
    print("DEBUT DU PIPELINE SMS")
    print("=" * 60)
    
    # Étape 1: Vérifier/créer les répertoires
    print("\n[1/5] Vérification des répertoires...")
    ensure_directories()
    
    # Étape 2: Concaténer les fichiers Brutes
    print("\n[2/5] Concaténation des fichiers Brutes...")
    concat_file = build_concat()
    
    if concat_file is None:
        print("ERREUR FATALE: Impossible de créer le fichier concaténé")
        return False
    
    # Étape 3: Archiver les fichiers Brutes
    print("\n[3/5] Archivage des fichiers Brutes...")
    if not archive_brutes():
        print("AVERTISSEMENT: L'archivage a échoué, continuation...")
    
    # Étape 4: Charger le fichier TP
    print("\n[4/5] Chargement du fichier TP...")
    tp_df = load_tp_file()
    
    if tp_df is None:
        print("ERREUR FATALE: Impossible de charger le fichier TP")
        return False
    
    # Étape 5: Créer les onglets et router les SMS
    print("\n[5/5] Création des onglets et routing SMS...")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"TP_SMS_{ts}.xlsx"
    
    sheet_map = create_tp_sheets(tp_df, output_file)
    
    if sheet_map is None:
        print("ERREUR FATALE: Impossible de créer les onglets TP")
        return False
    
    if not route_sms(concat_file, tp_df, output_file):
        print("ERREUR FATALE: Le routing SMS a échoué")
        return False
    
    print("\n" + "=" * 60)
    print("PIPELINE TERMINE AVEC SUCCES")
    print(f"Fichier de sortie: {output_file}")
    print("=" * 60)
    
    return True


def run_from_existing_concat():
    """Exécuter à partir du dernier fichier concat existant"""
    print("=" * 60)
    print("DEBUT DU TRAITEMENT A PARTIR DU DERNIER CONCAT")
    print("=" * 60)
    
    # Récupérer le dernier concat
    print("\n[1/3] Recherche du dernier fichier concaténé...")
    concat_file = get_latest_concat()
    
    if concat_file is None:
        print("ERREUR: Aucun fichier concaténé trouvé dans {CONCAT_DIR}")
        print("Exécutez d'abord run_pipeline() pour créer un fichier concaténé")
        return False
    
    print(f"Dernier concat trouvé: {concat_file}")
    
    # Charger le fichier TP
    print("\n[2/3] Chargement du fichier TP...")
    tp_df = load_tp_file()
    
    if tp_df is None:
        print("ERREUR FATALE: Impossible de charger le fichier TP")
        return False
    
    # Créer les onglets et router les SMS
    print("\n[3/3] Création des onglets et routing SMS...")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"TP_SMS_{ts}.xlsx"
    
    sheet_map = create_tp_sheets(tp_df, output_file)
    
    if sheet_map is None:
        print("ERREUR FATALE: Impossible de créer les onglets TP")
        return False
    
    if not route_sms(concat_file, tp_df, output_file):
        print("ERREUR FATALE: Le routing SMS a échoué")
        return False
    
    print("\n" + "=" * 60)
    print("TRAITEMENT TERMINE AVEC SUCCES")
    print(f"Fichier de sortie: {output_file}")
    print("=" * 60)
    
    return True


if __name__ == "__main__":
    # Mode d'emploi
    if len(sys.argv) > 1 and sys.argv[1] == "--from-existing":
        success = run_from_existing_concat()
    else:
        success = run_pipeline()
    
    sys.exit(0 if success else 1)
