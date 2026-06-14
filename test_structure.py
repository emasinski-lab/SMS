#!/usr/bin/env python3
"""
Script de test pour vérifier la structure sans dépendances
"""

import os
import sys

def test_directories():
    """Tester la création des répertoires"""
    print("Test des répertoires...")
    
    required_dirs = ["Brutes", "TP"]
    optional_dirs = ["Concat", "Brutes_archive"]
    
    for d in required_dirs:
        if os.path.exists(d):
            print(f"  ✓ {d} existe")
        else:
            print(f"  ✗ {d} manquant")
            return False
    
    for d in optional_dirs:
        if os.path.exists(d):
            print(f"  ✓ {d} existe")
        else:
            print(f"  - {d} (sera créé par le script)")
    
    # Vérifier les fichiers
    if os.path.exists("Brutes/sms_part_1.csv"):
        print("  ✓ Fichiers brutes présents")
    else:
        print("  ✗ Fichiers brutes manquants")
        return False
    
    if os.path.exists("TP/TP_test.xlsx"):
        print("  ✓ Fichier TP présent")
    else:
        print("  ✗ Fichier TP manquant")
        return False
    
    return True


def test_file_structure():
    """Tester la structure des fichiers"""
    print("\nTest de la structure des fichiers...")
    
    # Tester les fichiers brutes
    import csv
    for i in range(1, 4):
        filepath = f"Brutes/sms_part_{i}.csv"
        if os.path.exists(filepath):
            with open(filepath, 'r') as f:
                reader = csv.reader(f)
                header = next(reader)
                print(f"  {filepath}: colonnes = {header}")
                # Vérifier les colonnes requises
                required = ['ActionDate', 'MSISDN APT', 'MSISDN APE', 'Preview', 'Traduction']
                for col in required:
                    if col not in header:
                        print(f"    ✗ Colonne manquante: {col}")
                        return False
                print(f"    ✓ Toutes les colonnes requises présentes")
    
    return True


def test_main_script():
    """Tester que le script main.py est syntaxiquement correct"""
    print("\nTest du script main.py...")
    
    try:
        with open("main.py", "r") as f:
            code = f.read()
        
        # Vérifier la syntaxe
        compile(code, "main.py", "exec")
        print("  ✓ Syntaxe Python valide")
        
        # Vérifier les fonctions principales
        required_functions = [
            "load_file", "load_all_brutes", "build_concat", 
            "archive_brutes", "load_tp_file", "create_tp_sheets",
            "route_sms", "get_latest_concat", "run_pipeline",
            "run_from_existing_concat"
        ]
        
        for func in required_functions:
            if f"def {func}" in code:
                print(f"  ✓ Fonction {func} présente")
            else:
                print(f"  ✗ Fonction {func} manquante")
                return False
        
        return True
        
    except Exception as e:
        print(f"  ✗ Erreur: {e}")
        return False


def main():
    print("=" * 60)
    print("TESTS DE STRUCTURE")
    print("=" * 60)
    
    all_passed = True
    
    all_passed &= test_directories()
    all_passed &= test_file_structure()
    all_passed &= test_main_script()
    
    print("\n" + "=" * 60)
    if all_passed:
        print("TOUS LES TESTS PASSES ✓")
    else:
        print("CERTAINS TESTS ECHOUES ✗")
    print("=" * 60)
    
    return all_passed


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
