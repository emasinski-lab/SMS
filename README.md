# SMS Processing Pipeline

Script Python pour le traitement des SMS en environnement hors réseau (Python 3.9+).

## Structure du projet

```
SMS/
├── Brutes/                  # Dossier des fichiers bruts (CSV/XLSX)
│   ├── sms_part_1.csv
│   ├── sms_part_2.csv
│   └── sms_part_3.csv
├── TP/                      # Dossier du fichier TP
│   └── TP_test.xlsx
├── Concat/                  # Dossier des fichiers concaténés (créé automatiquement)
├── Sortie/                  # Dossier des fichiers de sortie (créé automatiquement)
├── Brutes_archive/          # Archive des fichiers bruts (créé automatiquement)
├── Scripts/                 # Scripts originaux
├── main.py                  # Script principal
├── test_structure.py        # Tests de structure
└── requirements.txt         # Dépendances
```

## Prérequis

- Python 3.9.x
- Dépendances: `pandas`, `openpyxl`

### Installation des dépendances (si en ligne)

```bash
pip install -r requirements.txt
```

### Installation hors ligne

Téléchargez les packages manuellement et installez-les:

```bash
pip install pandas-1.3.5-cp39-cp39-win_amd64.whl
pip install openpyxl-3.0.9-py2.py3-none-any.whl
```

## Utilisation

### 1. Pipeline complet (concaténation + traitement)

```bash
python main.py
```

Ce mode:
- Concatène tous les fichiers du dossier `Brutes/` en un seul fichier Excel dans `Concat/`
- Dédoublonne les données
- Archive les fichiers bruts dans `Brutes_archive/`
- Crée un fichier Excel avec un onglet par MSISDN du fichier TP
- Route les SMS dans les onglets correspondants

### 2. Traitement à partir du dernier concat existant

```bash
python main.py --from-existing
```

Ce mode:
- Utilise le dernier fichier concaténé créé dans `Concat/`
- Crée un fichier Excel avec un onglet par MSISDN du fichier TP
- Route les SMS dans les onglets correspondants

## Format des fichiers

### Fichiers Brutes (CSV/XLSX)

Les fichiers doivent contenir les colonnes suivantes:
- `ActionDate`: Date et heure de l'action
- `MSISDN APT`: MSISDN de l'expéditeur
- `MSISDN APE`: MSISDN du destinataire
- `Preview`: Message original
- `Traduction`: Traduction du message

Exemple:
```csv
ActionDate,MSISDN APT,MSISDN APE,Preview,Traduction
2026/02/17 13:16:43,33612345678,4915234567890,Ваш код подтверждения 918273,Your verification code is 918273
```

### Fichier TP (XLSX)

Le fichier TP doit contenir au minimum les colonnes:
- `Identité`: Nom de la personne
- `MSISDN`: Numéro de téléphone

Exemple de structure:
```
Identité    | Nationalité | MSISDN      | IMSI          | IMEI          | Niveau | Lien
Dupont Jean | Française    | 33612345678 | 208011234567890 | ... | Père
Dupont Marie| Française    | 33623456789 | 208011234567891 | ... | Mère
```

## Sortie

### Fichier concaténé

Généré dans `Concat/SMS_Concat_<timestamp>.xlsx`
- Contient toutes les données des fichiers bruts
- Dédoublonné
- Format Excel

### Fichier de sortie

Généré dans `Sortie/TP_SMS_<timestamp>.xlsx`
- Un onglet par MSISDN du fichier TP
- Nom des onglets: `Identité - MSISDN` (troncature à 31 caractères)
- Chaque ligne de SMS est copiée dans:
  - L'onglet correspondant au MSISDN APT (expéditeur)
  - L'onglet correspondant au MSISDN APE (destinataire), si différent
- **Format des valeurs**: Les numéros (MSISDN, etc.) sont formatés en texte pour éviter la notation scientifique

## Gestion des erreurs

Le script gère les erreurs suivantes:
- Fichiers manquants
- Formats non supportés
- Colonnes manquantes
- Doublons de noms d'onglets (ajout de suffixe numérique)
- Conflits de noms lors de l'archivage

## Exemple de sortie

```
============================================================
DEBUT DU PIPELINE SMS
============================================================

[1/5] Vérification des répertoires...

[2/5] Concaténation des fichiers Brutes...
Dédoublonnage: 5 doublons supprimés
SUCCESS: Fichier concaténé créé - Concat/SMS_Concat_20240115_143022.xlsx

[3/5] Archivage des fichiers Brutes...
Archivé: sms_part_1.csv -> Brutes_archive/sms_part_1.csv
Archivé: sms_part_2.csv -> Brutes_archive/sms_part_2.csv
Archivé: sms_part_3.csv -> Brutes_archive/sms_part_3.csv
SUCCESS: Tous les fichiers Brutes archivés

[4/5] Chargement du fichier TP...
SUCCESS: Fichier TP chargé (15 lignes)

[5/5] Création des onglets et routing SMS...
SUCCESS: Fichier avec onglets TP créé - TP_SMS_20240115_143025.xlsx
Mapping MSISDN -> Onglets: 15 entrées
SUCCESS: Routing SMS terminé - TP_SMS_20240115_143025.xlsx

============================================================
PIPELINE TERMINE AVEC SUCCES
Fichier de sortie: TP_SMS_20240115_143025.xlsx
============================================================
```

## Notes

- Les noms d'onglets Excel sont limités à 31 caractères
- Les doublons de noms d'onglets sont résolus par ajout de suffixe (`_1`, `_2`, etc.)
- Les fichiers bruts sont déplacés vers l'archive après traitement réussi
- Le timestamp utilise le format `YYYYMMDD_HHMMSS`
