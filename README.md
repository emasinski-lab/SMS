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
├── Reference/               # Dossier du fichier de référence
│   └── msisdn_reference_v1.1.csv
├── Error/                   # Dossier des fichiers d'erreur (créé automatiquement)
│   ├── errorlog_<timestamp>.txt    # Log des erreurs
│   └── SMS_error_<timestamp>.xlsx  # SMS en erreur (si applicable)
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

**Comportement intelligent:**
- Si le dossier `Brutes/` est **vide**, le script bascule automatiquement sur le **dernier fichier concaténé** existant dans `Concat/`
- Dans ce cas, l'archivage est sauté (pas de fichiers à archiver)
- Le traitement continue normalement avec les fonctions d'analyse (Accueil, PB_Format, coloration)

Exemple:
```
[2/5] Concaténation des fichiers Brutes...
AVERTISSEMENT: Dossier Brutes vide, recherche du dernier concat existant...
Dernier concat trouvé: Concat/SMS_Concat_20240115_143022.xlsx

[3/5] Archivage des fichiers Brutes: non nécessaire (mode analyse)
```

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

### Colorisation des onglets

Pour une **visualisation rapide des relations**, les MSISDN sont colorés:

- **🟡 JAUNE**: Le MSISDN correspond au MSISDN de l'onglet (correspondance directe)
  - Exemple: Dans l'onglet "Dupont Jean - 33612345678", la cellule contenant `33612345678` sera en jaune

- **🔴 ROUGE**: Le MSISDN est un autre membre du TP (relation entre membres)
  - Exemple: Dans l'onglet "Dupont Jean - 33612345678", si un SMS contient `33623456789` (qui est dans le TP), cette cellule sera en rouge

Cette coloration permet d'identifier **en un coup d'œil** quels SMS impliquent:
- Le propriétaire de l'onglet (jaune)
- D'autres membres du TP (rouge)
- Des numéros externes (non colorés)

### Onglet Accueil - Statistiques

Le fichier de sortie contient un **onglet "Accueil"** en première position avec les statistiques suivantes pour chaque membre du TP:

| Colonne | Description |
|---------|-------------|
| **Identité** | Nom de la personne |
| **MSISDN** | Numéro de téléphone |
| **Statut** | Détecté / Non détecté (si le MSISDN apparaît dans les SMS) |
| **SMS reçus** | Nombre total de SMS où le MSISDN est expéditeur (APT) ou destinataire (APE) |
| **Correspondants uniques** | Nombre de numéros différents avec qui la personne a échangé |
| **Détails correspondants** | Liste des correspondants du TP + nombre de correspondants externes |

**Exemple:**
```
Identité      | MSISDN      | Statut   | SMS reçus | Correspondants uniques | Détails correspondants
-------------|-------------|----------|-----------|------------------------|--------------------------
Dupont Jean   | 33612345678 | Détecté  | 15        | 5                      | 33623456789, 33634567890 (+2 externes)
Dupont Marie  | 33623456789 | Détecté  | 8         | 3                      | 33612345678, 34612345678
Martin Carlos | 34612345678 | Non détecté | 0         | 0                      | Aucun
```

Cet onglet permet d'avoir **une vue d'ensemble** de l'activité SMS de chaque membre du TP.

### Onglet PB_Format - Formatage des MSISDN

Le fichier de sortie contient un **onglet "PB_Format"** qui permet de standardiser les numéros de téléphone selon le fichier de référence (`Reference/msisdn_reference_v1.1.csv`).

| Colonne | Description |
|---------|-------------|
| **Donnée brute** | MSISDN tel qu'il apparaît dans les fichiers bruts |
| **Donnée formatée** | MSISDN formaté selon la norme E.164 (`+<indicatif><numéro>`) ou "Pas de correspondance" |

**Exemple:**
```
Donnée brute    | Donnée formatée
---------------|------------------
33612345678    | +33612345678
4915234567890  | +4915234567890
BANK-XYZ        | Pas de correspondance
0612345678     | Pas de correspondance
```

**Fonctionnement:**
- Le script utilise le fichier de référence pour identifier le pays et l'indicatif
- Les numéros sont formatés selon la norme **E.164** (ex: `+33612345678` pour la France)
- Si aucun format ne correspond (numéro local, format invalide, etc.), la valeur est "**Pas de correspondance**"
- Les numéros locaux (commencant par 0) ne peuvent pas être formatés sans connaître le pays

**Fichier de référence:**
- Emplacement: `Reference/msisdn_reference_v1.1.csv`
- Contient les préfixes, indicatifs et regex pour chaque pays
- Peut être mis à jour sans modifier le script

### Gestion des erreurs

Le script génère automatiquement un **dossier `Error/`** contenant :

1. **`errorlog_<timestamp>.txt`** - Fichier de log des erreurs
   - Contient tous les messages d'erreur, avertissements et informations
   - Format: `[timestamp] [type] message`
   - Types: `INFO`, `AVERTISSEMENT`, `ERREUR`, `ERREUR_FATALE`

2. **`SMS_error_<timestamp>.xlsx`** - Fichier des SMS en erreur (si applicable)
   - Contient les SMS qui n'ont pas pu être traités
   - Créé uniquement en cas d'erreur de traitement

**Exemple de log:**
```
============================================================
FICHIER DE LOG DES ERREURS
============================================================
Date de création: 2024-01-15 14:30:22
============================================================

[2024-01-15 14:30:22] [INFO] DEBUT DU PIPELINE SMS
[2024-01-15 14:30:22] [INFO] Répertoires vérifiés/créés
[2024-01-15 14:30:23] [INFO] Fichier concaténé créé: Concat/SMS_Concat_20240115_143022.xlsx
[2024-01-15 14:30:23] [INFO] Fichiers Brutes archivés avec succès
[2024-01-15 14:30:24] [INFO] Fichier TP chargé: 15 lignes
[2024-01-15 14:30:25] [INFO] PIPELINE TERMINE AVEC SUCCES
```

**Affichage console:**
```
============================================================
DEBUT DU PIPELINE SMS
============================================================

[1/5] Vérification des répertoires...

[2/5] Concaténation des fichiers Brutes...
Dédoublonnage: 5 doublons supprimés
SUCCESS: Fichier concaténé créé - Concat/SMS_Concat_20240115_143022.xlsx

[3/5] Archivage des fichiers Brutes...
SUCCESS: Tous les fichiers Brutes archivés

[4/5] Chargement du fichier TP...
SUCCESS: Fichier TP chargé (15 lignes)

[5/5] Création des onglets, routing SMS, statistiques et formatage...
SUCCESS: Routing SMS terminé avec coloration
SUCCESS: Fichier de référence chargé (1234 entrées)
SUCCESS: Onglet Accueil créé avec statistiques pour 15 membres
SUCCESS: Onglet PB_Format créé avec 45 MSISDN

============================================================
PIPELINE TERMINE AVEC SUCCES
Fichier de sortie: Sortie/TP_SMS_20240115_143025.xlsx
Fichier de log: Error/errorlog_20240115_143022.txt
============================================================
```

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
