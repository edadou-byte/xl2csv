# xl2csv

Outil Python d'export de plages de cellules Excel (.xlsx) vers des fichiers CSV, piloté par un fichier de configuration JSON (template).

---

## Fonctionnement général

Le script lit un fichier `.xlsx`, puis pour chaque **étape** définie dans le template JSON :

1. Ouvre la feuille Excel ciblée
2. Extrait la plage des **en-têtes** (headers)
3. Extrait la plage des **données**
4. Écrit le tout dans un fichier `.csv` dans le dossier de sortie

En cas de succès, le fichier `.xlsx` source est déplacé dans `dir_out`. En cas d'erreur sur une étape, il est déplacé dans `dir_rej`.

---

## Structure du projet

```
xl2csv/
├── xl2csv.py           # Script principal
├── xl2csv.ini          # Fichier de configuration (chemins, séparateur)
├── template.json       # Définition des exports (feuilles, plages, noms CSV)
├── data_in/            # Dossier de dépôt des fichiers xlsx à traiter
├── data_out/           # Dossier de sortie (CSV exportés + xlsx traités)
├── work/               # Dossier de travail intermédiaire
└── rej/                # Dossier de rejet (xlsx en erreur)
```

---

## Prérequis

**Python 3.8+** et les dépendances suivantes :

```bash
pip install openpyxl python-dotenv
```

---

## Configuration

### `xl2csv.ini`

Fichier de configuration principal, passé au script via l'argument `-cf`.

| Clé | Description |
|---|---|
| `default_path` | Chemin racine du projet (ex: `C:/projets/xl2csv/`) |
| `dir_in` | Dossier des fichiers xlsx en entrée (ex: `data_in/`) |
| `dir_out` | Dossier de sortie des CSV et xlsx traités (ex: `data_out/`) |
| `dir_work` | Dossier de travail temporaire (ex: `work/`) |
| `dir_rej` | Dossier de rejet en cas d'erreur (ex: `rej/`) |
| `S_TEMPLATE_PATH` | Chemin absolu vers le fichier `template.json` |
| `S_SEPARATOR` | Code ASCII du séparateur CSV (`59` = `;`, `44` = `,`, `9` = tabulation) |

Exemple :
```ini
default_path=C:/projets/xl2csv/
dir_in=data_in/
dir_out=data_out/
dir_work=work/
dir_rej=rej/
S_TEMPLATE_PATH=C:/projets/xl2csv/template.json
S_SEPARATOR=59
```

---

### `template.json`

Définit les exports à réaliser sous forme d'une liste d'**étapes** (`steps`). Chaque étape correspond à un fichier CSV produit.

```json
{
  "steps": [
    {
      "sheet_name": "NomDeLaFeuille",
      "export_name": "nom_du_fichier.csv",
      "headers_dimension": {
        "col_start": 1,
        "col_end": 4,
        "row_start": 1,
        "row_end": 1
      },
      "data_dimension": {
        "col_start": 1,
        "col_end": 4,
        "row_start": 2,
        "row_end": 10
      }
    }
  ]
}
```

| Champ | Description |
|---|---|
| `sheet_name` | Nom exact de la feuille dans le fichier xlsx |
| `export_name` | Nom du fichier CSV à générer |
| `headers_dimension` | Plage de cellules contenant les en-têtes (lignes/colonnes) |
| `data_dimension` | Plage de cellules contenant les données à exporter |

Les dimensions `col_start`, `col_end`, `row_start`, `row_end` utilisent la numérotation Excel (colonne A = 1, ligne 1 = 1).

Il est possible de définir **plusieurs étapes sur la même feuille**, avec des plages différentes, pour produire plusieurs CSV distincts (voir exemple avec `clients.csv` et `commandes.csv` dans le template d'exemple).

---

## Utilisation

```bash
python .\xl2csv.py -cf <chemin vers le fichier .ini> -fn "<chemin vers le fichier .xlsx>"
```

### Arguments

| Argument | Description |
|---|---|
| `-cf` / `--context_file` | Chemin vers le fichier `.ini` de configuration |
| `-fn` / `--input_file` | Chemin vers le fichier `.xlsx` à traiter |

### Exemples

```bash
# Windows
python .\xl2csv.py -cf .\xl2csv.ini -fn "C:\data\mon_fichier.xlsx"

# Linux / macOS
python ./xl2csv.py -cf ./xl2csv.ini -fn "/data/mon_fichier.xlsx"
```

---

## Résultat

- Les fichiers CSV sont générés dans `{default_path}{dir_out}` avec le séparateur configuré.
- Le fichier `.xlsx` source est **déplacé** dans `{default_path}{dir_out}` après traitement.
- En cas d'erreur sur une étape, le `.xlsx` est déplacé dans `{default_path}{dir_rej}`.
- Les logs sont affichés en console avec horodatage et niveau (`INFO`, `ERROR`).

---

## Logs

Le script produit des logs structurés en console, par exemple :

```
(+0200)2026-04-17T10:23:01.042: INFO: xl2csv.py: STARTING...
(+0200)2026-04-17T10:23:01.043: INFO: xl2csv.py: Openning mon_fichier.xlsx
(+0200)2026-04-17T10:23:01.120: INFO: xl2csv.py: Ecriture du fichier : data_out/clients.csv - OK
```
