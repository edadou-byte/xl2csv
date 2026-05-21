#!/usr/bin/env python
'''
Created: 2026-05-21
@author: edadou
documentation : lire une plage de cellule (xlsx) et la transformer en fichier csv
'''

#############
# CONSTANTS
#############
PARAMS_SPEC = {
    'context_file': {'type': 'str', 'var_name': 'context_file'},
    'input_file': {'type': 'str', 'var_name': 'input_file'}
}

global SHEET_NAME
global HEADERS_DIMENSION
global DATA_DIMENSION

#############
# IMPORTS
#############
# From standard library
import sys, os
import logging
import argparse
import csv
import json
from collections import OrderedDict
from dotenv import dotenv_values
from pathlib import Path
import shutil

# From community
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

#############
# INIT
#############
LOG_FORMAT_STD = '{asctime}.{msecs:03.0f}: {levelname}: {filename}: {message}'
LOG_FORMAT_DEBUG = '{asctime}.{msecs:03.0f}: {levelname}: {filename} - {name} - L{lineno}: {message}'

logger = logging.getLogger()
logging.basicConfig(level=logging.INFO, style='{', format=LOG_FORMAT_STD,
                    datefmt='(%z)%Y-%m-%dT%H:%M:%S')


# logger.setLevel(logging.DEBUG)


def csv_file_preparation(filepath: str, headers: list, delimiter_int: int):
    """
    Prépare un fichier CSV en écrivant la ligne d'en-têtes (headers).

    Si headers est fourni, chaque tuple de la liste est aplati pour former
    la première ligne du fichier CSV. Si headers est None ou vide, aucune
    ligne d'en-tête n'est écrite.
    Le fichier est ouvert en mode écriture ('w'), ce qui écrase le fichier
    existant s'il y en a un.

    Args:
        filepath (str): Chemin complet vers le fichier CSV cible.
                        Exemple : "data/export.csv"
        headers (list): Liste de tuples contenant les noms de colonnes.
                        Chaque tuple est aplati pour former la ligne d'en-tête.
                        Si None ou "", aucun header n'est écrit.
        delimiter_int (int): Code ASCII du caractère séparateur de colonnes.
                             Exemple : 44 pour ',', 59 pour ';', 9 pour '\\t', 124 pour '|'

    Returns:
        None

    Raises:
        FileNotFoundError: Si le répertoire parent de filepath n'existe pas.
                           L'erreur est catchée et loggée sans être reraisée.
        Exception: Toute autre erreur inattendue est catchée et loggée.

    Example:
        >>> csv_file_preparation("export.csv", [("email", "nom", "date")], 59)
        # Crée ou écrase export.csv avec la ligne : email;nom;date

        >>> csv_file_preparation("export.csv", None, 59)
        # Aucun header écrit, fichier non modifié

    Notes:
        - Les tuples de headers sont aplatis en une liste plate avant écriture.
        - Le bloc finally logue la fin de la préparation, qu'il y ait eu une erreur ou non.
    """
    logger.info(f"Préparation du fichier csv {filepath}")
    try:
        if headers is not None and headers != "":
            logger.info(f"  Préparation des headers pour le fichier csv -> {headers}")

            headers_list = [header for tuple_ in headers for header in tuple_]

            with open(filepath, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file, delimiter=chr(delimiter_int))
                writer.writerow(headers_list)
        else:
            logger.info(f"  Pas de headers a spécifier")
    except FileNotFoundError:
        logger.error(f" Fichier introuvable : {filepath}")
    except Exception as e:
        logger.error(f" Erreur : {e}")
    finally:
        logger.info(f"  Fin de préparation des headers pour le fichier csv")


def csv_write_row(filepath: str, delimiter_int: int, rows: list):
    """
    Ajoute une ligne dans un fichier CSV existant ou le crée s'il n'existe pas.

    Le fichier est ouvert en mode append ('a'), ce qui signifie que les données
    sont ajoutées à la suite du contenu existant sans l'écraser.

    Args:
        filepath (str): Chemin complet vers le fichier CSV cible.
                        Exemple : "data/export.csv"
        delimiter_int (int): Code ASCII du caractère séparateur de colonnes.
                             Exemple : 44 pour ',', 59 pour ';', 9 pour '\\t', 124 pour '|'
        rows (list): Liste de valeurs représentant une seule ligne à écrire.
                     Exemple : ["test@mail.com", "Dupont"]

    Returns:
        None

    Raises:
        FileNotFoundError: Si le répertoire parent de filepath n'existe pas.
                           L'erreur est catchée et loggée sans être reraisée.
        PermissionError: Si l'accès en écriture au fichier est refusé.
                         L'erreur est catchée et loggée sans être reraisée.
        ValueError: Si delimiter_int ne correspond pas à un caractère valide.
                    L'erreur est catchée et loggée sans être reraisée.
        Exception: Toute autre erreur inattendue est catchée et loggée.

    Example:
        >>> row = ["test@email.com", "Dupont"]
        >>> csv_write_row("export.csv", 59, row)  # séparateur ';'

    Notes:
        - Un seul appel à writer.writerow() est effectué : rows représente une
          ligne unique, et non une liste de lignes.
        - Le bloc finally logue la fin de l'écriture, qu'il y ait eu une erreur ou non.
    """
    logger.info(f"Ecriture de 1 ligne dans le fichier csv : {filepath}")
    try:
        with open(filepath, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, delimiter=chr(delimiter_int))
            writer.writerow(rows)

    except FileNotFoundError:
        logger.error(f" Fichier introuvable : {filepath}")
    except PermissionError:
        logger.error(f" Permission refusée pour écrire dans : {filepath}")
    except ValueError as e:
        logger.error(f" Délimiteur invalide (delimiter_int={delimiter_int}) : {e}")
    except Exception as e:
        logger.error(f" Erreur inattendue : {e}")
    finally:
        logger.info(f"  Fin d'écriture de la ligne dans : {filepath}")


def get_rows_from_xlsx(ws: openpyxl.worksheet, dimensions: tuple) -> list:
    """
    Extrait un ensemble de lignes d'une feuille Excel à partir de dimensions données.

    Parcourt la plage de cellules définie par le tuple dimensions et retourne
    les valeurs sous forme de liste de tuples, où chaque tuple représente une ligne.

    Args:
        ws (openpyxl.worksheet): Feuille Excel active issue d'un workbook openpyxl.
                                 Doit être ouverte avec data_only=True pour obtenir
                                 les valeurs calculées plutôt que les formules brutes.
                                 Exemple : wb = openpyxl.load_workbook(filepath, data_only=True)

        dimensions (tuple): Tuple de 4 entiers définissant la plage de lecture.
                            Format : (min_col, max_col, min_row, max_row)
                            Exemple : (1, 5, 1, 10) → colonnes 1 à 5, lignes 1 à 10

    Returns:
        list: Liste de tuples représentant les lignes lues.
              Exemple : [('email', 'nom', 'date'), ('a@mail.com', 'Dupont', '2024-01-01')]
              Retourne None en cas d'erreur.

    Raises:
        IndexError: Si le tuple dimensions est mal formé ou incomplet.
                    L'erreur est catchée et loggée sans être reraisée.
        Exception: Toute autre erreur inattendue est catchée et loggée sans être reraisée.

    Example:
        >>> wb = openpyxl.load_workbook("export.xlsx", data_only=True)
        >>> ws = wb.active
        >>> rows = get_rows_from_xlsx(ws, (1, 5, 1, 10))
        >>> print(rows)  # [('email', 'nom', ...), ('a@mail.com', 'Dupont', ...)]

    Notes:
        - Le tuple dimensions suit l'ordre (min_col, max_col, min_row, max_row).
        - Si le fichier n'a jamais été sauvegardé par Excel, data_only=True
          peut retourner des None à la place des valeurs calculées.
        - Le bloc finally logue la fin de l'opération qu'il y ait une erreur ou non.
    """
    try:
        data_tuple = list(
            ws.iter_rows(min_col=dimensions[0], max_col=dimensions[1], min_row=dimensions[2], max_row=dimensions[3],
                         values_only=True))
        return data_tuple
    except IndexError:
        logger.error(
            f" Vérifiez le fichier template.json et que headers_dimension / datas_dimension soit au bon format ")
    except Exception as e:
        logger.error(f" Erreur inattendue : {e}")
    finally:
        logger.info(f" Récupération des datas - OK")


def parse_template(contexts: OrderedDict):
    """
    Parse un fichier template JSON et en extrait les paramètres de chaque étape.

    Lit le fichier JSON situé à l'emplacement défini par contexts["S_TEMPLATE_PATH"]
    et parcourt chaque étape ("steps") pour en extraire le nom de la feuille,
    les plages de headers, les plages de données et les noms d'export.

    Args:
        contexts (OrderedDict): Dictionnaire de contexte contenant au minimum la clé
                                "S_TEMPLATE_PATH" pointant vers le fichier JSON template.
                                Exemple : {"S_TEMPLATE_PATH": "config/template.json"}

    Returns:
        tuple: Un tuple de 4 éléments, dans l'ordre :
            - sheets (tuple)           : Noms des feuilles Excel par étape.
            - headers_dimension (tuple): Plages des headers par étape, chacune sous la forme
                                         (col_start, col_end, row_start, row_end).
            - data_dimension (tuple)   : Plages des données par étape, chacune sous la forme
                                         (col_start, col_end, row_start, row_end).
            - exports_names (tuple)    : Noms des fichiers d'export par étape.
        Retourne None en cas d'erreur.

    Raises:
        FileNotFoundError: Si le fichier template est introuvable.
                           L'erreur est catchée et loggée sans être reraisée.
        PermissionError: Si l'accès en lecture au fichier est refusé.
                         L'erreur est catchée et loggée sans être reraisée.
        Exception: Toute autre erreur inattendue est catchée et loggée sans être reraisée.

    Example:
        >>> contexts = {"S_TEMPLATE_PATH": "config/template.json"}
        >>> sheets, headers_dim, data_dim, exports = parse_template(contexts)
        >>> print(sheets)        # ('Feuille1', 'Feuille2')
        >>> print(headers_dim)   # ((1, 5, 1, 1), (1, 3, 1, 1))
        >>> print(data_dim)      # ((1, 5, 2, 100), (1, 3, 2, 50))
        >>> print(exports)       # ('export_step1.csv', 'export_step2.csv')

    Notes:
        - Le fichier JSON doit contenir une clé "steps" avec pour chaque étape :
          "sheet_name", "headers_dimension" (col_start, col_end, row_start, row_end),
          "data_dimension" (col_start, col_end, row_start, row_end) et "export_name".
        - Les plages sont au format (col_start, col_end, row_start, row_end),
          compatible avec la fonction get_rows_from_xlsx().
    """
    try:
        with open(contexts["S_TEMPLATE_PATH"], 'r', encoding='utf-8') as file:
            data = json.load(file)

        sheets = tuple()
        headers_dimension = tuple()
        data_dimension = tuple()
        exports_names = tuple()

        for i in range(len(data["steps"])):
            sheets = (*sheets, data["steps"][i]["sheet_name"])

            logger.info(f" Récupération du nom de la feuille pour l'étape {i} - OK")

            headers_d = (
                data["steps"][i]["headers_dimension"]["col_start"],
                data["steps"][i]["headers_dimension"]["col_end"],
                data["steps"][i]["headers_dimension"]["row_start"],
                data["steps"][i]["headers_dimension"]["row_end"]
            )

            headers_dimension = (*headers_dimension, headers_d)

            logger.info(f" Récupération des plages pour les headers pour l'étape {i} - OK")

            data_d = (
                data["steps"][i]["data_dimension"]["col_start"],
                data["steps"][i]["data_dimension"]["col_end"],
                data["steps"][i]["data_dimension"]["row_start"],
                data["steps"][i]["data_dimension"]["row_end"]
            )

            data_dimension = (*data_dimension, data_d)

            logger.info(f" Récupération des plages de données pour l'étape {i} - OK")

            exports_names = (*exports_names, data["steps"][i]["export_name"])

            logger.info(f" Récupération du nom des exports données pour l'étape {i} - OK")
        print(exports_names)
        return sheets, headers_dimension, data_dimension, exports_names
    except FileNotFoundError:
        logger.error(f" Fichier introuvable : {contexts["S_TEMPLATE_PATH"]}")
    except PermissionError:
        logger.error(f" Permission refusée pour écrire dans : {contexts["S_TEMPLATE_PATH"]}")
    except Exception as e:
        logger.error(f" Erreur inattendue : {e}")


#############
# MAIN
#############
if __name__ == '__main__':
    logger.info("STARTING...\n")
    logger.info("  Récupération des paramètres dans les variables d'environnement...")

    try:
        params = {}
        for item, item_spec in PARAMS_SPEC.items():
            params[item] = None
    except Exception as e:
        logger.error("  Erreur lors de l'initialisation du dictionnaire de variables...")

    parser = argparse.ArgumentParser(description='Convert a delimited file (tsv, csv, dsv...) to xlsx')
    parser.add_argument('-cf', '--context_file', type=str, help='Path to the context file')
    parser.add_argument('-fn', '--input_file', type=str, help='XLSX file to process')
    args = parser.parse_args()

    logger.info(f"Arguments en lignes de commande (écrasent les variables d'environnement) :")
    logger.info(f"  {args}")
    for key, val in vars(args).items():
        if val is not None:
            params[key] = val

    config = dotenv_values(params["context_file"])

    input_filepath_without_ext, input_ext = os.path.splitext(params['input_file'])
    filename = Path(params['input_file']).stem

    s_separator = ""
    s_output_filepath = ""

    if config["S_SEPARATOR"] == "":
        s_separator = config["S_DELIMITEUR"] or ";"
    if config["dir_out"] is None or config["dir_out"] == "":
        s_output_filepath = input_filepath_without_ext + ".csv"

    logger.info(f"Params after applying defaults :")
    logger.info(f"  {params}")

    logger.info(f"Openning {params['input_file']}")

    try:
        wb = openpyxl.load_workbook(params['input_file'], data_only=True)
        sheets, headers_dimension, datas_dimension, exports_names = parse_template(config)

        if len(sheets) == len(datas_dimension):
            for i in range(len(sheets)):
                try:
                    pathfile = config["default_path"] + config["dir_out"] + exports_names[i]
                    Path(pathfile).parent.mkdir(parents=True, exist_ok=True)

                    ws = wb[sheets[i]]

                    if (headers_dimension != None) and (sum(headers_dimension[i]) > 0):
                        headers_string = get_rows_from_xlsx(ws, headers_dimension[i])
                        csv_file_preparation(pathfile, headers_string, int(config["S_SEPARATOR"]))
                    if (datas_dimension != None) and (sum(datas_dimension[i]) > 0):
                        datas_string = get_rows_from_xlsx(ws, datas_dimension[i])
                        for i in range(len(datas_string)):
                            csv_write_row(pathfile, int(config["S_SEPARATOR"]), datas_string[i])

                    source = params['input_file']
                    destination = config["default_path"] + config["dir_out"] + filename + ".xlsx"

                    if os.path.exists(source):
                        shutil.move(source, destination)
                        print("Fichier déplacé avec succès")
                    else:
                        print("Fichier introuvable")

                except Exception as e:
                    logger.error(f" Erreur inattendue : {e}")

                    source = params['input_file']
                    destination = config["default_path"] + config["dir_rej"] + filename + ".xlsx"

                    if os.path.exists(source):
                        shutil.move(source, destination)
                        print("Fichier déplacé avec succès")
                    else:
                        print("Fichier introuvable")
                finally:
                    logger.info(f"Ecriture du fichier : {pathfile} - OK")

    except InvalidFileException:
        logger.error(f"Failed to open the file {params['input_filepath']}")
    finally:
        logger.info(f"Transformation du fichier XLSX en CSV - OK")
