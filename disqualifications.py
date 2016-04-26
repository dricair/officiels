# -* -coding: utf-8 -*-

"""
.. module: officiels.py
   :platform: Unix, Windows
   :synopsys: Parse list of competitions to get disqualifications

.. moduleauthor: Cedric Airaud <cairaud@gmail.com>
"""

import os.path
import argparse
import zipfile
import logging
import tempfile
import xml.etree.ElementTree
import datetime
import pandas as pd
import shutil

from officiels_from_ffnex import Configuration
from openpyxl import load_workbook

logging.basicConfig(level=logging.DEBUG, format="%(levelname)-9s %(lineno)-4s %(message)s")

# List of categories per age
categories = dict()
year = datetime.date.today().year

for keys, value in {range(0,  10): "Avenir",
                    range(10, 12): "Poussin",
                    range(12, 14): "Benjamin",
                    range(14, 99): "Minime+"}.items():
    for key in keys:
        categories[year - key] = value


def get_disqualifications(filename, conf):
    """
    :param filename: FFNEX file to parse
    :type filename: string
    :param conf: Configuration structure
    :type conf: Configuration
    :return: List of disqualifications
    :rtype: list of Disqualification
    """
    try:
        if zipfile.is_zipfile(filename):
            z = zipfile.ZipFile(filename, 'r')
            if "ffnex.xml" not in z.namelist():
                logging.error("Le fichier {} devrait contenir un fichier ffnex.xml")
                return
            filename = z.extract('ffnex.xml', tempfile.gettempdir())
            z.close()

    except zipfile.BadZipfile:
        logging.error("Le fichier {} ne peut pas être lu correctement".format(filename))
        return

    # Header
    e = xml.etree.ElementTree.parse(filename).getroot()
    competition = e.find("MEETS").find("MEET")
    competition_id = int(competition.attrib["id"])
    nom = competition.attrib["name"]
    startdate = datetime.datetime.strptime(competition.attrib["startdate"], "%Y-%m-%d")
    stopdate = datetime.datetime.strptime(competition.attrib["stopdate"], "%Y-%m-%d")
    ville = competition.attrib["city"]
    par_equipe = True if competition.attrib.get("byteam", "false") == "true" else 0
    niveau = conf.type_competitions[int(competition.attrib["typeid"])][1]

    date_str = startdate.strftime("%d/%m/%Y")
    if stopdate != startdate:
        date_str += " - " + stopdate.strftime("%d/%m/%Y")

    logging.info("Lecture de la compétition {} ({}) - {} à {} - {}".format(nom, competition_id, date_str, ville, niveau))

    # Clubs
    clubs = {}
    for c in competition.find("CLUBS").findall("CLUB"):
        clubs[int(c.attrib["id"])] = c.attrib["name"]

    # Swimmers
    swimmers = {}
    for s in competition.find("SWIMMERS").findall("SWIMMER"):
        idx = int(s.attrib["id"])
        swimmers[idx] = {"année": datetime.datetime.strptime(s.attrib["birthdate"], "%Y-%m-%d").year,
                         "club": clubs[int(s.attrib["clubid"])],
                         "sexe": s.attrib["gender"]}
        swimmers[idx]["catégorie"] = categories[swimmers[idx]["année"]]

    # Sessions - Store date/time for each race - Indexes by (raceid, roundid) strings.
    events = dict()
    for s in competition.find("SESSIONS").findall("SESSION"):
        for e in s.find("EVENTS").findall("EVENT"):
            if "raceid" in e.attrib:
                date = datetime.datetime.strptime(e.attrib["datetime"], "%Y-%m-%d %H:%M:%S")
                events[(e.attrib["raceid"], e.attrib["roundid"])] = date

    # Races and disqualifications
    disqualifications = []
    for r in competition.find("RESULTS").findall("RESULT"):
        disqualification = int(r.attrib["disqualificationid"])
        if disqualification == 0:
            continue

        reason, libelle, relayeur = conf.disqualifications[disqualification]
        if reason in ("DNS exc", "DNS dec", "DNS Nd", "DSQ", "DNS", "FD", "DNF", "EPR Supp"):
            continue

        race, nage, sexe = conf.nages[int(r.attrib["raceid"])]

        heat, lane = int(r.attrib["heat"]), int(r.attrib["lane"])

        if relayeur is None:
            if r.find("RELAY"):
                raise Exception("Disqualification {}: relayeur=0 pour RELAY".format(reason))
            swimmer = swimmers[int(r.find("SOLO").attrib["swimmerid"])]

        else:
            if r.find("RELAY") is None:
                raise Exception("Disqualification {}: relayeur={} pour SOLO".format(reason, relayeur))
            relayposition = r.find("RELAY").find("RELAYPOSITIONS").findall("RELAYPOSITION")[relayeur-1]
            swimmer = swimmers[int(relayposition.attrib["swimmerid"])]

        date = events[(r.attrib["raceid"], r.attrib["roundid"])]

        relayeur_str = ""
        if relayeur is not None:
            relayeur_str = ", relayeur {}".format(relayeur)
        logging.info("- Disqualification {} en {} (Ligne {}, série {}{})".format(reason, race, lane, heat,
                                                                                 relayeur_str))
        disqualification = {"Compétition": nom, "Date": date.strftime("%Y-%m-%d %H:%M:%S"), "Niveau": niveau,
                            "Année naissance": swimmer["année"], "Club": swimmer["club"], "Sexe": swimmer["sexe"],
                            "Catégorie": swimmer["catégorie"], "Disqualification": reason,
                            "Disqualification (libellé)": libelle,
                            "Nage (Complet)": race, "Nage": nage, "Série": heat, "Ligne": lane}
        disqualifications.append(disqualification)

    return disqualifications


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Liste des compétitions')
    parser.add_argument("--conf", default="Officiels.xlsx", help="Fichier de configuration")
    parser.add_argument("--output", default="Disqualifications.xlsx", help="Fichier de sortie")
    parser.add_argument("--sheet", default="Données brutes", help="Page dans le fichier Excel")
    parser.add_argument("ffnex_files", metavar="fichiers", nargs="+", help="Liste des fichiers ou répertoires " +
                                                                           "à analyser")

    args = parser.parse_args()
    conf = Configuration(args.conf)

    ffnex_files = []
    for f in args.ffnex_files:
        if os.path.isdir(f):
            ffnex_files += [os.path.join(f, file) for file in os.listdir(f) if os.path.isfile(os.path.join(f, file))]
        else:
            ffnex_files.append(f)

    ffnex_files.sort()

    disqualifications = []
    for f in ffnex_files:
        disqualifications += get_disqualifications(f, conf)

    disq_df = pd.DataFrame(disqualifications)

    writer = pd.ExcelWriter(args.output, engine='openpyxl')
    if os.path.exists(args.output):
        f, ext = os.path.splitext(args.output)
        shutil.copyfile(args.output, f + "-saved" + ext)
        writer.book = load_workbook(args.output)
        writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)

    disq_df.to_excel(writer, sheet_name=args.sheet)
    writer.save()