#! /usr/bin/env python
# -* -coding: utf-8 -*-

"""
.. module: officiels.py
   :platform: Unix, Windows
   :synopsys: Parse list of competitions to get officiels

.. moduleauthor: Cedric Airaud <cairaud@gmail.com>
"""

import argparse
import datetime
import pandas as pd
import numpy as np
import xlrd.biffh
import zipfile
import tempfile
import xml.etree.ElementTree
import copy
import os.path
import re

import logging
logging.basicConfig(level=logging.INFO, format="%(levelname)-9s %(lineno)-4s %(message)s")


class OfficielException(Exception):
    pass


class Club:
    def __init__(self, index, nom, departement):
        self.index = index
        self.nom = nom
        self.departement = departement
        self.competitions = []
        self.officiels = {}

    def __str__(self):
        return "{} ({})".format(self.nom, self.departement)

    def link(self):
        return "Club{}".format(self.index)

    def add_officiel(self, officiel, reunion, poste):
        # officiel is a copy
        for o in self.officiels.keys():
            if o.index == officiel.index:
                officiel = o

        if officiel not in self.officiels:
            self.officiels[officiel] = {}
        self.officiels[officiel][reunion] = poste.nom

    def departement_name(self):
        return "Département {}".format(self.departement)


class Niveau:
    def __init__(self, index, nom, valeur):
        self.index = index
        self.nom = nom
        self.valeur = valeur

    def __lt__(self, other):
        return self.valeur < other.valeur

    def __eq__(self, other):
        return self.valeur == other.valeur

    def __le__(self, other):
        return self.valeur <= other.valeur

    def __str__(self):
        return self.nom


class Poste:
    def __init__(self, index, nom, niveau, depart, regional):
        self.index = index
        self.nom = nom
        self.niveau = niveau

        # Empty, Licencié or Officiel: when does it count?
        self.depart = depart
        self.regional = regional

    def __str__(self):
        return "{}".format(self.nom)

    def valid_for(self, officiel):
        """
        Indicates if this poste is valid for an officiel, for depart and regional levels
        :param officiel:
        :return: (depart, regional)
        """
        return (self.depart == "Licencié" or self.depart == "Officiel" and officiel.real_officiel,
                self.regional == "Licencié" or self.regional == "Officiel" and officiel.real_officiel)

    def preferred_to(self, other):
        """
        Return True if this post should be preferred to the other one
        :param other: Other poste to look at
        :return: bool
        """
        scores = {"Licencié": 2, "Officiel": 1}
        score_self = self.niveau.valeur + scores.get(self.depart, 0) + scores.get(self.regional, 0)
        score_other = other.niveau.valeur + scores.get(other.depart, 0) + scores.get(other.regional, 0)
        logging.debug("{}: {}, {}: {}".format(str(self), score_self, str(other), score_other))
        if score_self != score_other:
            return score_self > score_other
        else:
            return self.index < other.index


class Officiel:
    """
    Represent an Officiel
    """
    def __init__(self, index, nom, prenom, club, niveau, niveau_c):
        self.nom = nom
        self.prenom = prenom
        self.club = club
        self.index = index
        self.niveau = copy.deepcopy(niveau)
        self.poste = None
        self.real_officiel = niveau_c <= self.niveau
        self.valid = None

    def set_poste(self, poste, reunion):
        """
        Set the poste. If required level is more than officiel level, change the level
        Add officiel/poste to the list in the corresponding club
        """
        if self.poste is None:
            self.poste = poste
        else:
            if self.poste.preferred_to(poste):
                logging.debug("Pour {}, le poste {} est préféré à {}".format(str(self), str(self.poste), str(poste)))
                return
            else:
                logging.debug("Pour {}, le poste {} est préféré à {}".format(str(self), str(poste), str(self.poste)))
                self.poste = poste

        if self.niveau < poste.niveau:
            logging.warning("{}: le poste {} requiert un niveau {}".format(str(self), str(poste), str(poste.niveau)))
        self.valid = self.poste.valid_for(self)

        self.club.add_officiel(self, reunion, poste)

    def is_valid(self, depart):
        """
        Return True if officiel at given post is valid
        :param depart: True if Departemental, False if Regional or more
        :return: bool
        """
        return self.valid[0 if depart else 1]

    def get_level(self):
        return self.niveau

    def __str__(self):
        return "{} {} ({} {})".format(self.prenom, self.nom, str(self.niveau), self.club.nom)


class Configuration:
    """
    Read configuration from the given filename and stores it
    """
    def __init__(self, filename):
        """
        Read configuration from file

        :param filename: Name of the configuration file
        :type filename: String
        """
        self.clubs = {}
        self.postes = {}
        self.niveaux = {}
        self.type_competitions = {}
        self.niveau_competitions = {}
        self.engagements = {}
        self.nages = {}
        self.disqualifications = {}
        self.epreuves = {}
        self.club_override = {}
        self.reunions = []
        self.filename = filename
        self.grille = {}
        self.eleves = {}

        logging.info("Lecture du fichier de configuration")

        for index, row in self.read_sheet("Clubs", ["Club", "Département"], 0).iterrows():
            self.clubs[index] = Club(index=index, nom=row["Club"], departement="{:02d}".format(row["Département"]))
            logging.debug("Club {}: {}".format(index, str(self.clubs[index])))

        for index, row in self.read_sheet("Niveaux", ["Niveau", "Valeur"], 0).iterrows():
            if np.isnan(index):
                continue
            self.niveaux[index] = Niveau(index, row["Niveau"], row["Valeur"])
            logging.debug("Niveau {}: {}".format(index, self.niveaux[index]))
            if row["Niveau"] == "C":
                self.niveau_c = self.niveaux[index]
            if row["Niveau"] == "B":
                self.niveau_b = self.niveaux[index]
            if row["Niveau"] == "Elève Chrono":
                self.niveau_c_next = self.niveaux[index]

        for index, row in self.read_sheet("Postes", ["Poste", "Niveau", "Départemental", "Régional"], 0).iterrows():
            niveau = min(self.niveaux.values())
            n = row["Niveau"] if not isinstance(row["Niveau"], float) else ""
            if n != "":
                niveau_l = [item for item in self.niveaux.values() if item.nom == n]
                if len(niveau_l) != 1:
                    raise OfficielException("Le niveau {} pour le poste {} n'est pas correct"
                                            .format(n, row["Poste"]))
                else:
                    niveau = niveau_l[0]

            self.postes[index] = Poste(index=index, nom=row["Poste"], niveau=niveau, depart=row["Départemental"],
                                       regional=row["Régional"])
            logging.debug("Poste {}: {}".format(index, str(self.postes[index])))

        for index, row in self.read_sheet("Epreuves", ["Nom"], 0).iterrows():
            self.epreuves[index] = row["Nom"]

        for index, row in self.read_sheet("Niveau compétitions", ["Niveau"], 0).iterrows():
            self.niveau_competitions[index] = row["Niveau"]
            self.engagements[row["Niveau"]] = {"Individuels": row["Individuels"],
                                               "Relais": row["Relais"],
                                               "Equipes": row["Equipes"]}

        for index, row in self.read_sheet("Types compétitions", ["Description", "Niveau"], 0).iterrows():
            niveau = int(row["Niveau"])
            if niveau not in self.niveau_competitions:
                logging.error("Pour la feuille 'Types compétition', ligne '{}', le niveau {} n'existe pas"
                              .format(row["Description"], niveau))
            self.type_competitions[index] = (row["Description"], self.niveau_competitions[niveau])

        for index, row in self.read_sheet("Changement Club", ["Nom", "Prénom", "Club"], 0).iterrows():
            if int(row["Club"]) not in self.clubs:
                logging.fatal("Le club {} n'existe pas pour forcer un club à {} {}"
                              .format(row["Club"], row["Prénom"], row["Nom"]))
            club = self.clubs[int(row["Club"])]
            self.club_override[index] = {"Club": club, "Nom": row["Nom"], "Prénom": row["Prénom"]}
            logging.warning("Club {} forcé pour {} {} ({})".format(club.nom, index, row["Prénom"], row["Nom"]))

        for index, row in self.read_sheet("Elèves", ["Nom", "Prénom", "Chrono"], 0).iterrows():
            if index is not None and not np.isnan(index):
                if not row["Chrono"] is pd.NaT:
                    self.eleves[int(index)] = {"Nom": row["Nom"], "Prénom": row["Prénom"], 
                                               "Chrono": row["Chrono"].to_pydatetime()}

        nages = ["Nage Libre", "Dos", "Brasse", "Papillon", "4 Nages"]
        for index, row in self.read_sheet("Nages", ["Nage"], 0).iterrows():
            nage = None
            for n in nages:
                if n.lower() in row["Nage"].lower():
                    nage = n
                    break

            if nage is None:
                logging.error("Nage non trouvée dans {}".format(row["Nage"]))

            if "messieurs" in row["Nage"].lower():
                sexe = "H"
            elif "dame" in row["Nage"].lower():
                sexe = "D"
            elif "mixte" in row["Nage"].lower():
                sexe = "M"
            else:
                logging.error("Sexe non trouvé dans {}".format(row["Nage"]))
                sexe = None

            self.nages[index] = row["Nage"], nage, sexe

        r1 = re.compile(r"([DH]) (\d+)-(\d+)")
        r2 = re.compile(r"([DH]) (\d+)\+")
        year = datetime.date.today().year
        if datetime.date.today().month >= 9:
            year += 1

        nages_indexes = {row[0]: idx for idx, row in self.nages.items()}
        nages_sexe = {"D": " Dames", "H": " Messieurs"}
        self.grille_max = {"D": 18, "H": 19}
        for index, row in self.read_sheet("Grilles", ["D 14-15", "D 16-17", "D 18+", "H 15-16", "H 17-18",
                                                      "H 19+", "Tolérance"], 0).iterrows():
            for key in row.keys():
                if key == "Tolérance":
                    for value in nages_sexe.values():
                        nage_idx = nages_indexes[index + value]
                        self.grille[nage_idx]["Tolérance"] = row[key]
                    continue

                m = r1.match(key)
                if m is not None:
                    # Different index for H/F
                    nage_idx = nages_indexes[index + nages_sexe[m.group(1)]]
                    if nage_idx not in self.grille:
                        self.grille[nage_idx] = {}
                    for i in range(year - int(m.group(3)), year - int(m.group(2))+1):
                        self.grille[nage_idx][i] = row[key]
                    continue

                m = r2.match(key)
                if m is not None:
                    # Different index for H/F
                    nage_idx = nages_indexes[index + nages_sexe[m.group(1)]]
                    if nage_idx not in self.grille:
                        self.grille[nage_idx] = {}
                    self.grille[nage_idx][year - int(m.group(2))] = row[key]
                    self.grille[nage_idx]["Min"] = year - int(m.group(2))
                    continue

                else:
                    logging.fatal("Impossible de comprendre {} dans la page Grilles".format(key))

        r = re.compile(r"DSQr(\d+)")
        for index, row in self.read_sheet("Disqualifications", ["Code", "Libellé"], 0).iterrows():
            code = row["Code"]
            m = r.match(code)
            relayeur = None
            if m is not None:
                relayeur = int(m.group(1))
                code = r.sub("DSQ", code)

            self.disqualifications[index] = (code, row["Libellé"], relayeur)

    def read_sheet(self, sheet_name, columns, index_col=None):
        """
        Read sheet of given name in file and checks that the colums are as expected.

        :param sheet_name: Name of the sheet to read
        :type sheet_name: string
        :param columns: List of expected columns
        :type columns: List
        :param index_col: Column to use as index (None if None)
        :type index_col: int|None
        :return: Read table
        :rtype: DataFrame
        """
        try:
            sheet = pd.read_excel(self.filename, sheet_name=sheet_name, parse_dates=True, index_col=index_col, engine='openpyxl')
        except xlrd.biffh.XLRDError:
            raise OfficielException("Pas de feuille '{}' trouvée".format(sheet_name))

        sheet_columns = list(sheet.columns.values)
        if index_col is not None:
            sheet_columns.insert(0, "Index")
            columns.insert(0, "Index")

        if len(set(sheet_columns).intersection(columns)) != len(columns):
            raise OfficielException("Pour la feuille {}, les colonnes attendues sont:\n{}\nles colonnes trouvées "
                                    "sont:\n{}".format(sheet_name, ", ".join(columns), ", ".join(sheet_columns)))

        return sheet


class Competition:
    """
    Represent a competition, composed of several Reunions
    """
    def __init__(self, conf, filename):
        """
        Read a competition from a FFNEX file
        """
        self.conf = conf
        self.filename = filename
        self.reunions = []
        self.competition_link = None  # Link from this competition to another one
        self.linked = []  # List of competitions linked to it
        self.depassements = []

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
        if e.tag != "FFNEX":
            raise OfficielException("Le fichier {} n'est pas compatible: FFNEX attendu, {} trouvé"
                                    .format(self.filename, e.tag))
        if e.attrib["version"] != "1.1.1":
            logging.warning("Le fichier {} utilise la version {} alors que le script attend la version 1.1.1"
                            .format(self.filename, e.attrib["version"]))

        # Competition
        competition = e.find("MEETS").find("MEET")
        self.id = int(competition.attrib["id"])
        self.nom = competition.attrib["name"]
        self.startdate = datetime.datetime.strptime(competition.attrib["startdate"], "%Y-%m-%d")
        self.stopdate = datetime.datetime.strptime(competition.attrib["stopdate"], "%Y-%m-%d")
        self.ville = competition.attrib["city"]
        self.par_equipe = True if competition.attrib.get("byteam", "false") == "true" else 1
        self.type, self.niveau = conf.type_competitions[int(competition.attrib["typeid"])]
        self.clubs = []
        pool = competition.find("POOL")
        self.lanes = int(pool.attrib["lanes"])
        self.length = int(pool.attrib["size"])

        logging.info("Lecture de la compétition {} ({}) - {} à {} - {}".format(self.nom, self.id,
                                                                               self.date_str(), self.ville,
                                                                               self.niveau))

        # Competition can be linked to another one: jury is ignored but number of participations are added
        link = competition.find("LINK")
        if link is not None:
            self.competition_link = int(link.attrib["rel"])
            logging.info("Compétition liée à la compétition {}".format(self.competition_link))

        # Check list of clubs
        for o in competition.find("CLUBS").findall("CLUB"):
            code, clubid, name = o.attrib["code"], int(o.attrib["id"]), o.attrib["name"]
            club = self.conf.clubs.get(clubid, None)
            if club is not None:
                continue
            departement = code[4:6]
            departements = set([c.departement for c in self.conf.clubs.values()])
            if departement in departements:
                logging.error("Le club {} ({}) n'est pas dans la liste:\n{};{};{}".format(name, code, clubid, name,
                                                                                          departement))
            else:
                logging.debug("Le club {} n'est pas dans la région ({}: {})".format(name, departement, code))

        # List of officials
        officiels = {}
        for o in competition.find("OFFICIALS").findall("OFFICIAL"):
            index, clubid, gradeid = int(o.attrib["id"]), int(o.attrib["clubid"]), int(o.attrib["gradeid"])
            if index in self.conf.club_override:
                d = self.conf.club_override[index]
                club, nom, prenom = d["Club"], d["Nom"], d["Prénom"]
                if nom != o.attrib["lastname"] or prenom != o.attrib["firstname"]:
                    logging.fatal("Le nom/prénom ne correspond pas pour l'ID {}: {} {} vs. {} {}"
                                  .format(index, nom, prenom, o.attrib["lastname"], o.attrib["firstname"]))
                else:
                    logging.warning("Club {} forcé pour {} {} ({})".format(club.nom, prenom, nom, index))
            else:
                club = self.conf.clubs.get(clubid, None)
            try:
                niveau = self.conf.niveaux[gradeid]
            except KeyError:
                logging.fatal(f"Le niveau {gradeid} pour l'officiel {prenom} {nom} n'est pas listé dans le fichier de configuration")
            if club is not None and club.departement != '99':
                officiels[index] = Officiel(index=index, nom=o.attrib["lastname"], prenom=o.attrib["firstname"],
                                            club=club, niveau=niveau, niveau_c=conf.niveau_c)
                logging.debug("Officiel trouvé: {}".format(str(officiels[index])))
                if club not in self.clubs:
                    self.clubs.append(club)
            else:
                logging.debug("Officiel ignoré: {} {} ({})".format(o.attrib["firstname"], o.attrib["lastname"], clubid))

        # List of clubs declared as banniere
        for o in competition.find("CLUBS").findall("CLUB"):
            index, clubid, name = int(o.attrib["id"]), o.attrib.get("clubid", None), o.attrib["name"]
            if index < 0:
                if clubid is None:
                    logging.info("Bannière trouvée: {}. Rajouter clubid='<id>' s'il représente un club".format(name))
                else:
                    clubid = int(clubid)
                    logging.info("Club déclaré en bannière: {} ({} -> {})".format(name, index, clubid))
                    club = self.conf.clubs.get(clubid, None)
                    if club is None:
                        logging.fatal("Ce club est invalide")
                    if club.nom != name:
                        logging.warning("Le nom ne correspond pas: '{}' vs '{}'".format(name, club.nom))
                    self.conf.clubs[index] = club

        # List of swimmers
        nageurs = {}
        nom_nageurs = {}
        nageurs_year = {}
        for n in competition.find("SWIMMERS").findall("SWIMMER"):
            if "clubid" in n.attrib:
                index, clubid = int(n.attrib["id"]), int(n.attrib["clubid"])
                club = self.conf.clubs.get(clubid, None)
                nageurs[index] = club
                nom_nageurs[index] = n.attrib["firstname"] + " " + n.attrib["lastname"]
                nageurs_year[index] = datetime.datetime.strptime(n.attrib["birthdate"], "%Y-%m-%d").year
                if club is not None and club.departement != '99' and club not in self.clubs:
                    self.clubs.append(club)
            else:
                logging.warning("Le nageur {} {} ({}) est ignoré (Pas de clubid)".format(n.attrib["firstname"],
                                n.attrib["lastname"], n.attrib["nation"]))

        # List of sessions
        def race_id(item):
            return "{}_{}".format(item.attrib["raceid"], item.attrib["roundid"])

        races = {}
        finals = {}
        for session in competition.find("SESSIONS").findall("SESSION"):
            # List of races, with an index to the reunion
            reunion = Reunion(int(session.attrib["number"]), self)
            race_found = False
            for event in session.find("EVENTS").findall("EVENT"):
                if event.attrib["type"] == "RACE":
                    race_found = True
                    races[race_id(event)] = reunion
                    finals[race_id(event)] = "Final" in conf.epreuves[int(event.attrib["roundid"])]

            reunion.participations = {club: 0 for club in self.clubs}
            reunion.participants = {club: [] for club in self.clubs}
            reunion.engagements = {club: 0 for club in self.clubs}
            reunion.financier = {club: dict(individuel=0, relais=0, equipe=0) for club in self.clubs}
            for key in reunion.forfaits:
                reunion.forfaits[key] = {club: 0 for club in self.clubs}

            if race_found:
                self.reunions.append(reunion)
                reunion_start = datetime.datetime.strptime(session.attrib["datetime"], "%Y-%m-%d %H:%M:%S")
                for judge in session.find("JUDGES").findall("JUDGE"):
                    officielid, roleid = int(judge.attrib["officialid"]), int(judge.attrib["roleid"])
                    poste = conf.postes.get(roleid, None)
                    officiel = officiels.get(officielid, None)

                    if poste is None:
                        logging.error("Officiel {}: poste {} non trouvé".format(str(officiel), roleid))
                    if officiel is None:
                        logging.warning("Officiel ID {} (role {}) non trouvé".format(officielid, roleid))
                    else:
                        logging.debug("{}: {}".format(str(officiel), str(poste)))
                        officiel = copy.copy(officiel)

                        if officiel.niveau < self.conf.niveau_c and officiel.index in self.conf.eleves:
                            eleve = self.conf.eleves[officiel.index]
                            if eleve["Chrono"] < reunion_start:
                                logging.info(f"Officiel {officiel.prenom} {officiel.nom} passé en élève chrono")
                                officiel.niveau = self.conf.niveau_c_next
                                officiel.real_officiel = True

                        if officielid in reunion.officiels:
                            reunion.officiels[officielid].set_poste(poste, reunion)
                        else:
                            officiel.set_poste(poste, reunion)
                            reunion.officiels[officielid] = officiel

            else:
                logging.debug("Session {} ignorée: pas suffisamment d'events".format(session.attrib["number"]))

        # Size of teams
        if self.par_equipe is True:
            for result in competition.find("RESULTS").findall("RESULT"):
                relay = result.find("RELAY")
                if relay and result.attrib["disqualificationid"] == "0" and relay.find("RELAYPOSITIONS") is not None:
                    self.par_equipe = len(list(relay.find("RELAYPOSITIONS").findall("RELAYPOSITION")))
                    break

            if self.par_equipe == 1:
                logging.error("Taille d'équipe non trouvée")

        # Swimmers
        for result in competition.find("RESULTS").findall("RESULT"):
            reunion = races[race_id(result)]
            is_final = finals[race_id(result)]

            for record in list(result):
                if self.par_equipe != 1:
                    club = self.conf.clubs.get(int(result.attrib["clubid"]), None)
                    team = int(result.attrib["team"])
                    sexe = conf.nages[int(result.attrib["raceid"])][2]
                    if club is not None and not is_final:
                        reunion.participants[club].append("{} {}".format(team, sexe))
                        disqualification = int(result.attrib["disqualificationid"])
                        if disqualification in reunion.forfaits:
                            reunion.forfaits[disqualification][club] += 1

                elif record.tag == "SOLO":
                    nageurid = int(record.attrib["swimmerid"])
                    # club = nageurs[nageurid] Bug: declaration of swimmers does not contain correct club
                    club = self.conf.clubs.get(int(result.attrib["clubid"]), None)
                    if club is not None and club in self.clubs:
                        if club not in reunion.participants:
                            logging.error("Club {} not in participants list".format(str(club)))
                        reunion.participants[club].append(nageurid)
                        reunion.engagements[club] += 1
                        if not is_final:
                            reunion.financier[club]["individuel"] += 1
                        disqualification = int(result.attrib["disqualificationid"])
                        if disqualification in reunion.forfaits:
                            reunion.forfaits[disqualification][club] += 1

                        # Dépassement par rapport à la grille
                        #nage = int(result.attrib["raceid"])
                        #if nage in conf.grille and disqualification == 0:
                        #    year = nageurs_year[nageurid]
                        #    nage = conf.grille[nage]
                        #    if year < nage['Min']:
                        #        year = nage['Min']
                        #    minutes, secondes = [int(x) for x in result.attrib["swimtime"].split(".")]
                        #    temps = datetime.time(hour=minutes // 60, minute=minutes % 60, second=secondes // 100,
                        #                          microsecond=(secondes % 100) * 10)
                        #    tolerance = datetime.timedelta(seconds=nage["Tolérance"])
                        #    temps_tolerance = ((datetime.datetime.combine(datetime.date.today(),
                        #                                                  nage[year]) + tolerance).time())

                        #    if temps > nage[year]:
                        #        depassement = temps > temps_tolerance
                        #        self.depassements.append({"Club": club, "Nageur": nom_nageurs[nageurid],
                        #                                  "Temps": temps, "Limite": nage[year],
                        #                                  "Temps avec tolérance": temps_tolerance,
                        #                                  "Dépassement": "Oui" if depassement else "Non"})
                        #        logging.info("Dépassement pour {} ({}): {} (Limite: {}, avec tolérance {}".format(
                        #            nom_nageurs[nageurid], str(club), str(temps), str(nage[year]),
                        #            str(temps_tolerance)))

                elif record.tag == "RELAY":
                    positions = record.find("RELAYPOSITIONS")
                    if positions:
                        club = None
                        for relay_position in positions:
                            nageurid = int(relay_position.attrib["swimmerid"])
                            # club = nageurs[nageurid] Bug: declaration of swimmers does not contain correct club
                            club = self.conf.clubs.get(int(result.attrib["clubid"]), None)
                            if club is not None and club in self.clubs:
                                reunion.participants[club].append(nageurid)
                                reunion.engagements[club] += 1
                        if club is not None and club in reunion.financier and not is_final:
                            reunion.financier[club]["relais"] += 1
                            disqualification = int(result.attrib["disqualificationid"])
                            if disqualification in reunion.forfaits:
                                reunion.forfaits[disqualification][club] += 1

                elif record.tag == "SPLIT":
                    pass

        # Counts number of participations per club
        for reunion_num, reunion in enumerate(self.reunions):
            for club, l in reunion.participants.items():
                reunion.participations[club] = len(set(l))
                if self.par_equipe != 1:
                    reunion.engagements[club] = reunion.participations[club] * self.par_equipe
                    if reunion_num == 0:  # Only first reunion when by team
                        reunion.financier[club]["equipe"] += reunion.participations[club]

        # Update list of competitions for each club
        for club in self.clubs:
            club.competitions.append(self)

    def titre(self):
        """
        Return the Title as a string
        :return: Title
        :rtype: string
        """
        return "{} - {}".format(self.nom, self.ville)

    def date_str(self):
        """
        Date as a string. Either a single date or start - stop
        """
        if self.startdate == self.stopdate:
            return self.startdate.strftime("%d/%m/%Y")
        else:
            return "{} au {}".format(self.startdate.strftime("%d/%m/%Y"), self.stopdate.strftime("%d/%m/%Y"))

    def departemental(self):
        """
        Return true if competition is of Departement level
        """
        return "département" in self.niveau.lower()

    def __str__(self):
        return ("{}\n{}: {}\n\n".format(self.nom, self.ville, self.date_str()) +
                "\n\n".join(map(str, self.reunions)))

    def link(self):
        return "C{}".format(self.id)

    def weblink(self):
        return "http://ffn.extranat.fr/webffn/resultats.php?idact=nat&idcpt={}".format(self.id)


class Reunion:
    """
    Represent a Reunion, base for the calculation
    """
    # Forfaits
    DECL = 2      # Déclaré
    NON_DECL = 4  # Non-déclaré
    CERT = 1      # Certificat médical

    def __init__(self, index, competition):
        self.index = index
        self.competition = competition
        self.titre = "Réunion N°{}".format(index)
        self.officiels = {}
        self.participants = None
        self.participations = None
        self.forfaits = {self.DECL: 0, self.NON_DECL: 0, self.CERT: 0}
        self.engagements = None
        self.financier = None
        self._officiels_per_club = None
        self.pts = {}
        self.details = {}

    def __str__(self):
        return self.titre + "\n  " + "\n  ".join(map(str, self.officiels.values()))

    def officiels_per_club(self):
        """
        Sort officiels per club
        """
        if self._officiels_per_club is not None:
            return self._officiels_per_club

        self._officiels_per_club = {}
        for officiel in self.officiels.values():
            if officiel.club not in self._officiels_per_club:
                self._officiels_per_club[officiel.club] = []
            self._officiels_per_club[officiel.club].append(officiel)

        return self._officiels_per_club

    def points(self, club, details=None):
        """
        :param club: Club to look for
        :type club: Club
        :param details: Optional list to get detail about calculation
        :type details: None|List
        :return: Number of points
        :rtype: int
        """
        if club in self.pts and (club in self.details or details is None):
            if details is not None:
                details += self.details[club]
            return self.pts[club]

        participations = self.participations.get(club, 0)

        # needed = (Num of A/B, Total num)
        if self.competition.departemental():
            num_equipes = participations
            participations *= self.competition.par_equipe
            if participations == 0:
                needed = (0, 0)
            elif self.competition.par_equipe == 10:
                # Special case of equipes by 10 (Interclubs TC)
                num_officiels = num_equipes
                needed = (num_officiels // 2, num_officiels)
            else:
                num_officiels = (participations + 7) // 8
                needed = (num_officiels // 2, num_officiels)

        elif self.competition.par_equipe != 1:
            if participations <= 1:
                needed = (0, participations)
            else:
                needed = (1, min(3, participations))

        else:
            if participations <= 1:
                needed = (0, 0)
            elif participations <= 10:
                needed = (0, 1)
            else:
                needed = (1, 2)

        if type(details) is list:
            s = "{} officiels requis".format(needed[1])
            if needed[0] > 0:
                s += ", dont {} A ou B".format(needed[0])
            details.append(s)

        num_ab, num = 0, 0
        club_officiels = self.officiels_per_club().get(club, [])
        for officiel in club_officiels:
            if not officiel.is_valid(self.competition.departemental()):
                logging.warning("Le licencié/officiel {} n'est pas pas pris en compte au poste {}".format(
                                str(officiel), str(officiel.poste)))
                continue
            num += 1
            if conf.niveau_b <= officiel.get_level():
                num_ab += 1

        if not self.competition.departemental() and num > 5:
            if type(details) is list:
                details.append("5 officiels retenus sur les {} présentés".format(num))
            num = 5

        if num < needed[1]:
            missing = needed[1] - num
            pts = missing * -4
            if type(details) is list:
                details.append("{} points négatifs pour {} officiels manquants".format(-pts, missing))
        else:
            extra = num - needed[1]
            pts = extra * 2
            if extra > 0 and type(details) is list:
                details.append("{} points supplémentaires pour {} officiels".format(pts, extra))
            if num_ab < needed[0]:
                missing = needed[0] - num_ab
                pts += missing * -2
                if type(details) is list:
                    details.append("{} points de malus par manque d'officiel A/B".format(missing*2))

        if details is not None:
            self.details[club] = details
        self.pts[club] = pts
        return pts

    def link(self):
        competition_id = self.competition.id
        if self.competition.competition_link is not None:
            competition_id = self.competition.competition_link.id
        return "C{}_R{}".format(competition_id, self.index)


if __name__ == "__main__":
    import gen_pdf

    parser = argparse.ArgumentParser(description='Liste des compétitions')
    parser.add_argument("--conf", default="Officiels.xlsx", help="Fichier de configuration")
    parser.add_argument("--format", default=False, action="store_true",
                        help="Rend un fichier FFNEX plus lisible à l'intérieur du fichier Zip)")
    parser.add_argument("--competition", default=None, help="Génération pour cette compétition seulement. " +
                                                            "Pas de résumé des clubs.")
    parser.add_argument("--output", default="Compétitions.pdf", help="Fichier PDF de sortie")
    parser.add_argument("ffnex_files", metavar="fichiers", nargs="+", help="Liste des fichiers ou répertoires " +
                                                                           "à analyser")

    args = parser.parse_args()

    competitions = []
    conf = Configuration(args.conf)

    files = []
    for f in args.ffnex_files:
        if os.path.isdir(f):
            files += [os.path.join(f, file) for file in os.listdir(f) if os.path.isfile(os.path.join(f, file))]
        else:
            files.append(f)

    ffnex_files = [f for f in files if os.path.splitext(f)[1] == ".xml"]
    for f in [f for f in files if os.path.splitext(f)[1] == ".zip"]:
        if os.path.splitext(f)[0] + ".xml" in ffnex_files:
            logging.info("Fichier {} ignoré car déjà présent en .xml".format(f))
        else:
            ffnex_files.append(f)

    ffnex_files = [f for f in ffnex_files if os.path.splitext(f)[1] in (".zip", ".xml")]
    ffnex_files.sort()

    if args.format:
        import xml.dom.minidom
        for f in ffnex_files:
            logging.info("Extraction du fichier {}".format(f))
            backup_file = f + ".bak"
            filename = None

            try:
                z = zipfile.ZipFile(f)
                if "ffnex.xml" not in z.namelist():
                    logging.error("Le fichier {} devrait contenir un fichier ffnex.xml")
                else:
                    filename = z.extract('ffnex.xml', tempfile.gettempdir())

                    data = xml.dom.minidom.parse(filename)
                    z.close()
                    with open(filename, "w") as of:
                        of.write(data.toprettyxml())

            except zipfile.BadZipfile as e:
                logging.error("Le fichier {} ne peut pas être lu correctement:\n{}".format(filename, str(e)))
                continue

            logging.info("Fichier de sauvegarde {}".format(backup_file))
            os.rename(f, backup_file)

            logging.info("Recompression du fichier de sortie {}".format(f))
            with zipfile.ZipFile(f, "w") as z:
                z.write(filename, "ffnex.xml")

        exit(0)

    if args.competition is not None and args.competition not in ffnex_files:
        parser.error("Le fichier {} n'est pas dans la liste des fichiers source pour le paramètre 'competition'"
                     .format(args.competition))
        exit(-1)

    for f in ffnex_files:
        competition = Competition(conf, f)
        competitions.append(competition)
        if args.competition == f:
            logging.debug("Specific competition found: " + f)
            args.competition = competition

    competitions_ids = sorted([competition.id for competition in competitions])
    if len(competitions) != len(set(competitions_ids)):
        duplicates = []
        for i in range(1, len(competitions_ids)):
            if competitions_ids[i] == competitions_ids[i-1]:
                duplicates.append(competitions_ids[i])
        logging.fatal("Des compétitions sont dupliquées: {}".format(", ".join(map(str, duplicates))))
        exit(-1)

    competitions_by_id = {competition.id: competition for competition in competitions}
    if args.competition is not None:
        competitions = [args.competition]

    # Create links: linked competitions are removed from the list
    link_list = [competition for competition in competitions if competition.competition_link is not None]
    for competition in link_list:
        master = competitions_by_id[competition.competition_link]
        master.linked.append(competition)
        competition.competition_link = master

        # Add participations
        if len(competition.reunions) != len(master.reunions):
            logging.fatal("La compétition {} est liée à la compétition {} mais elles n'ont pas le même nombre de"
                          "réunions".format(competition.id, master.id))
            exit(-1)

        for i, creunion in enumerate(competition.reunions):
            mreunion = master.reunions[i]

            for club in set(list(creunion.participations.keys()) + list(mreunion.participations.keys())):
                if master not in club.competitions:
                    club.competitions.append(master)
                mreunion.participations[club] = (mreunion.participations.get(club, 0) +
                                                 creunion.participations.get(club, 0))

            for club in set(list(creunion.engagements.keys()) + list(mreunion.engagements.keys())):
                if master not in club.competitions:
                    club.competitions.append(master)
                mreunion.engagements[club] = mreunion.engagements.get(club, 0) + creunion.engagements.get(club, 0)

            for officielid, officiel in creunion.officiels.items():
                if officielid not in mreunion.officiels:
                    mreunion.officiels[officielid] = officiel

    departements = set([c.departement_name() for c in conf.clubs.values()])
    points = {"Régional": {"participations": 0, "engagements": 0, "total_bonus": 0}}
    for d in list(departements):
        points[d] = {"participations": 0, "engagements": 0, "total_bonus": 0}

    raw_df = []

    logging.info("Génération du fichier PDF {}".format(args.output))
    doc = gen_pdf.DocTemplate(conf, args.output, "Liste des compétitions", "Cédric Airaud")
    for competition in competitions:

        for reunion in competition.reunions:
            for club in reunion.participations.keys():
                pts = reunion.points(club, details=[])
                participations = reunion.participations.get(club, 0)
                engagements = reunion.engagements.get(club, 0)
                officiels = reunion.officiels_per_club().get(club, [])
                num_officiels = len([o for o in officiels if o.is_valid(competition.departemental())])
                officiels_per_categorie = {}
                for officiel in officiels:
                    if officiel.is_valid(competition.departemental()):
                        niveau = officiel.niveau.nom
                        if niveau not in officiels_per_categorie:
                            officiels_per_categorie[niveau] = 0
                        officiels_per_categorie[niveau] += 1

                if competition.competition_link:
                    pts, num_officiels, engagements, participations = 0, 0, 0, 0
                    officiels_per_categorie = {}

                if competition.departemental():
                    niveau = club.departement_name()
                else:
                    niveau = "Régional"
                l = points[niveau]
                if club not in l:
                    l[club] = 0

                l["participations"] += participations
                l["engagements"] += engagements
                l[club] += pts

                raw = {"Niveau": competition.niveau, "Structure": niveau,
                       "Par Equipe": competition.par_equipe != 1,
                       "Compétition": competition.titre(), "Date": competition.startdate,
                       "Réunion": reunion.index, "Club": club.nom,
                       "Participations": participations,
                       "Engagements": engagements,
                       "Points": pts, "Officiels": num_officiels,
                       "Individuels": reunion.financier.get(club, {}).get("individuel", 0),
                       "Relais": reunion.financier.get(club, {}).get("relais", 0),
                       "Equipes": reunion.financier.get(club, {}).get("equipe", 0),
                       "Lignes": competition.lanes, "Longueur": competition.length,
                       "Disq-Médical": reunion.forfaits[reunion.CERT].get(club, 0),
                       "Disq-Déclaré": reunion.forfaits[reunion.DECL].get(club, 0),
                       "Disq-NonDéclaré": reunion.forfaits[reunion.NON_DECL].get(club, 0),
                       }

                for key, value in officiels_per_categorie.items():
                    raw["Officiels-" + key] = value

                raw_df.append(raw)

    if args.competition is None:
        raw_df = pd.DataFrame(raw_df)

        def total_engagements(row):
            prices = conf.engagements[row["Niveau"]]
            return (row["Individuels"] * prices["Individuels"] + row["Relais"] * prices["Relais"] +
                    row["Equipes"] * prices["Equipes"])

        raw_df["Total"] = raw_df.apply(total_engagements, axis=1)

        writer = pd.ExcelWriter("export.xlsx")

        points_df = raw_df[['Date', 'Niveau', 'Structure', 'Compétition', 'Réunion', 'Club',
                            'Par Equipe', 'Participations', 'Engagements', 'Officiels', 'Points']]
        points_df.to_excel(writer, sheet_name="Points")

        columns = (['Participations', 'Engagements', 'Officiels', 'Points'] +
                   [c for c in raw_df.columns if "Officiels-" in c])
        officiels_df = raw_df.groupby(['Structure', 'Club'])[columns].sum()
        officiels_df.to_excel(writer, sheet_name="Officiels par compétition")

        etat_df = raw_df.groupby(['Structure', 'Club'])['Individuels', 'Relais', 'Equipes', 'Total', 'Points',
                                                        'Disq-Médical', 'Disq-Déclaré', 'Disq-NonDéclaré'].sum()
        etat_df.rename(columns={'Points': 'Points Bonus/Malus'}, inplace=True)
        etat_df.to_excel(writer, sheet_name="Etat financier")

        def first_item(x):
            return x.iloc[0]

        competitions_df = raw_df.groupby(['Compétition', 'Réunion'])
        competitions_df = competitions_df.agg({'Participations': np.sum,
                                               'Engagements': np.sum,
                                               'Officiels': np.sum,
                                               'Lignes': first_item,
                                               'Longueur': first_item,
                                               'Niveau': first_item,
                                               'Par Equipe': first_item})
        competitions_df['Officiels voulus'] = competitions_df['Lignes'] * 3 + 9
        competitions_df = competitions_df[["Niveau", "Participations", "Engagements", "Officiels", "Lignes", "Longueur",
                                           "Officiels voulus"]]
        competitions_df.to_excel(writer, sheet_name="Compétitions")

        bonus_df = officiels_df[officiels_df["Points"] > 0].reset_index()
        for key, l in points.items():
            l["total_bonus"] = bonus_df[bonus_df["Structure"] == key]["Points"].sum()

        doc.bonus = {level: 0.50 * l["engagements"] / l["total_bonus"] if l["total_bonus"] else 0
                     for level, l in points.items()}
        for level, value in doc.bonus.items():
            logging.info("Valeur du point bonus pour {}: {:.2f} € (Total engagements: {}, total bonus: {})"
                         .format(level, value, points[level]["engagements"], points[level]["total_bonus"]))

        officiels_df = []
        for club in sorted(conf.clubs.values(), key=lambda x: "{} {}".format(x.departement, x.nom)):
            for officiel, reunions in club.officiels.items():
                for reunion, poste in reunions.items():
                    officiels_df.append({"Officiel": "{} {}".format(officiel.nom, officiel.prenom),
                                         "Club": club.nom,
                                         "Date": reunion.competition.startdate,
                                         "Compétition": reunion.competition.titre(),
                                         "Réunion": reunion.index,
                                         "Poste": poste,
                                         "Niveau": str(officiel.niveau)})

        officiels_df = pd.DataFrame(officiels_df)
        officiels_df.to_excel(writer, sheet_name="Officiels")

        raw_df.to_excel(writer, sheet_name="Données brutes")
        writer.save()

    if args.competition is None:
        for club in sorted(conf.clubs.values(), key=lambda x: "{} {}".format(x.departement, x.nom)):
            doc.new_club(club)

    for competition in sorted(competitions, key=lambda x: x.startdate):
        if competition.competition_link is None:
            doc.new_competition(competition)

    doc.build()
