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

    def __str__(self):
        return "{} ({})".format(self.nom, self.departement)

    def link(self):
        return "Club{}".format(self.index)


class Niveau:
    def __init__(self, index, nom, valeur):
        self.index = index
        self.nom = nom
        self.valeur = valeur

    def __lt__(self, other):
        return self.valeur < other.valeur

    def __eq__(self, other):
        return self.valeur == other.valeur

    def __str__(self):
        return self.nom


class Poste:
    def __init__(self, index, nom, niveau):
        self.index = index
        self.nom = nom
        self.niveau = niveau

    def __str__(self):
        return "{}".format(self.nom)


class Officiel:
    """
    Represent an Officiel
    """
    def __init__(self, index, nom, prenom, club, niveau):
        self.nom = nom
        self.prenom = prenom
        self.club = club
        self.index = index
        self.niveau = copy.deepcopy(niveau)
        self.poste = None
        self.valid = niveau.valeur > 0  # 0 = Seulement licencié

    def set_poste(self, poste):
        """
        Set the poste. If required level is more than officiel level, change the level
        """
        self.poste = poste
        if self.valid and poste.niveau > self.niveau:
            logging.warning("{}: le poste {} requiert un niveau {}".format(str(self), str(poste), str(poste.niveau)))

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
        self.nages = {}
        self.disqualifications = {}
        self.reunions = []
        self.filename = filename

        logging.info("Lecture du fichier de configuration")

        for index, row in self.read_sheet("Clubs", ["Club", "Département"], 0).iterrows():
            self.clubs[index] = Club(index=index, nom=row["Club"], departement=row["Département"])
            logging.debug("Club {}: {}".format(index, str(self.clubs[index])))

        for index, row in self.read_sheet("Niveaux", ["Niveau", "Valeur"], 0).iterrows():
            self.niveaux[index] = Niveau(index, row["Niveau"], row["Valeur"])
            logging.debug("Niveau {}: {}".format(index, str(self.niveaux[index])))
            if row["Niveau"] == "C":
                self.niveau_c = self.niveaux[index]

        for index, row in self.read_sheet("Postes", ["Poste", "Niveau"], 0).iterrows():
            niveau = min(self.niveaux.values())
            n = row["Niveau"] if not isinstance(row["Niveau"], float) else ""
            if n != "":
                l = [item for item in self.niveaux.values() if item.nom == n]
                if len(l) != 1:
                    raise OfficielException("Le niveau {} pour le poste {} n'est pas correct"
                                            .format(n, row["Poste"]))
                else:
                    niveau = l[0]

            self.postes[index] = Poste(index=index, nom=row["Poste"], niveau=niveau)
            logging.debug("Poste {}: {}".format(index, str(self.postes[index])))

        for index, row in self.read_sheet("Niveau compétitions", ["Niveau"], 0).iterrows():
            self.niveau_competitions[index] = row["Niveau"]

        for index, row in self.read_sheet("Types compétitions", ["Description", "Niveau"], 0).iterrows():
            niveau = int(row["Niveau"])
            if niveau not in self.niveau_competitions:
                logging.error("Pour la feuille 'Types compétition', ligne '{}', le niveau {} n'existe pas"
                              .format(row["Description"], niveau))
            self.type_competitions[index] = (row["Description"], self.niveau_competitions[niveau])

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

            self.nages[index] = row["Nage"], nage, sexe

        r = re.compile("DSQr(\d+)")
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
            sheet = pd.read_excel(self.filename, sheetname=sheet_name, convert_dates=True, index_col=index_col)
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
        if e.attrib["version"] != "1.1.0":
            logging.warning("Le fichier {} utilise la version {} alors que le script attend la version 1.1.0"
                            .format(self.filename, e.attrib["version"]))

        # Competition
        competition = e.find("MEETS").find("MEET")
        self.id = int(competition.attrib["id"])
        self.nom = competition.attrib["name"]
        self.startdate = datetime.datetime.strptime(competition.attrib["startdate"], "%Y-%m-%d")
        self.stopdate = datetime.datetime.strptime(competition.attrib["stopdate"], "%Y-%m-%d")
        self.ville = competition.attrib["city"]
        self.par_equipe = True if competition.attrib.get("byteam", "false") == "true" else 0
        self.type, self.niveau = conf.type_competitions[int(competition.attrib["typeid"])]
        self.clubs = []

        logging.info("Lecture de la compétition {} ({}) - {} à {} - {}".format(self.nom, self.id,
                     self.date_str(), self.ville, self.niveau))

        # List of officials
        officiels = {}
        for o in competition.find("OFFICIALS").findall("OFFICIAL"):
            index, clubid, gradeid = int(o.attrib["id"]), int(o.attrib["clubid"]), int(o.attrib["gradeid"])
            club = self.conf.clubs.get(clubid, None)
            niveau = self.conf.niveaux.get(gradeid, None)
            if club is not None:
                officiels[index] = Officiel(index=index, nom=o.attrib["lastname"], prenom=o.attrib["firstname"],
                                            club=club, niveau=niveau)
                logging.debug("Officiel trouvé: {}".format(str(officiels[index])))
                if club not in self.clubs:
                    self.clubs.append(club)
            else:
                logging.debug("Officiel ignoré: {} {} ({})".format(o.attrib["firstname"], o.attrib["lastname"], clubid))

        # List of swimmers
        nageurs = {}
        for n in competition.find("SWIMMERS").findall("SWIMMER"):
            index, clubid = int(n.attrib["id"]), int(n.attrib["clubid"])
            club = self.conf.clubs.get(clubid, None)
            nageurs[index] = club
            if club is not None and club not in self.clubs:
                self.clubs.append(club)

        # List of sessions
        races = {}
        for session in competition.find("SESSIONS").findall("SESSION"):
            # List of races, with an index to the reunion
            reunion = Reunion(int(session.attrib["number"]), self)
            race_found = False
            for event in session.find("EVENTS").findall("EVENT"):
                if event.attrib["type"] == "RACE":
                    race_found = True
                    races[event.attrib["raceid"]] = reunion

            reunion.participations = {club: 0 for club in self.clubs}
            reunion.participants = {club: [] for club in self.clubs}
            reunion.engagements = {club: 0 for club in self.clubs}

            if race_found:
                self.reunions.append(reunion)
                for judge in session.find("JUDGES").findall("JUDGE"):
                    officielid, roleid = int(judge.attrib["officialid"]), int(judge.attrib["roleid"])
                    poste = conf.postes.get(roleid, None)
                    officiel = officiels.get(officielid, None)
                    if poste is None:
                        logging.error("Officiel {}: poste {} non trouvé".format(str(officiel), roleid))
                    if officiel is not None:
                        logging.debug("{}: {}".format(str(officiel), str(poste)))

                        if officielid in reunion.officiels:
                            reunion.officiels[officielid].set_poste(poste)
                        else:
                            officiel = copy.copy(officiel)
                            officiel.set_poste(poste)
                            reunion.officiels[officielid] = officiel

            else:
                logging.debug("Session {} ignorée: pas suffisamment d'events".format(session.attrib["number"]))

        # Size of teams
        if self.par_equipe is True:
            for result in competition.find("RESULTS").findall("RESULT"):
                relay = result.find("RELAY")
                if relay and result.attrib["disqualificationid"] != "0" and relay.find("RELAYPOSITIONS") is not None:
                    self.par_equipe = len(list(relay.find("RELAYPOSITIONS").findall("RELAYPOSITION")))
                    break

            if self.par_equipe == 0:
                logging.error("Taille d'équipe non trouvée")

        # Swimmers
        for result in competition.find("RESULTS").findall("RESULT"):
            reunion = races[result.attrib["raceid"]]
            for record in list(result):
                if self.par_equipe != 0:
                    club = self.conf.clubs.get(int(result.attrib["clubid"]), None)
                    team = int(result.attrib["team"])
                    sexe = conf.nages[int(result.attrib["raceid"])][2]
                    if club is not None:
                        reunion.participants[club].append("{} {}".format(team, sexe))

                elif record.tag == "SOLO":
                    nageurid = int(record.attrib["swimmerid"])
                    club = nageurs[nageurid]
                    if club is not None:
                        reunion.participants[club].append(nageurid)
                        reunion.engagements[club] += 1

                elif record.tag == "RELAY":
                    positions = record.find("RELAYPOSITIONS")
                    if positions:
                        for relay_position in positions:
                            nageurid = int(relay_position.attrib["swimmerid"])
                            club = nageurs[nageurid]
                            if club is not None:
                                reunion.participants[club].append(nageurid)
                                reunion.engagements[club] += 1

                elif record.tag == "SPLIT":
                    pass

        # Counts number of participations per club
        for reunion in self.reunions:
            for club, l in reunion.participants.items():
                reunion.participations[club] = len(set(l))
                if self.par_equipe != 0:
                    reunion.engagements[club] = reunion.participations[club] * self.par_equipe

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
    def __init__(self, index, competition):
        self.index = index
        self.competition = competition
        self.titre = "Réunion N°{}".format(index)
        self.officiels = {}
        self.participants = None
        self.participations = None
        self.engagements = None
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
        if self.competition.par_equipe:
            if participations <= 1:
                needed = (participations, participations)
            else:
                needed = (1, min(3, participations))

        else:
            if participations <= 1:
                needed = (0, 0)
            elif participations <= 10:
                needed = (0, 1)
            elif participations <= 20 or not self.competition.departemental():
                needed = (1, 2)
            else:
                needed = (1, 3)

        if type(details) is list:
            s = "{} officiels requis".format(needed[1])
            if needed[0] > 0:
                s += ", dont {} A ou B".format(needed[0])
            details.append(s)

        num_ab, num = 0, 0
        club_officiels = self.officiels_per_club().get(club, [])
        for officiel in club_officiels:
            if not officiel.valid:
                if self.competition.departemental():
                    logging.warning("L'officiel {} n'est pas valide et est ignoré".format(str(officiel)))
                    continue
                else:
                    # Régional: all officials
                    officiel.valid = True
            num += 1
            if officiel.get_level() > conf.niveau_c:
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
        return "C{}_R{}".format(self.competition.id, self.index)

if __name__ == "__main__":
    import gen_pdf

    parser = argparse.ArgumentParser(description='Liste des compétitions')
    parser.add_argument("--conf", default="Officiels.xlsx", help="Fichier de configuration")
    parser.add_argument("--plot", default=False, action="store_true", help="Ajout des graphiques")
    parser.add_argument("--extract", default=False, action="store_true", help="Extrait le fichier ZIP en XML")
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

    if args.extract:
        import xml.dom.minidom
        for f in ffnex_files:
            logging.info("Extraction du fichier {}".format(f))
            output_file = os.path.splitext(f)[0] + ".xml"
            backup_file = f + ".bak"
            filename = None
            try:
                if zipfile.is_zipfile(f):
                    z = zipfile.ZipFile(f, 'r')
                    if "ffnex.xml" not in z.namelist():
                        logging.error("Le fichier {} devrait contenir un fichier ffnex.xml")
                    else:
                        filename = z.extract('ffnex.xml', tempfile.gettempdir())
                    z.close()

                logging.debug("Fichier de sortie {}".format(output_file))
                data = xml.dom.minidom.parse(filename)
                with open(output_file, "w+") as of:
                    of.write(data.toprettyxml())

                logging.debug("Fichier de sauvegarde {}".format(backup_file))
                os.rename(f, backup_file)

            except zipfile.BadZipfile:
                logging.error("Le fichier {} ne peut pas être lu correctement".format(filename))

        exit(0)

    for f in ffnex_files:
        competitions.append(Competition(conf, f))

    competition_ids = sorted([competition.id for competition in competitions])
    if len(competitions) != len(set(competition_ids)):
      duplicates = []
      for i in range(1, len(competition_ids)):
        if competition_ids[i] == competition_ids[i-1]:
          duplicates.append(competition_ids[i])
      logging.fatal("Des compétitions sont dupliquées: {}".format(", ".join(map(str, duplicates))))
      exit(-1)

    points = {"Départemental": {"participations": 0, "engagements": 0, "total_bonus": 0},
              "Régional":      {"participations": 0, "engagements": 0, "total_bonus": 0}}

    points_df = []

    logging.info("Génération du fichier PDF {}".format(args.output))
    doc = gen_pdf.DocTemplate(conf, args.output, "Liste des compétitions", "Cédric Airaud")
    for competition in competitions:
        if competition.departemental():
            niveau = "Départemental"
        else:
            niveau = "Régional"

        l = points[niveau]

        for club in competition.clubs:
            if club not in l:
                l[club] = 0

        for reunion in competition.reunions:
            for club in reunion.participations.keys():
                pts = reunion.points(club, details=[])
                participations = reunion.participations.get(club, 0)
                engagements = reunion.engagements.get(club, 0)
                officiels = reunion.officiels_per_club().get(club, [])
                num_officiels = len([o for o in officiels if o.valid])

                l["participations"] += participations
                l["engagements"] += engagements
                if pts > 0:
                    l["total_bonus"] += pts
                l[club] += pts

                points_df.append({"Niveau": niveau, "Equipe": competition.par_equipe != 0,
                                  "Compétition": competition.titre(), "Date": competition.startdate,
                                  "Réunion": reunion.index, "Club": club.nom, "Participations": participations,
                                  "Engagements": engagements, "Points": pts, "Officiels": num_officiels})

    points_df = pd.DataFrame(points_df)
    points_df.to_excel("export.xlsx", sheet_name="Points")

    doc.bonus = {level: 0.50 * l["engagements"] / l["total_bonus"] if l["total_bonus"] else 0
                 for level, l in points.items()}
    for level, value in doc.bonus.items():
        logging.info("Valeur du point bonus pour {}: {:.2f} € (Total engagements: {}, total bonus: {:.2f})"
                     .format(level, value, points[level]["engagements"], points[level]["total_bonus"]))

    for club in sorted(conf.clubs.values(), key=lambda x: "{} {}".format(x.departement, x.nom)):
        doc.new_club(club)

    for competition in sorted(competitions, key=lambda x: x.startdate):
        doc.new_competition(competition)

    doc.build()
