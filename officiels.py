# -* -coding: utf-8 -*-

"""
.. module: models.py
   :platform: Unix, Windows
   :synopsys: List of the models for tickets structure

.. moduleauthor: Cedric Airaud <cairaud@gmail.com>
"""

import os
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import requests

import logging
logging.basicConfig(level=logging.INFO, format="%(levelname)-9s %(lineno)-4s %(message)s")

jury_url = "http://ffn.extranat.fr/webffn/resultats.php?idact=nat&idcpt={competition}&go=off"
clubs_url = "http://ffn.extranat.fr/webffn/resultats.php?idact=nat&idcpt={competition}&go=clb"


class Competition:
    """
    Represent a competition, composed of several Reunions
    """

    def __init__(self, conf, competition_index):
        """
        :param conf Configuration structure
        :type conf Configuration
        :param competition_index: Index of the competition
        :type competition_index int
        """
        self.index = competition_index

        # Get some info from configuration
        self.date = conf.competitions[competition_index]["date"]
        self.equipe = conf.competitions[competition_index]["equipe"]
        self.regional = conf.competitions[competition_index]["regional"]

        if self.equipe:
            try:
                self.equipe = int(self.equipe)
            except ValueError as e:
                logging.fatal("La colonne Equipe doit être un nombre pour une compétition par équipe")

        url = jury_url.format(competition=competition_index)
        logging.debug("Jury et réunions: " + url)
        data = requests.get(url).text
        soup = BeautifulSoup(data, 'html.parser')

        entete = soup.find("fieldset", class_="enteteCompetition")
        spans = entete.find_all("span")
        self.type, self.titre = spans[0].text, entete.text.splitlines()[-1]
        logging.debug("{} - {} - {} ".format(self.type, self.titre, self.date))

        self.reunions = []
        self.participations = {}
        table = entete.find_next_sibling("table")

        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if tds[0]['id'] == "mainResEpr":
                reunion = Reunion(self, titre=tds[0].text.strip())
                self.reunions.append(reunion)
                logging.debug("Réunion trouvée: " + str(reunion))
            else:
                if len(tds) != 3:
                    logging.fatal("Besoin de 3 colonnes par officiel: " + tds.text)
                if not reunion:
                    logging.fatal("Pas d'entête de réunion trouvé: " + tds.text)
                poste, nom, club = tds[0].text.replace(":", "").strip(), tds[1].text, tds[2].text
                if poste in conf.postes and not conf.postes[poste]:
                    logging.debug("{} au poste {} est ignoré".format(nom, poste))
                elif club in conf.clubs:
                    officiel = conf.find_officiel(nom=nom, club=club)
                    logging.debug("Officiel trouvé: " + str(officiel))
                    if officiel not in reunion.officiels and conf.checkPoste(officiel, poste):
                        reunion.officiels.append(officiel)
                elif club != "NATATION AZUR":
                    logging.warning("Officiel ignoré: {} car le club {} n'est pas dans la liste".format(nom, club))

        # Parse participations
        url = clubs_url.format(competition=competition_index)
        data = requests.get(url).text

        soup = BeautifulSoup(data, 'html.parser')

        table = soup.find("table", class_="tableau")
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) == 13:
                tds[1].b.clear()
                club, num = tds[1].a.text.strip(), int(tds[4].text)
                if club in conf.clubs:
                    self.participations[club] = num
                else:
                    logging.warning("Club {} ignoré pour les participations car pas dans la liste".format(club))

    def __str__(self):
        return "{titre}\n{type}\n{date}\n\n".format(**self.__dict__) + "\n\n".join(map(str, self.reunions))


class Reunion:
    """
    Represent a Reunion, base for the calculation
    """
    def __init__(self, competition, titre):
        self.competition = competition
        self.titre = titre
        self.officiels = []
        self._officiels_per_club = None

    def __str__(self):
        return self.titre + "\n  " + "\n  ".join(map(str, self.officiels))

    def officiels_per_club(self):
        """
        Sort officiels per club
        """
        if self._officiels_per_club:
            return self._officiels_per_club

        self._officiels_per_club = {}
        for officiel in self.officiels:
            if officiel.club.nom not in self._officiels_per_club:
                self._officiels_per_club[officiel.club.nom] = []
            self._officiels_per_club[officiel.club.nom].append(officiel)

        return self._officiels_per_club

    def points(self, club):
        """
        :param club: Club to look for
        :type club: Club
        :return: Number of points
        :rtype: int
        """
        participations = self.competition.participations.get(club.nom, 0)

        # needed = (Num of A/B, Total num)
        if self.competition.equipe:
            participations = participations // self.competition.equipe
            if participations <= 1:
                needed = (participations, participations)
            else:
                needed = (1, min(3, participations))

        else:
            if participations <= 1:
                needed = (0, 0)
            elif participations < 10:
                needed = (0, 1)
            elif participations < 20 or competition.regional:
                needed = (1, 2)
            else:
                needed = (1, 3)

        num_ab, num = 0, 0
        club_officiels = reunion.officiels_per_club().get(club.nom, [])
        for officiel in club_officiels:
            num += 1
            level = officiel.get_level(self.competition.date)
            if level == 'A' or level == 'B':
                num_ab += 1

        if competition.regional:
            num = min(5, num)

        # TODO: Officiel manquant -4 point, sans bon statut -2 points. Dans quel ordre ?
        if num < needed[1]:
            pts = (needed[1] - num) * -4
        else:
            pts = (num - needed[1]) * 2
            if num_ab < needed[0]:
                pts += (needed[0] - num_ab) * -2

        return pts


class Officiel:
    """
    Represent an Officiel
    """
    def __init__(self, nom, club, b_depuis=None, a_depuis=None):
        self.nom = nom
        self.club = club
        self.b_depuis = b_depuis
        self.a_depuis = a_depuis

    def __str__(self):
        return "{nom} ({club})".format(**self.__dict__)

    def get_level(self, date):
        """
        Return the level for an officiel as a string, at the given date (if specified)
        """
        if self.a_depuis and self.a_depuis < date:
            return 'A'
        elif self.b_depuis and self.b_depuis < date:
            return 'B'
        else:
            return 'C'


class Club:
    """
    Club
    """
    def __init__(self, nom, departement):
        self.nom = nom
        self.departement = departement

    def __str__(self):
        return "{} ({})".format(self.nom, self.departement)


class Configuration:
    """
    Global configuration
    """
    def __init__(self, filename):
        self.officiels = {}
        self.clubs = {}
        self.postes = {}
        self.competitions = {}
        self.dirty = False
        self.filename = filename

        self.wb = load_workbook(filename, guess_types=True)
        logging.info("Configuration depuis le fichier '{}'".format(filename))

        self.sheets = {'clubs': 'Clubs', 'officiels': 'Officiels', 'postes': 'Postes', 'competitions': 'Compétitions'}
        if len(set(self.wb.get_sheet_names()) & set(self.sheets.values())) != 4:
            raise Exception("Le fichier {} doit contenir les pages {} (Trouvées: {})".format(
                filename, ', '.join(self.sheets.values()), ', '.join(self.wb.get_sheet_names())))

        logging.info("- Lecture des clubs")
        xl_sheet = self.wb.get_sheet_by_name(self.sheets['clubs'])
        row = xl_sheet.rows[0]
        if row[0].value != "Club" or row[1].value != "Département":
            raise Exception("La page 'Clubs' doit contenir des colonnes 'Club' et 'Département' (Trouvées: {})".format(
                ", ".join([cell.value for cell in row])))
        for row in xl_sheet.rows[1:]:
            if row[0].value != "":
                club = Club(nom=row[0].value, departement=row[1].value)
                self.clubs[club.nom] = club

        logging.info("- Lecture des officiels")
        xl_sheet = self.wb.get_sheet_by_name(self.sheets['officiels'])
        row = xl_sheet.rows[0]
        if row[0].value != "Nom" or row[1].value != "Club" or row[2].value != "A depuis" or row[3].value != "B depuis":
            raise Exception("La page 'Officiels' doit contenir des colonnes 'Nom', 'Club', 'A depuis' et 'B depuis' "
                            "(Trouvées: {})".format(", ".join([cell.value for cell in row])))
        for index, row in enumerate(xl_sheet.rows[1:]):
            if row[0].value != "":
                club = row[1].value
                if club not in self.clubs:
                    print("WARNING: Le club {} pour l'officiel {} n'a pas été trouvé".format(club, row[0].value))
                else:
                    club = self.clubs[club]
                    officiel = Officiel(nom=row[0].value, club=club, a_depuis=row[2].value, b_depuis=row[3].value)
                    self.officiels[officiel.nom] = officiel

        logging.info("- Lecture des postes")
        xl_sheet = self.wb.get_sheet_by_name(self.sheets['postes'])
        row = xl_sheet.rows[0]
        if row[0].value != "Poste" or row[1].value != "Niveau":
            raise Exception("La page 'Postes' doit contenir des colonnes 'Postes' et 'Niveau' "
                            "(Trouvées: {})".format(", ".join([cell.value for cell in row])))
        for row in xl_sheet.rows[1:]:
            if row[0].value != "":
                self.postes[row[0].value] = row[1].value

        logging.info("- Lecture des compétitions")
        xl_sheet = self.wb.get_sheet_by_name(self.sheets["competitions"])
        row = xl_sheet.rows[0]
        if (row[0].value != "Numéro" or row[1].value != "Date" or row[2].value != "Compétition" or
                row[3].value != "Régional" or row[4].value != "Équipe"):
            raise Exception("La page 'Compétition' doit contenir des colonnes 'Numéro', 'Date' "
                            "'Compétition', 'Régional' et 'Équipe' "
                            "(Trouvées: {})".format(", ".join([cell.value for cell in row])))
        for row in xl_sheet.rows[1:]:
            if row[0].value != "":
                self.competitions[row[0].value] = {"date": row[1].value,
                                                   "titre": row[2].value,
                                                   "regional": row[3].value == "x",
                                                   "equipe": row[4].value}

    def find_officiel(self, nom, club):
        """
        Find an officiel by name if it exists
        """
        if nom not in self.officiels:
            logging.warning("L'officiel {} (Club {}) n'existe pas".format(nom, club))
            officiel = Officiel(nom, club)
            self.officiels[nom] = officiel
            sheet = self.wb.get_sheet_by_name(self.sheets['officiels'])
            num_rows = len(sheet.rows)
            sheet.cell(row=num_rows+1, column=1, value=nom)
            sheet.cell(row=num_rows+1, column=2, value=club)
            self.dirty = True

        return self.officiels[nom]


    def checkPoste(self, officiel, poste):
        """
        Check that the poste matches the level for the Officiel
        :param officiel Officiel to check
        :type officiel Officiel
        :param poste Name of the poste
        :type poste basestring
        """
        if poste not in self.postes:
            logging.error("Le poste '{}' n'est pas listé dans le fichier de configuration".format(poste))
            return False

        niveau = conf.postes[poste]
        if niveau == 'A' and not officiel.a_depuis:
            logging.error("L'officiel {} semble avoir le niveau A".format(officiel.nom))
            return False

        elif niveau == 'B' and not officiel.b_depuis:
            logging.error("L'officiel {} semble avoir le niveau B".format(officiel.nom))
            return False

        return True

    def save(self):
        """
        Save the file if it has been updated
        """
        if self.dirty:
            backup_filename = self.filename + ".bak"
            logging.info("Mise à jour du fichier {} (Sauvegarde: {})".format(self.filename, backup_filename))
            os.rename(self.filename, backup_filename)
            try:
                self.wb.save(self.filename)
            except Exception as e:
                os.rename(backup_filename, self.filename)
                logging.error("Erreur lors de la mise à jour, restoration de la sauvegarde.\n" + str(e))
            self.dirty = False


if __name__ == "__main__":
    conf = Configuration('Officiels.xlsx')
    competition = Competition(conf, 34325)
    conf.save()

    for reunion in competition.reunions:
        print(reunion.titre)

        off_per_club = reunion.officiels_per_club()
        for club, num in sorted(competition.participations.items()):
            officiels = off_per_club.get(club, [])
            if competition.equipe:
                participations = "{} équipes".format(num // competition.equipe)
            else:
                participations = "{} participations".format(num)

            print(" {club:30s}: {participations}, {officiels} officiels {officiels_str} -->"
                  " {points} points".format(
                    club=club, participations=participations, officiels=len(officiels),
                    officiels_str=" ({})".format(", ".join(["{}: {}".format(off.nom, off.get_level(competition.date))
                    for off in officiels])) if len(officiels) else "",
                    points=reunion.points(conf.clubs[club])
            ))

        print("\n")










