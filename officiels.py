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
log = logging.getLogger(__name__)
log.setLevel(logging.DEBUG)

jury_url = "http://ffn.extranat.fr/webffn/resultats.php?idact=nat&idcpt={competition}&go=off"
clubs_url = "http://ffn.extranat.fr/webffn/resultats.php?idact=nat&idcpt={competition}&go=clb"

"""
Represent a competition, composed of several Reunions
"""
class Competition:
    def __init__(self, conf, competition_index):
        """
        :param conf: Configuration structure
        :type Configuration
        :param competition_index: Index of the competition
        :type int
        """
        self.index = competition_index
        url = jury_url.format(competition=competition_index)
        log.debug("Jury et réunions: " + url)
        data = requests.get(url).text
        soup = BeautifulSoup(data, 'html.parser')

        entete = soup.find("fieldset", class_="enteteCompetition")
        spans = entete.find_all("span")
        self.type, self.titre, self.date = spans[0].text, spans[1].text, entete.text.splitlines()[-1]
        log.debug("{} - {} - {} ".format(self.type, self.titre, self.date))

        self.per_equipe = False
        self.regionale = False

        self.reunions = []
        self.participations = {}
        table = entete.find_next_sibling("table")

        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if tds[0]['id'] == "mainResEpr":
                reunion = Reunion(titre=tds[0].text)
                self.reunions.append(reunion)
                log.debug("Réunion trouvée: " + str(reunion))
            else:
                if len(tds) != 3:
                    log.fatal("Besoin de 3 colonnes par officiel: " + tds.text)
                if not reunion:
                    log.fatal("Pas d'entête de réunion trouvé: " + tds.text)
                nom, club = tds[1].text, tds[2].text
                if club in conf.clubs:
                    reunion.officiels.append(conf.findOfficiel(nom=nom, club=club))
                    poste = tds[1].text, tds[0].text.replace(":", "").strip()
                else:
                    log.warning("Officiel ignoré: {} car le club {} n'est pas dans la liste".format(nom, club))

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
                    log.warning("Club {} ignoré pour les participations car pas dans la liste".format(club))


    def __str__(self):
        return "{titre}\n{type}\n{date}\n\n".format(**self.__dict__) + "\n\n".join(map(str, self.reunions))



"""
Represent a Reunion, base for the calculation
"""
class Reunion:
    def __init__(self, titre):
        self.titre = titre
        self.officiels = []
        self.officiels_per_club = None

    def __str__(self):
        return self.titre + "\n  " + "\n  ".join(map(str, self.officiels))

    def officielsPerClub(self):
        """
        Sort officiels per club
        """
        if self.officiels_per_club:
            return self.officiels_per_club

        self.officiels_per_club = {}
        for officiel in self.officiels:
            if not officiel.club in self.officiels_per_club:
                self.officiels_per_club[officiel.club] = []
            self.officiels_per_club[officiel.club].append(officiel)

        return self.officiels_per_club


"""
Represent an Officiel
"""
class Officiel:
    def __init__(self, nom, club, index, b_depuis=None, a_depuis=None):
        self.nom = nom
        self.club = club
        self.index = index
        self.b_depuis = b_depuis
        self.a_depuis = a_depuis

    def __str__(self):
        return "{nom} ({club})".format(**self.__dict__)



"""
Club
"""
class Club:
    def __init__(self, nom, departement):
        self.nom = nom
        self.departement = departement

    def __str__(self):
        return "{} ({})".format(self.nom, self.departement)


"""
Global configuration
"""
class Configuration:
    def __init__(self, filename):
        self.officiels = {}
        self.clubs = {}
        self.postes = {}
        self.dirty = False
        self.filename = filename

        self.wb = load_workbook(filename, guess_types=True)
        log.info("Configuration depuis le fichier '{}:".format(filename))

        self.sheets = {'clubs': 'Clubs', 'officiels': 'Officiels', 'postes': 'Postes', 'competitions': 'Compétitions'}
        if len(set(self.wb.get_sheet_names()) & set(self.sheets.values())) != 4:
            raise Exception("Le fichier {} doit contenir les pages {} (Trouvées: {})".format(
                filename, ', '.join(self.sheets.values()), ', '.join(self.wb.get_sheet_names())))

        log.info("- Lecture des clubs")
        xl_sheet = self.wb.get_sheet_by_name(self.sheets['clubs'])
        header = True
        row = xl_sheet.rows[0]
        if row[0].value != "Club" or row[1].value != "Département":
            raise Exception("La page 'Clubs' doit contenir des colonnes 'Club' et 'Département' (Trouvées: {})".format(
                ", ".join([cell.value for cell in row])))
        for row in xl_sheet.rows[1:]:
            if row[0].value != "":
                club = Club(nom=row[0].value, departement=row[1].value)
                self.clubs[club.nom] = club

        log.info("- Lecture des officiels")
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
                    officiel = Officiel(nom=row[0].value, club=club, a_depuis=row[2].value, b_depuis=row[3])
                    self.officiels[officiel.nom] = officiel

        log.info("- Lecture des postes")
        xl_sheet = self.wb.get_sheet_by_name(self.sheets['postes'])
        row = xl_sheet.rows[0]
        if row[0].value != "Poste" or row[1].value != "Niveau":
            raise Exception("La page 'Postes' doit contenir des colonnes 'Postes' et 'Niveau' "
                            "(Trouvées: {})".format(", ".join([cell.value for cell in row])))
        for row in xl_sheet.rows[1:]:
            if row[0].value != "":
                self.postes[row[0].value] = row[1].value


    def findOfficiel(self, nom, club):
        """
        Find an officiel by name if it exists
        """
        if not nom in self.officiels:
            log.warning("L'officiel {} (Club {}) n'existe pas".format(nom, club))
            officiel = Officiel(nom, club)
            self.officiels[nom] = officiel
            sheet = self.wb.get_sheet_by_name(self.sheets['officiels'])
            num_rows = len(sheet.rows)
            sheet.cell(row=num_rows+1, column=1, value=nom)
            sheet.cell(row=num_rows+1, column=2, value=club)
            self.dirty = True

        return self.officiels[nom]


    def checkRole(self, officiel, poste, date):
        """
        Check that the poste matches the level for the Officiel
        :param officiel Officiel to check
        :type officiel Officiel
        :param poste Name of the poste
        :type poste basestring
        :param date date of competition to check
        :type date datetime
        """
        if not poste in self.postes:
            log.error("Le poste '{}' n'est pas listé dans le fichier de configuration")
            return True

        niveau = conf.postes[poste]
        update = False
        if niveau == 'A' and not self.a_depuis:
            log.warning("L'officiel {} semble avoir le niveau A")
            officiel.a_depuis = date
            if not officiel.b_depuis: officiel.b_depuis =
            update = True

        elif niveau == 'B' and not self.b_depuis:
            log.warning("L'officiel {} semble avoir le niveau A")
            update = True

        if update:
            for row in self.wb.get_sheet_by_name(self.sheets['officiels']).rows[1]:
                if row[0] == officiel.nom:
                    row[2] = officiel.a_depuis
                    row[3] = officiel.b_depuis
                    break
            self.dirty = True



    def save(self):
        """
        Save the file if it has been updated
        """
        if self.dirty:
            backup_filename = self.filename + ".bak"
            log.info("Mise à jour du fichier {} (Sauvegarde: {})".format(self.filename, backup_filename))
            os.rename(self.filename, backup_filename)
            try:
                self.wb.save(self.filename)
            except Exception as e:
                os.rename(backup_filename, self.filename)
                log.error("Erreur lors de la mise à jour, restoration de la sauvegarde.\n" + str(e))
            self.dirty = False




def points(competition, reunion, club):
    """
    :param competition: Competition to use
    :type competition: Competition
    :param reunion: Reunion to use
    :type reunion: Reunion
    :param club: Club to look for
    :type club: string
    :return: Number of points
    :rtype: int
    """

    participations = competition.participations[club]
    officiels = reunion.officielsPerClub().get(club, [])






if __name__ == "__main__":
    conf = Configuration('Officiels.xlsx')
    competition = Competition(conf, 33007)
    conf.save()

    for reunion in competition.reunions:
        print(reunion.titre)

        officiels_per_club = reunion.officielsPerClub()
        for club, num in sorted(competition.participations.items()):
            officiels = officiels_per_club.get(club, [])
            print("  {club:30s}: {participations} participations, {officiels} officiels {officiels_str}".format(
                club=club, participations=num, officiels=len(officiels),
                officiels_str = " ({})".format(", ".join([off.nom for off in officiels])) if len(officiels) else ""
            ))

        print("\n")










