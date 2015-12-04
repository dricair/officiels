# -* -coding: utf-8 -*-

"""
.. module: models.py
   :platform: Unix, Windows
   :synopsys: List of the models for tickets structure

.. moduleauthor: Cedric Airaud <cairaud@gmail.com>
"""

from bs4 import BeautifulSoup
import requests
import xlrd

jury_url = "http://ffn.extranat.fr/webffn/resultats.php?idact=nat&idcpt={competition}&go=off"
clubs_url = "http://ffn.extranat.fr/webffn/resultats.php?idact=nat&idcpt={competition}&go=clb"

"""
Represent a competition, composed of several Reunions
"""
class Competition:
    def __init__(self, index, titre, type, date):
        self.index = index
        self.titre = titre
        self.type = type
        self.date = date
        self.reunions = []
        self.participations = {}

        self.per_equipe = False
        self.regionale = False

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
    def __init__(self, nom, club, b_depuis=None):
        self.nom = nom
        self.club = club
        self.b_depuis = b_depuis

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
    def __init__(self):
        self.officiels = {}
        self.clubs = {}
        self.postes = {}
        self.xl_workbook = None


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


def parseJury(competition_index):
    """
    :param competition_index: Index of the competition
    :type int
    :return: Competition with Jury
    :rtype: Competition
    """
    url = jury_url.format(competition=competition_index)
    data = requests.get(url).text

    soup = BeautifulSoup(data, 'html.parser')

    entete = soup.find("fieldset", class_="enteteCompetition")
    spans = entete.find_all("span")
    competition = Competition(index=competition_index, titre=spans[0].text, type=spans[1].text, date=entete.text.splitlines()[-1])

    reunions = []
    table = entete.find_next_sibling("table")

    for tr in table.find_all("tr"):
        tds = tr.find_all("td")
        if tds[0]['id'] == "mainResEpr":
            reunion = Reunion(titre=tds[0].text)
            competition.reunions.append(reunion)
        else:
            if len(tds) != 3:
                raise Exception("Besoin de 3 colonnes par officiel: " + tds.text)
            if not reunion:
                raise Exception("Pas d'entête de réunion trouvé: " + tds.text)
            reunion.officiels.append(Officiel(nom=tds[1].text, club=tds[2].text))

            poste = tds[0].text.replace(":", "").strip()

    return competition

def parseParticipations(competition):
    url = clubs_url.format(competition=competition.index)
    data = requests.get(url).text

    soup = BeautifulSoup(data, 'html.parser')

    table = soup.find("table", class_="tableau")
    for tr in table.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) == 13:
            tds[1].b.clear()
            competition.participations[tds[1].a.text.strip()] = int(tds[4].text)



def getConfiguration(filename):
    conf = Configuration()
    conf.xl_workbook = xlrd.open_workbook(filename)
    sheet_names = conf.xl_workbook.sheet_names()
    print("Configuration depuis le fichier '{}:".format(filename))

    sheets = {'clubs': 'Clubs', 'officiels': 'Officiels', 'postes': 'Postes'}
    if len(set(conf.xl_workbook.sheet_names()) & set(sheets.values())) != 3:
        raise Exception("Le fichier {} doit contenir les pages {} (Trouvées: {})".format(
            filename, ', '.join(sheets.values()), ', '.join(conf.xl_workbook.sheet_names())))

    print("- Lecture des clubs")
    xl_sheet = conf.xl_workbook.sheet_by_name(sheets['clubs'])
    row = xl_sheet.row(0)
    if row[0].value != "Club" or row[1].value != "Département":
        raise Exception("La page 'Clubs' doit contenir des colonnes 'Club' et 'Département' (Trouvées: {})".format(
            ", ".join([cell.value for cell in row])))
    for row_idx in range(1,xl_sheet.nrows):
        row = xl_sheet.row(row_idx)
        if row[0].value != "":
            club = Club(nom=row[0].value, departement=row[1].value)
            conf.clubs[club.nom] = club

    print("- Lecture des officiels")
    xl_sheet = conf.xl_workbook.sheet_by_name(sheets['officiels'])
    row = xl_sheet.row(0)
    if row[0].value != "Nom" or row[1].value != "Club" or row[2].value != "A depuis" or row[3].value != "B depuis":
        raise Exception("La page 'Officiels' doit contenir des colonnes 'Nom', 'Club', 'A depuis' et 'B depuis' "
                        "(Trouvées: {})".format(", ".join([cell.value for cell in row])))
    for row_idx in range(1,xl_sheet.nrows):
        row = xl_sheet.row(row_idx)
        if row[0].value != "":
            club = row[1].value
            if club not in clubs:
                print("WARNING: Le club {} pour l'officiel {} n'a pas été trouvé".format(club, row[0].value))
            else:
                club = clubs[club]
                officiel = Officiel(nom=row[0].value, club=club, b_depuis=row[2].value)
                conf.officiels[officiel.nom] = officiel

    print("- Lecture des postes")
    xl_sheet = conf.xl_workbook.sheet_by_name(sheets['postes'])
    row = xl_sheet.row(0)
    if row[0].value != "Poste" or row[1].value != "Niveau":
        raise Exception("La page 'Postes' doit contenir des colonnes 'Postes' et 'Niveau' "
                        "(Trouvées: {})".format(", ".join([cell.value for cell in row])))
    for row_idx in range(1,xl_sheet.nrows):
        row = xl_sheet.row(row_idx)
        if row[0].value != "":
            conf.postes[row[0].value] = row[1].value

    return conf





if __name__ == "__main__":
    conf = getConfiguration('Officiels.xls')
    for club in conf.clubs.values():
        print(club)
    for officiel in conf.officiels.values():
        print(officiel)
    for role in conf.postes:
        print("{} ({})".format(role, conf.postes[role]))
    exit()

    competition = parseJury(33007)
    parseParticipations(competition)

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










