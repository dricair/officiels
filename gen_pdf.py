# -* -coding: utf-8 -*-

"""
.. module: gen_pdf.py
   :platform: Unix, Windows
   :synopsys: Generate PDF for a list of competitions

.. moduleauthor: Cedric Airaud <cairaud@gmail.com>
"""

import logging
import datetime

from reportlab.platypus import BaseDocTemplate, PageTemplate, NextPageTemplate, PageBreak, Frame
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.platypus.flowables import ListItem, ListFlowable
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont, TTFError

from officiels import Club, Competition

PAGE_HEIGHT = A4[1]
PAGE_WIDTH = A4[0]
styles = getSampleStyleSheet()

try:
    pdfmetrics.registerFont(TTFont("Trebuchet", "/usr/share/fonts/truetype/msttcorefonts/Trebuchet_MS.ttf"))
    pdfmetrics.registerFont(TTFont("Trebuchet-bold", "/usr/share/fonts/truetype/msttcorefonts/Trebuchet_MS_Bold.ttf"))
    pdfmetrics.registerFont(TTFont("Trebuchet-italic",
                                   "/usr/share/fonts/truetype/msttcorefonts/Trebuchet_MS_Italic.ttf"))
    pdfmetrics.registerFont(TTFont("Trebuchet-bold-italic",
                                   "/usr/share/fonts/truetype/msttcorefonts/Trebuchet_MS_Bold_Italic.ttf"))
    trebuchet_available = True
except TTFError as e:
    logging.warning("Font Trebuchet non disponible")
    trebuchet_available = False

sNormal = styles['Normal']
sHeading2 = styles['Heading2']

header_table_style = {
    "Départemental": TableStyle([('BOX',        (0, 0), (-1, -1), 0.25, "#FF4500"),
                                 ('ALIGN',      (0, 0), (-1, -1), 'CENTER'),
                                 ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
                                 ('BACKGROUND', (0, 0), (-1, -1), "#FF8C00"),
                                 ('TEXTCOLOR',  (0, 0), (-1, -1), colors.white),
                                 ('FONTSIZE',   (0, 0), (-1,  0), 14),
                                 ('FONTSIZE',   (0, 1), (-1, -1), 10),
                                 ]),
    "Régional":      TableStyle([('BOX',        (0, 0), (-1, -1), 0.25, "#2B7739"),
                                 ('ALIGN',      (0, 0), (-1, -1), 'CENTER'),
                                 ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
                                 ('BACKGROUND', (0, 0), (-1, -1), "#48B35E",),
                                 ('TEXTCOLOR',  (0, 0), (-1, -1), colors.white),
                                 ('FONTSIZE',   (0, 0), (-1,  0), 14),
                                 ('FONTSIZE',   (0, 1), (-1, -1), 10),
                                 ]),

    "Club":          TableStyle([('BOX',        (0, 0), (-1, -1), 0.25, "#0e23a2"),
                                 ('ALIGN',      (0, 0), (-1, -1), 'CENTER'),
                                 ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
                                 ('BACKGROUND', (0, 0), (-1, -1), "#132fdb"),
                                 ('TEXTCOLOR',  (0, 0), (-1, -1), colors.white),
                                 ('FONTSIZE',   (0, 0), (-1, -1), 14)]),

    "Content":       TableStyle([('BOX',        (0, 0), (-1, -1), 0.25, colors.black),
                                 ('INNERGRID',  (0, 0), (-1, -1), 0.25, colors.black),
                                 ('BACKGROUND', (0, 0), (-1,  0), "#DCE2F1"),
                                 ('FONTSIZE',   (0, 0), (-1, -1), 10),
                                 ('VALIGN',     (0, 0), (-1, -1), "TOP"),
                                 ])
    }


class ClubTemplate(PageTemplate):
    """
    Default template for a Club
    """
    def __init__(self):
        super().__init__(id='club')
        self.page_width = PAGE_WIDTH-2*cm
        self.frames.append(Frame(x1=cm, y1=2*cm, height=PAGE_HEIGHT-4*cm, width=self.page_width))

    def beforeDrawPage(self, canv, doc):
        logging.debug("ClubTemplate.beforeDrawPage, club={}".format(str(doc.club)))

        canv.saveState()
        canv.setFont('Times-Roman', 9)
        canv.setFillColor(colors.grey)
        canv.setStrokeColor(colors.grey)

        # Header
        if not doc.newClub:
            canv.drawCentredString(PAGE_WIDTH/2.0, PAGE_HEIGHT-cm, doc.club.nom)
            canv.line(cm, PAGE_HEIGHT-1.3*cm, PAGE_WIDTH-cm, PAGE_HEIGHT-1.3*cm)
        doc.newClub = False

        # Footer
        canv.drawCentredString(PAGE_WIDTH/2.0, cm,
                               "Page {} - Mis à jour le {}".format(doc.page,
                                                                   datetime.date.today().strftime("%d/%m/%Y")))

        canv.restoreState()


class ReunionTemplate(PageTemplate):
    """
    Default template for a Competition/Reunion
    """
    def __init__(self):
        super().__init__(id='reunion')
        self.page_width = PAGE_WIDTH-2*cm
        self.frames.append(Frame(x1=cm, y1=2*cm, height=PAGE_HEIGHT-4*cm, width=self.page_width))

    def beforeDrawPage(self, canv, doc):
        logging.debug("ReunionTemplate.beforeDrawPage, newCompetition={}".format(str(doc.newCompetition)))

        canv.saveState()
        canv.setFont('Times-Roman', 9)
        canv.setFillColor(colors.grey)
        canv.setStrokeColor(colors.grey)

        # Header
        if not doc.newCompetition:
            canv.drawCentredString(PAGE_WIDTH/2.0, PAGE_HEIGHT-cm, doc.competition.nom)
            canv.line(cm, PAGE_HEIGHT-1.3*cm, PAGE_WIDTH-cm, PAGE_HEIGHT-1.3*cm)
        doc.newCompetition = False

        # Footer
        canv.drawCentredString(PAGE_WIDTH/2.0, cm,
                               "Page {} - Mis à jour le {}".format(doc.page,
                                                                   datetime.date.today().strftime("%d/%m/%Y")))

        canv.restoreState()


class DocTemplate(BaseDocTemplate):
    def __init__(self, conf, filename, title, author):
        super().__init__(filename, pagesize=A4, title=title, author=author)
        self.addPageTemplates([ClubTemplate(), ReunionTemplate()])
        self.story = []
        self.conf = conf
        self.competition = None
        self.newCompetition = True
        self.club = None
        self.newClub = True
        self.page_width = self.pageTemplates[0].page_width
        self.club_seen = 0
        self.competition_seen = False

    def new_club(self, club):
        """
        Add a club on a new page
        :param club: Club to print
        :type club: Club
        """
        logging.debug("New club: " + club.nom)
        if not self.story:
            # For the first page
            self.club = club
        else:
            self.story.append(NextPageTemplate(club))
            self.story.append(PageBreak())

        table_style = header_table_style["Club"]
        table = Table([[club.nom]], [self.page_width], 2 * cm, style=table_style)
        table.link_object = (club, club.nom)
        self.story.append(table)

        table_style = header_table_style["Content"]

        for departemental in (True, False):
            total = 0
            if departemental:
                competitions = [c for c in club.competitions if c.departemental()]
                self.story.append(Paragraph("Compétitions départementales", sHeading2))
                bonus = self.bonus["Départemental"]
            else:
                competitions = [c for c in club.competitions if not c.departemental()]
                self.story.append(Paragraph("Compétitions régionales et plus", sHeading2))
                bonus = self.bonus["Régional"]

            table_data = [["Compétition", "Réunions", "Points"]]

            for competition in sorted(competitions, key=lambda c: c.startdate):
                row = [Paragraph("{} - {}<br/>{}".format(competition.date_str(), competition.titre(),
                                                         competition.type), sNormal), [], []]
                for reunion in competition.reunions:
                    pts = reunion.points(club)
                    total += pts
                    row[1].append(Paragraph(reunion.titre, sNormal))
                    row[2].append(Paragraph("{} points".format(pts), sNormal))
                table_data.append(row)

            table = Table(table_data, [self.page_width * x for x in (0.70, 0.15, 0.15)], style=table_style)
            self.story.append(table)
            self.story.append(Paragraph("<br/>Total des points: {}".format(total), sNormal))

            if total < 0:
                self.story.append(Paragraph("Valeur du malus: {:.2f} €".format(total * 10), sNormal))
            else:
                self.story.append(Paragraph("Valeur du bonus (Estimation): {:.2f} €"
                                            .format(total * bonus), sNormal))

    def new_competition(self, competition):
        """
        Add a competition on a new page
        :param competition: New competition
        :type competition: Competition
        """
        logging.debug("New competition: " + competition.titre())

        if not self.story:
            # For the first page
            self.competition = competition
        else:
            self.story.append(NextPageTemplate(competition))
            self.story.append(PageBreak())

        if competition.departemental():
            table_style = header_table_style["Départemental"]
        else:
            table_style = header_table_style["Régional"]

        table_data = [[competition.titre()], [competition.type], [competition.date_str()]]
        table = Table(table_data, [self.page_width], [cm, 0.5*cm, 0.5*cm], style=table_style)
        table.link_object = (competition, competition.titre())
        self.story.append(table)

        if competition.reunions:
            for reunion in competition.reunions:
                self.new_reunion(reunion)
        else:
            self.story.append(Paragraph("Pas de résultats trouvés pour cette compétition", sNormal))

    def new_reunion(self, reunion):
        logging.debug("New reunion: " + reunion.titre)

        p = Paragraph(reunion.titre, styles["h2"])
        p.link_object = (reunion, reunion.titre)
        self.story.append(p)

        table_style = header_table_style["Content"]
        table_data = [["Club", "Officiels", "Points"]]
        off_per_club = reunion.officiels_per_club()
        total_participations, total_engagements = 0, 0
        for club, num in sorted(reunion.participations.items(), key=lambda c: c[0].nom):
            total_participations += num
            total_engagements += reunion.engagements.get(club, 0)
            if reunion.competition.par_equipe != 0:
                participations = "{} équipes".format(num)
            else:
                participations = "{} participations".format(num)

            details = []
            points = reunion.points(club, details)
            paragraph_points = [Paragraph("<b>{} points</b>".format(points), sNormal)]
            if len(details) > 0:
                paragraph_points.append(ListFlowable([ListItem(Paragraph(d, sNormal), leftIndex=20, value='-')
                                                      for d in details], bulletType='bullet'))
            else:
                print("No details")

            officiels = []
            for off in sorted(off_per_club.get(club, []), key=lambda o: o.nom):
                officiels.append("{}: {} {}".format(str(off.get_level()), off.prenom, off.nom))
                if not off.valid:
                    officiels[-1] = "<strike>{}</strike>".format(officiels[-1])
            paragraph_officiels = Paragraph("<br/>".join(officiels), sNormal)

            table_data.append([club.nom + "\n" + participations, paragraph_officiels, paragraph_points])

        self.story.append(Table(table_data, 3 * [self.page_width / 3.0], style=table_style))
        self.story.append(Paragraph("<br/>Total des participations: {}".format(total_participations), sNormal))
        self.story.append(Paragraph("Total des engagements: {}".format(total_engagements), sNormal))

    def build(self):
        """
        Create the PDF
        """
        super().build(self.story)

    # Handler functions
    def handle_pageBegin(self):
        """
        New page starting
        """
        logging.debug("doc.handle_pageBegin")
        self.newCompetition = False
        super().handle_pageBegin()

    def handle_nextPageTemplate(self, pt):
        """
        Changing page template
        """
        logging.debug("doc.handle_nextPageTemplate")
        if type(pt).__name__ not in ("Club", "Competition"):
            logging.fatal("pt is not of correct class: {}".format(type(pt).__name__))

        if type(pt).__name__ == "Club":
            self.newClub = True
            self.club = pt
            super().handle_nextPageTemplate('club')
        else:
            self.newCompetition = True
            self.competition = pt
            super().handle_nextPageTemplate('reunion')

    def handle_flowable(self, flowables):
        f = flowables[0]
        BaseDocTemplate.handle_flowable(self, flowables)

        if hasattr(f, "link_object"):
            (obj, text) = f.link_object

            if type(obj).__name__ == "Club":
                if obj.departement != self.club_seen:
                    key = "Departement{}".format(obj.departement)
                    self.canv.bookmarkPage(key)
                    self.canv.addOutlineEntry("Département {}".format(obj.departement), key, level=0, closed=1)
                    self.club_seen = obj.departement
                level = 1

            elif type(obj).__name__ == "Competition":
                if not self.competition_seen:
                    key = "Competition"
                    self.canv.bookmarkPage(key)
                    self.canv.addOutlineEntry("Détail des compétitions", key, level=0, closed=1)
                    self.competition_seen = True
                level = 1

            else:
                level = 2

            key = obj.link()
            self.canv.bookmarkPage(key)
            self.canv.addOutlineEntry(text, key, level=level, closed=1)




