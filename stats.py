# -* -coding: utf-8 -*-

"""
.. module: stats.py
   :platform: Unix, Windows
   :synopsys: Extracts statistics and create plots

.. moduleauthor: Cedric Airaud <cairaud@gmail.com>
"""

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import random
import argparse

def get_x(participations):
    return (participations // 10 + 1) if participations > 1 else 0

parser = argparse.ArgumentParser(description='Statistiques')
parser.add_argument("figure", type=int, help="Index de figure")
args = parser.parse_args()
fig_num = args.figure

filename = "export.xlsx"
stats = pd.read_excel(filename, sheetname="Données brutes")

min_range, max_range = (-5, 5) if fig_num in [1] else (-6, 6)
data = pd.Series({x: x*4 if x<0 else x*2 for x in range(min_range, max_range+1)})

if fig_num not in (8, 9, 10, 11, 12):
    plt.figure()
    plt.plot(data, 'b.', label=None)
    plt.plot(data, 'b', linewidth=1, label="Bonus/Malus" if fig_num in [1] else None )
    plt.axhline(0, color='black', linewidth=1)
    plt.axvline(0, color='black', linewidth=1)
    
    ax = plt.gca()
    ax.grid(True, color='b')
    plt.ylabel("Nombre de points")
    plt.xlabel("Officiels manquants ou en plus")

#ax.spines['left'].set_position('zero')
#ax.spines['bottom'].set_position('zero')

arrow = dict(facecolor='black', arrowstyle='->', connectionstyle='arc3, rad=.2')

if fig_num == 1:
    ax.annotate('30 nageurs, 3 officiels\n0 pts', xy=(0,0), xytext=(1,-5), arrowprops=arrow)
    ax.annotate('10 nageurs, 3 officiels\n2 supplémentaires\n4 pts', xy=(2,4), xytext=(-2,5), arrowprops=arrow)
    ax.annotate('30 nageurs, 1 officiels\n2 manquants\n-8 pts', xy=(-2,-8), xytext=(1,-10), arrowprops=arrow)

if fig_num in range(2,4):
    if fig_num == 2:
        clubs = [("DAUPHINS GRASSE", "r+"),
                 ("STADE ST-LAURENT-DU-VAR NAT", "gx"),
                 ("NICE LA SEMEUSE", "y.")]
    else:
        clubs = [("CN ANTIBES", "r+"),
                 ("OLYMPIC NICE NATATION", "gx"),
                 ("AS MONACO NATATION", "y.")]

    stats["required"] = stats["Participations"] // 10
    stats["x"] = stats["Officiels"] - stats["required"]
    stats["x"] += random.random()/10
    stats["Points"] += random.random()/10

    for club, color in clubs:
        club_stats = stats[stats.Club == club]
        print(club_stats[["Date", "Participations", "required", "Officiels", "Points", "x"]])
        plt.plot(club_stats["x"], club_stats["Points"], color, label=club)


def num_points(participations, officiels, depart, equipe):
    """
    Returns number of points for nageurs for each 'x' (Officiels +-)
    :param participations: number of nageurs
    :param officiels: number of officiels
    :param depart: True if Departement
    :param equipe: True if equipe
    :return: (x, points)
    """
    required = get_x(participations)
    x = officiels - required

    # needed = (Num of A/B, Total num)
    if equipe:
        participations //= 4
        if participations <= 1:
            needed = participations
        else:
            needed = min(3, participations)
    else:
        if participations <= 1:
            needed = 0
        elif participations <= 10:
            needed = 1
        elif participations <= 20 or not depart:
            needed = 2
        else:
            needed = 3

    if not depart and officiels > 5:
        officiels = 5

    if officiels < needed:
        missing = needed - officiels
        pts = missing * -4
    else:
        extra = officiels - needed
        pts = extra * 2

    return x, pts


if fig_num in (4, 5, 6, 7):
    index_start, index_end = -5, 6
    data = {}

    depart = fig_num in (4, 6)
    equipe = fig_num in (6, 7)

    if equipe:
        series = (("1 équipe",  4),
                  ("2 équipes", 8),
                  ("3 équipes", 12),
                  ("4 équipes", 16),
                  ("5 équipes", 20),
                  ("6 équipes", 24))
    else:
        series = (("1-9 nageurs", 2),
                  ("10-19 nageurs", 10),
                  ("20-29 nageurs", 20),
                  ("30-39 nageurs", 30),
                  ("40-49 nageurs", 40),
                  ("50-59 nageurs", 50))

    for label, participations in series:
        data[label] = [None] * (index_end - index_start + 1)
        for officiels in range(8):
            x, pts = num_points(participations, officiels, depart=depart, equipe=equipe)
            if x in range(index_start, index_end + 1):
                data[label][x-index_start] = pts

    data_figure = pd.DataFrame(data, index=range(index_start, index_end+1))
    print(data_figure)

    plt.plot(data_figure)
    data_figure.plot(kind='line', ax=ax)

    if fig_num == 6:
        ax.annotate('Clubs avec\n2, 3 ou 4 équipes\ndésavantagés', xy=(2, 2), xytext=(2, -5), arrowprops=arrow)

list_clubs = {"CNS VALLAURIS": "Vallauris", 
              "CN ANTIBES": "Antibes", 
              "DAUPHINS GRASSE": "Grasse", 
              "CN CANNES": "Cannes", 
              "STADE ST-LAURENT-DU-VAR NAT": "St Laurent",
              "AS MONACO NATATION": "Monaco", 
              "OLYMPIC NICE NATATION": "ONN", 
              "CN MENTON": "Menton", 
              "CARROS NATATION": "Carros"} 

if fig_num in (8, 9, 10):
    officiels_df = stats[stats['Club'].isin(list_clubs.keys())]

    if fig_num == 8:
      officiels_df = officiels_df[officiels_df["Structure"] == "Départemental"]
      max_nageurs = 21
    elif fig_num == 9:
      officiels_df = officiels_df[officiels_df["Structure"] == "Régional"]
      max_nageurs = 11
    elif fig_num == 10:
      officiels_df = officiels_df[officiels_df["Par Equipe"] == 1]
      max_nageurs = 12

    officiels_df.replace(to_replace=list_clubs, inplace=True)
    num_reunions = len(officiels_df.groupby(['Compétition', 'Réunion']).groups.keys())

    officiels_df = officiels_df.groupby(['Club'])['Participations', 'Engagements', 'Officiels'].sum()
    officiels_df.sort_values(by="Participations", inplace=True)
    officiels_df["Participations par réunion"] = officiels_df["Participations"] * 1.0 / num_reunions
    officiels_df["Nageurs par officiels"] = officiels_df["Participations"] * 1.0 / officiels_df["Officiels"]
    officiels_df["Engagements par officiels"] = officiels_df["Engagements"] * 1.0 / officiels_df["Officiels"]
    officiels_df = officiels_df[['Participations par réunion', 'Nageurs par officiels']]
    subplots = officiels_df.plot.bar(subplots=True, sharex=True)

    ax = plt.gca()
    subplots[0].axhline(max_nageurs, color='red')
    subplots[1].axhline(8, color='red')

    ax.grid(True, color='c')
    ax.set_xticklabels(officiels_df.index, rotation=45)
    plt.gcf().subplots_adjust(bottom=0.15)

if fig_num in (11, 12, 13):

    if fig_num == 11:
      officiels_df = stats[stats["Niveau"] == "Départemental"]
      max_nageurs = 21
    elif fig_num == 12:
      officiels_df = stats[stats["Niveau"] == "Régional"]
      max_nageurs = 11
    elif fig_num == 13:
      officiels_df = stats[stats["Par Equipe"] == 1]
      max_nageurs = 12

    competitions_df = officiels_df.groupby(['Compétition', 'Réunion'])
    competitions_df = competitions_df.agg({'Participations': np.sum,
                                           'Engagements': np.sum,
                                           'Officiels': np.sum,
                                           'Lignes': lambda x: x.iloc[0],
                                           'Longueur': lambda x: x.iloc[0],
                                           'Niveau': lambda x: x.iloc[0],
                                           'Par Equipe': lambda x: x.iloc[0]})
    competitions_df['Officiels voulus'] = competitions_df['Lignes'] * 3 + 9
    competitions_df = competitions_df[["Niveau", "Participations", "Engagements", "Officiels", "Lignes", "Longueur", "Officiels voulus"]]

    by_line = competitions_df.groupby(["Lignes"])["Participations", "Officiels", "Officiels voulus"].mean()
    by_line['Nageurs par officiels'] = by_line['Participations']  / by_line['Officiels']
    by_line['Nageurs voulus par officiels'] = by_line['Participations'] / by_line['Officiels voulus']
    by_line[["Officiels", "Officiels voulus", "Nageurs par officiels", "Nageurs voulus par officiels"]].plot.bar()
    print(by_line)
    ax = plt.gca()
    ax.grid(True, color='b')
    ax.axhline(by_line["Nageurs voulus par officiels"].mean(), color="red")
    

if fig_num not in (8, 9, 10, 11, 12, 13):
    legend = plt.legend(loc=4, ncol=2, prop={'size': 7}, fancybox=True, borderaxespad=0.)
    legend.get_frame().set_alpha(.4)

plt.show()




