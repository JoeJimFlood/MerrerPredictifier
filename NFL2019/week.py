import os
os.chdir(os.path.dirname(__file__))

import pandas as pd
import xlsxwriter
import sys
import time
import collections
import matplotlib.pyplot as plt
from math import log2

import matchup
import ranking

week_timer = time.time()

week_number = 12

matchups = collections.OrderedDict()

matchups['Thursday Night'] = [('HOU', 'IND')]
matchups['Sunday Morning'] = [('BUF', 'DEN'),
                              ('CHI', 'NYG'),
                              ('CIN', 'PIT'),
                              ('CLE', 'MIA'),
                              ('ATL', 'TB'),
                              ('NO', 'CAR'),
                              ('PHI', 'SEA'),
                              ('WAS', 'DET'),
                              ('NYJ', 'OAK')]
matchups['Sunday Afternoon'] = [('TEN', 'JAX'),
                                ('NE', 'DAL')]
matchups['Sunday Night'] = [('SF', 'GB')]
matchups['Monday Night'] = [('LAR', 'BAL')]

def rgb2hex(r, g, b):
    r_hex = hex(r)[-2:].replace('x', '0')
    g_hex = hex(g)[-2:].replace('x', '0')
    b_hex = hex(b)[-2:].replace('x', '0')
    return '#' + r_hex + g_hex + b_hex

location = os.path.split(__file__)[0]
stadium_file = location + '/StadiumLocs.csv'
output_file = location + '/Weekly Forecasts/Week' + str(week_number) + '.xlsx'
output_fig = location + '/Weekly Forecasts/Week' + str(week_number) + '.png'
stadium_file = location + '/StadiumLocs.csv'
stadiums = pd.read_csv(stadium_file, index_col = 0)

rankings = ranking.rank(os.path.join(location, 'Score Tables'), week_number)

n_games = 0
for day in matchups:
    n_games += len(matchups[day])

colors = {}
team_formats = {}
color_df = pd.DataFrame.from_csv(location + '/colors.csv')
teams = list(color_df.index)
for team in teams:
    primary = rgb2hex(int(color_df.loc[team, 'R1']), int(color_df.loc[team, 'G1']), int(color_df.loc[team, 'B1']))
    secondary = rgb2hex(int(color_df.loc[team, 'R2']), int(color_df.loc[team, 'G2']), int(color_df.loc[team, 'B2']))
    colors[team] = (primary, secondary)

name_map = pd.DataFrame.from_csv(location + '/names.csv')['NAME'].to_dict()

plt.figure(figsize = (24, 24), dpi = 96)
plt.title('Week ' + str(week_number))
counter = 0

week_book = xlsxwriter.Workbook(output_file)
header_format = week_book.add_format({'align': 'center', 'bold': True, 'bottom': True})
index_format = week_book.add_format({'align': 'right', 'bold': True})
score_format = week_book.add_format({'num_format': '#0', 'align': 'right'})
mean_format = week_book.add_format({'num_format': '#0.0', 'align': 'right'})
percent_format = week_book.add_format({'num_format': '#0%', 'align': 'right'})
merged_format = week_book.add_format({'num_format': '#0.00', 'align': 'center'})
merged_format2 = week_book.add_format({'num_format': '0.000', 'align': 'center'})
for team in teams:
    team_formats[team] = week_book.add_format({'align': 'center', 'bold': True, 'border': True,
                                                'bg_color': colors[team][0], 'font_color': colors[team][1]})

for game_time in matchups:
        
    sheet = week_book.add_worksheet(game_time)
    sheet.write_string(1, 0, 'City', index_format)
    sheet.write_string(2, 0, 'Quality', index_format)
    sheet.write_string(3, 0, 'Entropy', index_format)
    sheet.write_string(4, 0, 'Hype', index_format)
    sheet.write_string(5, 0, 'Chance of Winning', index_format)
    sheet.write_string(6, 0, 'Expected Score', index_format)
    for i in range(1, 20):
        sheet.write_string(6+i, 0, str(5*i) + 'th Percentile Score', index_format)
    sheet.freeze_panes(1, 1)
    games = matchups[game_time]

    for i in range(len(games)):
        home = games[i][0]
        away = games[i][1]

        try:
            venue = games[i][2]
        except IndexError:
            venue = games[i][0]
        stadium = stadiums.loc[venue, 'Venue']
        city = stadiums.loc[venue, 'City']
        state = stadiums.loc[venue, 'State']

        homecol = 3 * i + 1
        awaycol = 3 * i + 2
        sheet.write_string(0, homecol, name_map[home], team_formats[home])
        sheet.write_string(0, awaycol, name_map[away], team_formats[away])
            
        if len(games[i]) == 3:
            results = matchup.matchup(home, away, games[i][2])
        else:
            results = matchup.matchup(home, away)
        
        probwin = results['ProbWin']

        #Calculate hype
        home_ranking = rankings.loc[home, 'Quantile']
        away_ranking = rankings.loc[away, 'Quantile']
        ranking_factor = (home_ranking + away_ranking) / 2
        hwin = probwin[home]
        awin = probwin[away]
        entropy = -hwin*log2(hwin) - awin*log2(awin)
        hype = 100 * ranking_factor * entropy

        sheet.merge_range(1, homecol, 1, awaycol, city, merged_format)
        sheet.merge_range(2, homecol, 2, awaycol, ranking_factor, merged_format2)
        sheet.merge_range(3, homecol, 3, awaycol, entropy, merged_format2)
        sheet.merge_range(4, homecol, 4, awaycol, hype, merged_format)
        sheet.write_number(5, homecol, hwin, percent_format)
        sheet.write_number(5, awaycol, awin, percent_format)
        home_dist = results['Scores'][home]
        away_dist = results['Scores'][away]
        sheet.write_number(6, homecol, home_dist['mean'], mean_format)
        sheet.write_number(6, awaycol, away_dist['mean'], mean_format)
        for i in range(1, 20):
            sheet.write_number(6+i, homecol, home_dist[str(5*i)+'%'], score_format)
            sheet.write_number(6+i, awaycol, away_dist[str(5*i)+'%'], score_format)

        sheet.set_column(0, 0, 20)
        sheet.set_column(1, awaycol, 12)
        for i in range(3, awaycol, 3):
            sheet.set_column(i, i, 0.5)

        if i != len(games) - 1:
            sheet.write_string(0, 3 * i + 3, ' ')

        counter += 1
        hwin = probwin[home]
        awin = probwin[away]
        draw = 1 - hwin - awin

        if counter == 5:
            counter += 1

        if counter == 10:
            plt.savefig(output_fig.replace('.png', '-1.png'))
            plt.clf()
            plt.close()

            counter = 1
            plt.figure(figsize = (24, 24), dpi = 96)
            plt.title('Week ' + str(week_number))

        plt.subplot(3, 3, counter)
        labels = [home, away]
        values = [hwin, awin]
        c = [colors[home][0], colors[away][0]]
        ex = 0.05
        explode = [ex, ex]
        plt.pie(values,
                colors = c,
                labels = labels,
                explode = explode,
                autopct='%.0f%%',
                startangle = 90,
                labeldistance = 1,
                textprops = {'backgroundcolor': '#ffffff', 'ha': 'center', 'va': 'center', 'fontsize': 18})
        plt.title(name_map[home] + ' vs ' + name_map[away] + '\n' + stadium + '\n' + city + ', ' + state + '\n' + 'Hype: ' + str(int(round(hype, 0))), size = 18)
        plt.axis('equal')

week_book.close()

plt.savefig(output_fig.replace('.png', '-2.png'))
#plt.savefig(output_fig)

print('Week ' + str(week_number) + ' predictions calculated in ' + str(round((time.time() - week_timer) / 60, 2)) + ' minutes')