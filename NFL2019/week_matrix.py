import os
os.chdir(os.path.dirname(__file__))

import pandas as pd
import xlsxwriter
import sys
import time
import collections
import os
import matplotlib.pyplot as plt

import matchup

week_timer = time.time()

week_number = 'WCMatrix'

matchups = collections.OrderedDict()
matchups['BAL'] = [('BAL', 'KC'),
                   ('BAL', 'NE'),
                   ('BAL', 'HOU'),
                   ('BAL', 'BUF'),
                   ('BAL', 'TEN'),
                   ('BAL', 'SF', 'MIA'),
                   ('BAL', 'GB', 'MIA'),
                   ('BAL', 'NO', 'MIA'),
                   ('BAL', 'PHI', 'MIA'),
                   ('BAL', 'SEA', 'MIA'),
                   ('BAL', 'MIN', 'MIA')]
matchups['KC'] = [('KC', 'NE'),
                  ('KC', 'HOU'),
                  ('KC', 'BUF'),
                  ('KC', 'TEN'),
                  ('KC', 'SF', 'MIA'),
                  ('KC', 'GB', 'MIA'),
                  ('KC', 'NO', 'MIA'),
                  ('KC', 'PHI', 'MIA'),
                  ('KC', 'SEA', 'MIA'),
                  ('KC', 'MIN', 'MIA')]
matchups['NE'] = [('NE', 'HOU'),
                  ('NE', 'BUF'),
                  ('NE', 'TEN'),
                  ('NE', 'SF', 'MIA'),
                  ('NE', 'GB', 'MIA'),
                  ('NE', 'NO', 'MIA'),
                  ('NE', 'PHI', 'MIA'),
                  ('NE', 'SEA', 'MIA'),
                  ('NE', 'MIN', 'MIA')]
matchups['HOU'] = [('HOU', 'BUF'),
                   ('HOU', 'TEN'),
                   ('HOU', 'SF', 'MIA'),
                   ('HOU', 'GB', 'MIA'),
                   ('HOU', 'NO', 'MIA'),
                   ('HOU', 'PHI', 'MIA'),
                   ('HOU', 'SEA', 'MIA'),
                   ('HOU', 'MIN', 'MIA')]
matchups['BUF'] = [('BUF', 'TEN'),
                   ('BUF', 'SF', 'MIA'),
                   ('BUF', 'GB', 'MIA'),
                   ('BUF', 'NO', 'MIA'),
                   ('BUF', 'PHI', 'MIA'),
                   ('BUF', 'SEA', 'MIA'),
                   ('BUF', 'MIN', 'MIA')]
matchups['TEN'] = [('TEN', 'SF', 'MIA'),
                   ('TEN', 'GB', 'MIA'),
                   ('TEN', 'NO', 'MIA'),
                   ('TEN', 'PHI', 'MIA'),
                   ('TEN', 'SEA', 'MIA'),
                   ('TEN', 'MIN', 'MIA')]
matchups['SF'] = [('SF', 'GB'),
                  ('SF', 'NO'),
                  ('SF', 'PHI'),
                  ('SF', 'SEA'),
                  ('SF', 'MIN')]
matchups['GB'] = [('GB', 'NO'),
                  ('GB', 'PHI'),
                  ('GB', 'SEA'),
                  ('GB', 'MIN')]
matchups['NO'] = [('NO', 'PHI'),
                  ('NO', 'SEA'),
                  ('NO', 'MIN')]
matchups['PHI'] = [('PHI', 'SEA'),
                   ('PHI', 'MIN')]
matchups['SEA'] = [('SEA', 'MIN')]

def rgb2hex(r, g, b):
    r_hex = hex(r)[-2:].replace('x', '0')
    g_hex = hex(g)[-2:].replace('x', '0')
    b_hex = hex(b)[-2:].replace('x', '0')
    return '#' + r_hex + g_hex + b_hex

location = os.getcwd().replace('\\', '/')
output_file = location + '/Weekly Forecasts/Week' + str(week_number) + '.xlsx'
output_fig = location + '/Weekly Forecasts/Week' + str(week_number) + '.png'

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

plt.figure(figsize = (18, 18), dpi = 96)
plt.title('Week ' + str(week_number))
counter = 0

week_book = xlsxwriter.Workbook(output_file)
header_format = week_book.add_format({'align': 'center', 'bold': True, 'bottom': True})
index_format = week_book.add_format({'align': 'right', 'bold': True})
score_format = week_book.add_format({'num_format': '#0', 'align': 'right'})
mean_format = week_book.add_format({'num_format': '#0.0', 'align': 'right'})
percent_format = week_book.add_format({'num_format': '#0%', 'align': 'right'})
for team in teams:
    team_formats[team] = week_book.add_format({'align': 'center', 'bold': True, 'border': True,
                                                'bg_color': colors[team][0], 'font_color': colors[team][1]})

for game_time in matchups:
        
    sheet = week_book.add_worksheet(game_time)
    sheet.write_string(1, 0, 'Chance of Winning', index_format)
    sheet.write_string(2, 0, 'Expected Score', index_format)
    for i in range(1, 20):
        sheet.write_string(2+i, 0, str(5*i) + 'th Percentile Score', index_format)
    sheet.freeze_panes(0, 1)
    games = matchups[game_time]

    for i in range(len(games)):
        home = games[i][0]
        away = games[i][1]
        homecol = 3 * i + 1
        awaycol = 3 * i + 2
        sheet.write_string(0, homecol, name_map[home], team_formats[home])
        sheet.write_string(0, awaycol, name_map[away], team_formats[away])
            
        if len(games[i]) == 3:
            results = matchup.matchup(home, away, games[i][2])
        else:
            results = matchup.matchup(home, away)
        probwin = results['ProbWin']
        sheet.write_number(1, homecol, probwin[home], percent_format)
        sheet.write_number(1, awaycol, probwin[away], percent_format)
        home_dist = results['Scores'][home]
        away_dist = results['Scores'][away]
        sheet.write_number(2, homecol, home_dist['mean'], mean_format)
        sheet.write_number(2, awaycol, away_dist['mean'], mean_format)
        for i in range(1, 20):
            sheet.write_number(2+i, homecol, home_dist[str(5*i)+'%'], score_format)
            sheet.write_number(2+i, awaycol, away_dist[str(5*i)+'%'], score_format)

        sheet.set_column(0, 0, 20)
        sheet.set_column(1, awaycol, 12)
        for i in range(3, awaycol, 3):
            sheet.set_column(i, i, 0.5)

        if i != len(games) - 1:
            sheet.write_string(0, 3 * i + 3, ' ')

        #counter += 1
        #hwin = probwin[home]
        #awin = probwin[away]
        #draw = 1 - hwin - awin

        #plt.subplot(5, 6, counter)
        #labels = [home, away]
        #values = [hwin, awin]
        #c = [colors[home][0], colors[away][0]]
        #ex = 0.05
        #explode = [ex, ex]
        #plt.pie(values,
        #        colors = c,
        #        labels = labels,
        #        explode = explode,
        #        autopct='%.0f%%',
        #        startangle = 90,
        #        labeldistance = 1,
        #        textprops = {'backgroundcolor': '#ffffff', 'ha': 'center', 'va': 'center'})
        #plt.title(name_map[home] + ' vs ' + name_map[away], size = 18)
        #plt.axis('equal')

        time.sleep(5)

week_book.close()

plt.savefig(output_fig)

print('Week ' + str(week_number) + ' predictions calculated in ' + str(round((time.time() - week_timer) / 60, 2)) + ' minutes')