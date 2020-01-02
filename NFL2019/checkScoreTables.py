import os
os.chdir(os.path.dirname(__file__))

import pandas as pd

teamsheetpath = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'Score Tables')

for f in os.listdir(teamsheetpath):

    team = f.split('.')[0]
    print('Checking ' + team)

    df = pd.read_csv(os.path.join(teamsheetpath, f), index_col = 0)

    if df.shape[0] != 16:
        print('Wrong number of games')

    if team in list(df['OPP']):
        print('You done it again')

    n_home = df['VENUE'].value_counts()[team]
    if n_home != 8:
        print('%d HOME GAMES!'%(n_home))

    print('\n')

print('Done Checking')