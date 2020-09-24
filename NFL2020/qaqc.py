import os
os.chdir(os.path.dirname(__file__))

import pandas as pd

teamsheetpath = os.path.join(os.path.split(os.path.abspath(__file__))[0], 'Score Tables')

n_errors = 0

for f in os.listdir(teamsheetpath):

    team = f.split('.')[0]
    print('Checking ' + team)
    df = pd.read_csv(os.path.join(teamsheetpath, f), index_col = 0)

    for opp in df['OPP']:
        opp_df = pd.read_csv(os.path.join(teamsheetpath, opp + '.csv'), index_col = 0)

        if (df.query('OPP == @opp').T.iloc[2:10].values != opp_df.query('OPP == @team').T.iloc[10:18].values).any():
            print('Mismatch in ' + team + ' vs ' + opp)
            n_errors += 1
print('Done Checking')
print('%d Errors Found'%(n_errors))