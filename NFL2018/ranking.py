import pandas as pd
import numpy as np
import os
from scipy.stats.distributions import norm

def rank(wd, roundno):
    '''
    Rank the teams!
    '''
    #Create list of types of scores
    scores_for = ['TDF', 'PAT1FS', 'PAT2FS', 'FGF', 'SFF', 'D2CF']
    scores_against = ['TDA', 'PAT1AS', 'PAT2AS', 'FGA', 'SFA', 'D2CA']

    scores = np.array([6, 1, 2, 3, 2, 2]) #Number of points in each score type

    #Create dictionaries mapping each team to their average points for and against
    pf = {}
    pa = {}
    for team in os.listdir(wd):
        data = pd.read_csv(os.path.join(wd, team))
        pf[team[:-4]] = np.dot(data[scores_for], scores).mean()
        pa[team[:-4]] = np.dot(data[scores_against], scores).mean()

    #Create data frame for results
    results = pd.DataFrame(index = pf.keys(), columns = ['Attack', 'Defense', 'Overall'])

    #Calculate each teams residual points for, against, and differential
    for team in os.listdir(wd):
        data = pd.read_csv(os.path.join(wd, team))
        data['For'] = np.dot(data[scores_for], scores)
        data['Against'] = np.dot(data[scores_against], scores)
        data['OppFor'] = data['OPP'].map(pf)
        data['OppAgainst'] = data['OPP'].map(pa)
        data['Attack'] = data['For'] - data['OppAgainst']
        data['Defense'] = data['Against'] - data['OppFor']
        data['Overall'] = data['Attack'] - data['Defense']
    
        results.loc[team[:-4]] = data[results.columns].mean() #Add to results table

    #Standardize results and compute theoretical quantiles
    results['Standardized'] = (results['Overall'] - results['Overall'].mean())/results['Overall'].std()
    results['Quantile'] = norm.cdf(results['Standardized'].astype(float))

    #Write results to file
    outfile = os.path.join(os.path.split(wd)[0], 'Rankings', 'RankingsRound{}.csv'.format(roundno))
    results.sort_values('Overall', ascending = False).to_csv(outfile)
    return results
    #Popen(outfile, shell = True)