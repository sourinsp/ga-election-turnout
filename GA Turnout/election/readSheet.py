
import xlrd
import pandas as pd
from openpyxl import load_workbook
import xlsxwriter
import csv
from openpyxl.utils.cell import col

'''
Created on Aug 13, 2020    

@author: Souri
'''

def main():

    base_file = ("/Users/sourinpaturi/Dropbox/Projects/GA Turnout/base files/ElectionData2016.xlsx")
    
    new_file = ("/Users/sourinpaturi/Dropbox/Projects/GA Turnout/base files/ElectionData2018.xlsx")
 
    other_file = ("/Users/sourinpaturi/Dropbox/Projects/GA Turnout/base files/ElectionData2014.xlsx")
    
    #turnout = pd.read_excel(base_file, dtype = str)
    #print(turnout)
    
    #df = turnout.drop(['COUNTY NAME', 'COUNTY CODE'], axis=1)
    #print(df)    

    #print(df['BM%'])
    
    print(demoAvg(base_file, 'WM'))
    print(demoTotal(base_file, 'WM'))
    print(demoTotal(base_file, 'WF'))
    print(raceTotal(base_file,'H'))
    
    
def demoAvg(year_sheet, type):
    
    demoAvg = 0.0
    
    turnout = pd.read_excel(year_sheet, dtype = str)
    
    df = turnout.drop([col for col in turnout.columns if not "%" in col], axis = 1)

    # Changes the type to the searchable turnout percentage
    searchable = type+'%'
    
    # Converts the values to be floats and sums them all up averages
    df[searchable] = pd.to_numeric(df[searchable])
    sum = df[searchable].to_numpy().sum()
        
    mean = sum/159
    demoAvg = round(mean,2)
    
    return demoAvg


def demoTotal(year_sheet, type):
    
    demoTotal = 0.0
    
    turnout = pd.read_excel(year_sheet, dtype = str)
    
    df = turnout.drop([col for col in turnout.columns if "%" in col], axis = 1)
    
    num_reg = type+'R'
    num_vote = type+'V'
    
    df[num_reg] = pd.to_numeric(df[num_reg])
    reg_sum = df[num_reg].to_numpy().sum()
    
    
    df[num_vote] = pd.to_numeric(df[num_vote])
    vote_sum = df[num_vote].to_numpy().sum()
    
    total = (vote_sum / reg_sum) * 100.0
    demoTotal = round(total, 2)
    
    return demoTotal
    
def countyTotal(year_sheet, county):
    
    countyTotal = 0.0
    
    turnout = pd.read_excel(year_sheet, dtype = str)
    
    rowNum = turnout[turnout["COUNTY NAME"] == county].index[0]
    countyTotal = turnout.at[rowNum,'T%']
      
    return countyTotal


def raceTotal(year_sheet, type):
    
    raceTotal = 0.0
    
    turnout = pd.read_excel(year_sheet, dtype = str)
    
    df = turnout.drop([col for col in turnout.columns if "%" in col], axis = 1)
    
    
    # Male Voters
    mal_reg = type +'MR'
    mal_vote = type +'MV'
    
    df[mal_reg] = pd.to_numeric(df[mal_reg])
    mal_reg_sum = df[mal_reg].to_numpy().sum()
    
    df[mal_vote] = pd.to_numeric(df[mal_vote])
    mal_vote_sum = df[mal_vote].to_numpy().sum()
    
    # Female Voters
    fem_reg = type +'FR'
    fem_vote = type +'FV'
    
    df[fem_reg] = pd.to_numeric(df[fem_reg])
    fem_reg_sum = df[fem_reg].to_numpy().sum()
    
    df[fem_vote] = pd.to_numeric(df[fem_vote])
    fem_vote_sum = df[fem_vote].to_numpy().sum()
    
    # Unkown Voters
    unk_reg = type +'UR'
    unk_vote = type +'UV'
    
    df[unk_reg] = pd.to_numeric(df[unk_reg])
    unk_reg_sum = df[unk_reg].to_numpy().sum()
    
    df[unk_vote] = pd.to_numeric(df[unk_vote])
    unk_vote_sum = df[unk_vote].to_numpy().sum()
    
    # Add totals and averaging
    
    total_reg = mal_reg_sum + fem_reg_sum + unk_reg_sum
    
    total_vote = mal_vote_sum + fem_vote_sum + unk_vote_sum
    
    subTotal = (total_vote / total_reg) * 100.0

    raceTotal = round(subTotal, 2)
    
    return raceTotal

def genderTotal(year_sheet, type):  
    
    print(0)
    
    
    
    

def exactEntry(year_sheet, countyName, type):
    
    turnout = pd.read_excel(year_sheet, dtype = str)
    
    
    headers = turnout.iloc[0]
    new_df  = pd.DataFrame(turnout.values[1:], columns=headers)
    
    print(new_df)

    
    
    
    
if __name__ == '__main__':
    main()