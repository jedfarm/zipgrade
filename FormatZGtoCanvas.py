# -*- coding: utf-8 -*-
"""
Created on Fri Oct 20 08:32:40 2017
Takes a .csv file from ZipGrade and using a roster ( in a .csv format)  to compare 
against, creates a .csv file ready to be upoaded into canvas.

@author: jedfarm
"""
import os
import pandas as pd

# The .csv file from zipgrade and the roster (also a .csv file) must be in the same folder
#Change the string between quotation marks in the following line to the given folder
os.chdir('/Users/macbookpro/Downloads/ZIPGRADE')

# Change the string in quotation marks for the roster file of interest
filename_roster = "Roster-18_FA-PHY-1025-15432.csv"

# Change the string in quotation marks for the file with the grades from zipgrade.
filename_quiz = "quiz-E2-standard20180510.csv"


filename_out = filename_quiz.split('-')[1] + "_toCanvas.csv"
roster = pd.read_csv(filename_roster)
quiz = pd.read_csv(filename_quiz)
quiz_name = quiz['Quiz Name'][1]

# 'External Id' must be there, the next column is the one that contains the grades
# of interest and could be changed upon necesity.
cols_of_interest = ['External Id', 'Num Correct']

quiz = quiz[cols_of_interest]
quiz.rename(columns={'Num Correct': quiz_name}, inplace=True)

new_df = roster.merge(quiz, left_on='SIS Login ID', right_on='External Id', how='left')
del new_df['External Id']
new_df.to_csv(filename_out, encoding='utf-8', index=False)
