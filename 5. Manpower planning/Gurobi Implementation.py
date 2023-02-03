# -*- coding: utf-8 -*-
"""
Created on Thu Jul 23 11:27:37 2020

@author: johan
"""
#*******************************
# MANPOWER PLANNING
#*******************************

#%% Importing Gurobi Shell and other libraries

import gurobipy as gb
import pandas as pd
from datetime import datetime


#%% Model Data

# Importing data from excel file
data = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = None)

# Creating dictionaries with the data

dic_skills = {}
for index, row in data['skills'].iterrows(): 
    dic_skills[row[0]] = (row[1], row[2], row[3], row[4], row[5])
    
dic_demand = {}
for index, row in data['demand'].iterrows(): 
    dic_demand[row[0],row[1]] = row[2]

years = list(data['demand'].YEAR.unique())
retrain_downgrade_years = []
for i in years:
    retrain_downgrade_years.append(('US -> SS',i))
    retrain_downgrade_years.append(('SS -> SK',i))
    retrain_downgrade_years.append(('SK -> SS',i))
    retrain_downgrade_years.append(('SK -> US',i))
    retrain_downgrade_years.append(('SS -> US',i))

# Creating model parameters
skills, initial, A, redCost, ovCost, stCost = gb.multidict(dic_skills) 
skills_years, B = gb.multidict(dic_demand)

del data, dic_skills, dic_demand, index, row, i

#%% Model Formulation

#-------------- Model Creation

model = gb.Model('Manpower Planning')

#-------------- Variables Creation
t = model.addVars(skills_years, obj = 0, vtype = gb.GRB.INTEGER, name='t')

u = model.addVars(skills_years, obj = 0, vtype = gb.GRB.INTEGER, name='u')
for i in years:
    for s in skills:
        # Upper bounds
        u[s,i].ub = A[s]
        
v = model.addVars(retrain_downgrade_years, obj = 0, vtype = gb.GRB.INTEGER, name='v')
for i in years:
    # Objective coefficients
    v['US -> SS',i].obj = 400
    v['SS -> SK',i].obj = 500
    # Upper bounds
    v['US -> SS',i].ub = 200
    
w = model.addVars(skills_years, obj = 0, vtype = gb.GRB.INTEGER, name='w')
for i in years:
    for s in skills:
        # Objective coefficients
        w[s,i].obj = redCost[s]
   
x = model.addVars(skills_years, obj = 0, vtype = gb.GRB.INTEGER, ub = 50, name='x')
for i in years:
    for s in skills:
        # Objective coefficients
        x[s,i].obj = stCost[s]
      
y = model.addVars(skills_years, obj = 0, vtype = gb.GRB.INTEGER, name='y')
for i in years:
    for s in skills:
        # Objective coefficients
        y[s,i].obj = ovCost[s]
        
model.ModelSense = gb.GRB.MINIMIZE

#------------- Constraints Creation

# 1.	Continuity: Workers at each time depend on the wastage, recruitment, retraining and redundancy
SK = 'Skilled'
SS = 'Semi-skilled'
US = 'Unskilled'
for year in range(len(years)):
    i = years[year]
    if i == 'Year 1':
        model.addConstr(t[SK,i] == 
                        0.95*initial[SK]+0.9*u[SK,i]+0.95*v['SS -> SK',i]-v['SK -> SS',i]-v['SK -> US',i]-w[SK,i])
        model.addConstr(t[SS,i] == 
                        0.95*initial[SS]+0.8*u[SS,i]+0.95*v['US -> SS',i]-v['SS -> SK',i]+0.5*v['SK -> SS',i]-v['SS -> US',i]-w[SS,i])
        model.addConstr(t[US,i] == 
                        0.9*initial[US]+0.75*u[US,i]-v['US -> SS',i]+0.5*v['SK -> US',i]+0.5*v['SS -> US',i]-w[US,i])
    else:
        j = years[year-1]
        model.addConstr(t[SK,i] == 
                        0.95*t[SK,j]+0.9*u[SK,i]+0.95*v['SS -> SK',i]-v['SK -> SS',i]-v['SK -> US',i]-w[SK,i])
        model.addConstr(t[SS,i] == 
                        0.95*t[SS,j]+0.8*u[SS,i]+0.95*v['US -> SS',i]-v['SS -> SK',i]+0.5*v['SK -> SS',i]-v['SS -> US',i]-w[SS,i])
        model.addConstr(t[US,i] == 
                        0.9*t[US,j]+0.75*u[US,i]-v['US -> SS',i]+0.5*v['SK -> US',i]+0.5*v['SS -> US',i]-w[US,i])

# 2.	Retraining Semi-skilled workers: The retraining of semi-skilled workers to make them skilled is limited 
#       to no more than one quarter of the skilled labour force at the time

model.addConstrs((v['SS -> SK',i] <= 0.25*t[SK,i] for i in years))

# 3.	Overmanning: There cannot be more than 150 workers more than needed at any year

model.addConstrs((gb.quicksum(y[s,i] for s in skills) <= 150 for i in years))

# 4.	Requirements: The number of workers required must be the workers available, the short time workers minus 
#       the overmanning workers

model.addConstrs((t[j] - y[j] - 0.5*x[j] == B[j] for j in skills_years))

#-------------- Model Execution

model.setParam('OutputFlag',0)    # Turns off the Optimization Details sheet print after the model.optimize() call
# Setting timer
begin = datetime.now()

model.optimize()

# Stopping timer
elapsed = datetime.now() - begin

#%% Results Report

# Reporting variables values
for i in years:
    print(f'---------------------------\n{i}\n---------------------------')
    print('Labour Force:------------')
    for s in skills:
        val = t[s,i].x
        if val > 0:
            print(f'{s} -> {val}')
            
    print('\nRecruitment:-------------')
    for s in skills:
        val = u[s,i].x
        if val > 0:
            print(f'{s} -> {val}')
            
    print('\nRetraining:--------------')
    val = v['US -> SS',i].x
    if val > 0:
        print(f'US -> SS -> {val}')
    val = v['SS -> SK',i].x
    if val > 0:
        print(f'SS -> SK -> {val}')  
        
    print('\nDowngrading:-------------')
    val = v['SK -> SS',i].x
    if val > 0:
        print(f'SK -> SS -> {val}')
    val = v['SK -> US',i].x
    if val > 0:
        print(f'SK -> US -> {val}')
    val = v['SS -> US',i].x
    if val > 0:
        print(f'SS -> US -> {val}')
    
    print('\nRedundancy:--------------')
    for s in skills:
        val = w[s,i].x
        if val > 0:
            print(f'{s} -> {val}')
            
    print('\nShort Time Working:------')
    for s in skills:
        val = x[s,i].x
        if val > 0:
            print(f'{s} -> {val}')
            
    print('\nOvermanning:-------------')
    for s in skills:
        val = y[s,i].x
        if val > 0:
            print(f'{s} -> {val}')

# Reporting objective function value
print('***************************************')
print(f'Objective function value: £{round(model.objval)}')
print(f'Time elapsed: {round(elapsed.microseconds/1000000,2)} seconds')
print('***************************************')

print('\nNOTE: The solution of the book: £498677 is obtained \
using continuous variables instead of integer variables')


#%% End of file
