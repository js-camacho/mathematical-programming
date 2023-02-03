# -*- coding: utf-8 -*-
"""
Created on Thu Jul  2 16:27:29 2020

@author: johan
"""
#%% Importing Gurobi Shell and other libraries

import gurobipy as gb
import pandas as pd
import collections
        
#%% Model Data

# Importing data from excel file

df_cost = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = 'Cost')
df_hardness = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = 'Hardness')
df_scalars = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = 'Scalars')


# Creating dictionaries with the data

dic_cost = {}

for index, row in df_cost.iterrows():
    dic_cost[row[0],row[1]] = row[2]
    
dic_hardness = {}

for index, row in df_hardness.iterrows():
    dic_hardness[row[0]] = [row[1], row[2]]
    
scalars = {}

for index, row in df_scalars.iterrows():
    scalars[row[0]] = row[1]


# Creating model parameters

oils_months, costs = gb.multidict(dic_cost)
oils, hardness, types_ = gb.multidict(dic_hardness)
months = df_cost.MONTH.unique()

types = collections.defaultdict(list)
for k, v in types_.items():
    types[v].append(k)

del df_cost, df_hardness, df_scalars, dic_cost, dic_hardness, types_

#%% Model Formulation

#-------------- Model Creation

fm1 = gb.Model('Food Manufacture I')

#-------------- Variables Creation

refine = fm1.addVars(oils_months, name = 'refine', obj = scalars['PRICE'])
buy = fm1.addVars(oils_months, name = 'buy', obj = costs)
inv = fm1.addVars(oils_months, name = 'inventory', obj = scalars['STORECOST'], ub = scalars['STORAGE'])

fm1.ModelSense = gb.GRB.MAXIMIZE

#------------- Constraints Creation

# 1. The inventory at the end of a month depends of the tons of raw oils refined and bought in that month

for o in oils:
    # 1.1 Initial inventory
    fm1.addConstr(inv[o,months[0]] == 500 - refine[o,months[0]] + buy[o,months[0]])
    # 1.1 Final inventory
    fm1.addConstr(500 == inv[o,months[-2]] - refine[o,months[-1]] + buy[o,months[-1]])
    
    for j in range(1, len(months)-1):
        m = months[j]
        m1 = months[j-1]
        fm1.addConstr(inv[o,m] == inv[o,m1] - refine[o,m] + buy[o,m])

# 2. There are maximum refining capacities for each type of raw oil each month

for key, t in types.items():
    if key == 'V':
        capacity = scalars['CAPACITY_V']
    else:
        capacity = scalars['CAPACITY_N']
        
    fm1.addConstrs((gb.quicksum([refine[o,m] for o in t]) <= capacity for m in months))

# 3. The maximum storage capacity cannot be surpassed

# This constraint is included as an upper bound of the <inv> variables
#fm1.addConstrs((inv[i] <= scalars['STORAGE'] for i in oils_months))

# 4. There are hardness bounds for the final product linearly dependent of the individual hardness of each raw oil used

fm1.addConstrs((gb.quicksum((hardness[o] - scalars['HL'])*refine[o,m] for o in oils) >= 0 for m in months))
fm1.addConstrs((gb.quicksum((hardness[o] - scalars['HU'])*refine[o,m] for o in oils) <= 0 for m in months))

#-------------- Model Execution

fm1.setParam('OutputFlag',0)    # Turns off the Optimization Details sheet print after the model.optimize() call
fm1.optimize()

#%% Results Report

# Reporting variables values
price = scalars['PRICE']
store = scalars['STORECOST']
for m in months:
    print(f'------------------\n{m}\n------------------')
    for o in oils:
        val = round(refine[o,m].x,2)
        if val > 0:
            print(f'{o}: Refine -> {val}')
            #print(f'{o}: Refine -> {val}, Profit: ${round(val*price,1)}')
        val = round(buy[o,m].x,2)
        if val > 0:
            #cost = costs[o,m]
            print(f'{o}: Buy -> {val}')
            #print(f'{o}: Buy -> {val}, Profit: ${round(val*cost,1)}')
        val = round(inv[o,m].x,2)
        if val > 0:
            print(f'{o}: Store -> {val}')
            #print(f'{o}: Store -> {val}, Profit: ${round(val*store,1)}')
        
# Reporting objective function value  

print('***************************************')
print(f'Objective function value: ${round(fm1.objval,1) - 12500}')
print('$12.500 were deducted of \nstore costs of the last month')
print('***************************************')


#%% End of file