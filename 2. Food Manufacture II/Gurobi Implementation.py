# -*- coding: utf-8 -*-
"""
Created on Tue Jul  7 16:25:06 2020

@author: johan
"""
#*******************************
# FOOD MANUFACTURE II
#*******************************
#%% Importing Gurobi Shell and other libraries

import gurobipy as gb
import pandas as pd
import collections
from datetime import datetime
        
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

fm2 = gb.Model('Food Manufacture II')

#-------------- Variables Creation

refine = fm2.addVars(oils_months, name = 'refine', obj = scalars['PRICE'])
buy = fm2.addVars(oils_months, name = 'buy', obj = costs)
inv = fm2.addVars(oils_months, name = 'inventory', obj = scalars['STORECOST'], ub = scalars['STORAGE'])
delta = fm2.addVars(oils_months, name = 'delta', obj = 0, vtype = gb.GRB.BINARY)

fm2.ModelSense = gb.GRB.MAXIMIZE

#------------- Constraints Creation

# 1. The inventory at the end of a month depends of the tons of raw oils refined and bought in that month

for o in oils:
    # 1.1 Initial inventory
    fm2.addConstr(inv[o,months[0]] == 500 - refine[o,months[0]] + buy[o,months[0]])
    # 1.1 Final inventory
    fm2.addConstr(500 == inv[o,months[-2]] - refine[o,months[-1]] + buy[o,months[-1]])
    
    for j in range(1, len(months)-1):
        m = months[j]
        m1 = months[j-1]
        fm2.addConstr(inv[o,m] == inv[o,m1] - refine[o,m] + buy[o,m])

# 2. There are maximum refining capacities for each type of raw oil each month

for key, t in types.items():
    if key == 'V':
        capacity = scalars['CAPACITY_V']
    else:
        capacity = scalars['CAPACITY_N']
        
    fm2.addConstrs((gb.quicksum([refine[o,m] for o in t]) <= capacity for m in months))

# 3. The maximum storage capacity cannot be surpassed

# This constraint is included as an upper bound of the <inv> variables
#fm2.addConstrs((inv[i] <= scalars['STORAGE'] for i in oils_months))

# 4. There are hardness bounds for the final product linearly dependent of the individual hardness of each raw oil used

fm2.addConstrs((gb.quicksum((hardness[o] - scalars['HL'])*refine[o,m] for o in oils) >= 0 for m in months))
fm2.addConstrs((gb.quicksum((hardness[o] - scalars['HU'])*refine[o,m] for o in oils) <= 0 for m in months))

# 5. Associating binary variables to refine variables

fm2.addConstrs((refine[j] <= delta[j]*max(scalars['CAPACITY_N'],scalars['CAPACITY_V']) for j in oils_months))

# 6. The food may be never made up of more than three oils in any month

fm2.addConstrs((delta.sum('*',m) <= 3 for m in months))

# 7. If an oil is used in any month, at least 20 tons must be used

fm2.addConstrs((refine[j] >= delta[j]*20 for j in oils_months))

# 8. If either VEG 1 or VEG 2 are used in a month then OIL 3 must also be used

# fm2.addConstrs((delta['VEG 1',m] + delta['VEG 2',m] <= 2*delta['OIL 3',m] for m in months))
fm2.addConstrs((delta['VEG 1',m] <= delta['OIL 3',m] for m in months))
fm2.addConstrs((delta['VEG 2',m] <= delta['OIL 3',m] for m in months))

#-------------- Model Execution

fm2.setParam('OutputFlag',0)    # Turns off the Optimization Details sheet print after the model.optimize() call
# Setting timer
begin = datetime.now()

fm2.optimize()

# Stopping timer
elapsed = datetime.now() - begin

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
print(f'Objective function value: ${round(fm2.objval,1) - 12500}')
print(f'Time elapsed: {round(elapsed.microseconds/1000000,2)} seconds')
print('***************************************')
print('$12.500 were deducted of \nstore costs of the last month')


#%% End of file