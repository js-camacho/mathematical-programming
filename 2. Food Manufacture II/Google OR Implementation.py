# -*- coding: utf-8 -*-
"""
Created on Tue Jul  7 16:50:09 2020

@author: johan
"""
#*******************************
# FOOD MANUFACTURE II
#*******************************

#%% Importing libraries

from ortools.linear_solver import pywraplp
import pandas as pd
from datetime import datetime

#%% Importing parameters

df_cost = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = 'Cost')
df_hardness = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = 'Hardness')
df_scalars = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = 'Scalars')

scalars = {}
for index, row in df_scalars.iterrows():
    scalars[row[0]] = row[1]

months = df_cost.MONTH.unique()
oils = df_hardness.OIL

#%% Model Formulation

# Instantiate a Glop solver
solver = pywraplp.Solver('Food Manufacture II', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)

# ------------------------ Declaring objective
objective = solver.Objective()
objective.SetMaximization()

# ------------------------ Creating variables
refine = {}
buy = {}
inv = {}
delta = {}
for index, row in df_cost.iterrows():
    # Creating <refine> variables
    var = solver.NumVar(0, solver.infinity(), f'Refine[{row[0]},{row[1]}]')
    refine[(row[0],row[1])] = var
    objective.SetCoefficient(var, scalars['PRICE'])
    # Creating <buy> variables
    var = solver.NumVar(0, solver.infinity(), f'Buy[{row[0]},{row[1]}]')
    buy[(row[0],row[1])] = var
    objective.SetCoefficient(var, row[2])
    # Creating <inv> variables
    var = solver.NumVar(0, scalars['STORAGE'], f'Inv[{row[0]},{row[1]}]')
    inv[(row[0],row[1])] = var
    objective.SetCoefficient(var, scalars['STORECOST'])
    # Creating <delta> variables
    var = solver.IntVar(0, 1, f'Delta[{row[0]},{row[1]}]')
    delta[(row[0],row[1])] = var
    objective.SetCoefficient(var, 0)

# ------------------------ Creating Constraints
# 1. The inventory at the end of a month depends of the tons of raw oils refined and bought in that month
for o in oils:
    # 1.1 Initial inventory: inv[o,'Jan'] + refine[o,'Jan'] - buy[o, 'Jan'] = 500
    const = solver.Constraint(scalars['INITIAL'], scalars['INITIAL'])
    const.SetCoefficient(inv[o,months[0]], 1)
    const.SetCoefficient(refine[o,months[0]], 1)
    const.SetCoefficient(buy[o,months[0]], -1)
    # 1.1 Final inventory: inv[o,'May'] - refine[o,'Jun'] + buy[o,'Jun'] = 500
    const = solver.Constraint(scalars['FINAL'], scalars['FINAL'])
    const.SetCoefficient(inv[o,months[-2]], 1)
    const.SetCoefficient(refine[o,months[-1]], -1)
    const.SetCoefficient(buy[o,months[-1]], 1)
    
    for j in range(1, len(months)-1):
        m = months[j]
        m1 = months[j-1]
        # inv[o,m] - inv[o,m-1] + refine[o,m] - buy[o,m] = 0
        const = solver.Constraint(0,0)
        const.SetCoefficient(inv[o,m], 1)
        const.SetCoefficient(inv[o,m1], -1)
        const.SetCoefficient(refine[o,m], 1)
        const.SetCoefficient(buy[o,m], -1)
        
# 2. There are maximum refining capacities for each type of raw oil each month

for m in months:
    # sum(o in T) refine[o,m] <= capacity[T]
    const_v = solver.Constraint(-solver.Infinity(),scalars['CAPACITY_V'])
    const_n = solver.Constraint(-solver.Infinity(),scalars['CAPACITY_N'])   
    for index, row in df_hardness.iterrows():
        o = row[0]
        typ = row[2]
        if typ == 'V':
            const_v.SetCoefficient(refine[o,m], 1)
        else:
            const_n.SetCoefficient(refine[o,m], 1)
     
# 3. The maximum storage capacity cannot be surpassed

# This constraint is included as an upper bound of the <inv> variables

# 4. There are hardness bounds for the final product linearly dependent of the individual hardness of each raw oil used

HL = scalars['HL']
HU = scalars['HU']
for m in months:
    # sum(o in oils) (h[o]-HL)*refine[o,m] >= 0, for m in months
    const_l = solver.Constraint(0,solver.Infinity())
    # sum(o in oils) (h[o]-HU)*refine[o,m] <= 0, for m in months
    const_u = solver.Constraint(-solver.Infinity(), 0)
    for index, row in df_hardness.iterrows():
        o = row[0]
        h = row[1]
        const_l.SetCoefficient(refine[o,m], h-HL)
        const_u.SetCoefficient(refine[o,m], h-HU)

# 5. Associating binary variables to refine variables

capacity = max(scalars['CAPACITY_N'],scalars['CAPACITY_V'])
for o in oils:
    for m in months:
        const = solver.Constraint(-solver.Infinity(), 0)
        const.SetCoefficient(refine[o,m], 1)
        const.SetCoefficient(delta[o,m], -capacity)

# 6. The food may be never made up of more than three oils in any month

for m in months:
    const = solver.Constraint(-solver.Infinity(), 3)
    for o in oils:
        const.SetCoefficient(delta[o,m], 1)

# 7. If an oil is used in any month, at least 20 tons must be used

for o in oils:
    for m in months:
        const = solver.Constraint(0,solver.Infinity())
        const.SetCoefficient(refine[o,m], 1)
        const.SetCoefficient(delta[o,m], -20)

# 8. If either VEG 1 or VEG 2 are used in a month then OIL 3 must also be used

for m in months:
    const1 = solver.Constraint(-solver.Infinity(), 0)
    const2 = solver.Constraint(-solver.Infinity(), 0)
    
    const1.SetCoefficient(delta['VEG 1',m], 1)
    const1.SetCoefficient(delta['OIL 3',m], -1)
    const2.SetCoefficient(delta['VEG 2',m], 1)
    const2.SetCoefficient(delta['OIL 3',m], -1)

# ------------------------ Model Execution
# Setting timer
begin = datetime.now()

status = solver.Solve()

# Stopping timer
elapsed = datetime.now() - begin

#%% Results report

# Reporting variables values
price = scalars['PRICE']
store = scalars['STORECOST']
for m in months:
    print(f'------------------\n{m}\n------------------')
    for o in oils:
        val = round(refine[o,m].solution_value(),2)
        if val > 0:
            print(f'{o}: Refine -> {val}')
            #print(f'{o}: Refine -> {val}, Profit: ${round(val*price,1)}')
        val = round(buy[o,m].solution_value(),2)
        if val > 0:
            #cost = costs[o,m]
            print(f'{o}: Buy -> {val}')
            #print(f'{o}: Buy -> {val}, Profit: ${round(val*cost,1)}')
        val = round(inv[o,m].solution_value(),2)
        if val > 0:
            print(f'{o}: Store -> {val}')
            #print(f'{o}: Store -> {val}, Profit: ${round(val*store,1)}')
        
# Reporting objective function value  

print('***************************************')
print(f'Objective function value: ${round(objective.Value(),1) - 12500}')
print(f'Time elapsed: {round(elapsed.microseconds/1000000,2)} seconds')
print('***************************************')
print('$12.500 were deducted of \nstore costs of the last month')

#%% End of file