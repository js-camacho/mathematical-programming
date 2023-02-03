# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 14:26:29 2020

@author: johan
"""
#*******************************
# FACTORY PLANNING II
#*******************************

#%% Importing Gurobi Shell and other libraries

import gurobipy as gb
import pandas as pd
from datetime import datetime
        
#%% Model Data

# Importing data from excel file
data = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = None)

# Creating dictionaries with the data

dic_profit = {}
for index, row in data['profit'].iterrows(): dic_profit[row[0]] = row[1]
    
dic_prod = {}
for index, row in data['production'].iterrows(): dic_prod[row[0],row[1]] = row[2]

dic_mach = {}
for index, row in data['machinery'].iterrows(): dic_mach[row[0]] = row[1]

dic_demand = {}
for index, row in data['demand'].iterrows(): dic_demand[row[0],row[1]] = row[2]

scalars = {}
for index, row in data['scalars'].iterrows(): scalars[row[0]] = row[1]

# Creating model parameters
products, c = gb.multidict(dic_profit) 
products_machines, a = gb.multidict(dic_prod)
machines, n = gb.multidict(dic_mach)
products_months, b = gb.multidict(dic_demand)

machines = pd.Series(machines)
months = data['demand']['MONTH'].unique()

prices = {}
for k, v in c.items():
    for t in months:
        prices[(k,t)] = v

del data, dic_demand, dic_mach, dic_prod, dic_profit, index, row

#%% Model Formulation

#-------------- Model Creation

model = gb.Model('Factory Planning I')

#-------------- Variables Creation
x = model.addVars(products_months, name = 'produce', obj = 0)
y = model.addVars(products_months, name = 'sell', obj = prices, ub = b)
q = model.addVars(products_months, name = 'store', obj = -scalars['StorageCost'], ub = scalars['StorageCapacity'])
z = model.addVars([(m,t) for m in machines for t in months], name = 'maintenance', obj = 0, vtype = gb.GRB.INTEGER, ub = 3)

model.ModelSense = gb.GRB.MAXIMIZE

#------------- Constraints Creation

# 1.	The production policy of a month cannot surpass the production hours availability of any machine

c1 = model.addConstrs((gb.quicksum(a[p,m]*x[p,t] for p in products) 
                  <= scalars['ProductiveHours']*(n[m] - z[m,t]) for m in machines for t in months), 'prod_capacity')

# 2.	The units sold in any month must be less or equal to the units demanded (upper bound constraint)

# 3.	Relationship between units produced, sold and stored:

model.addConstrs((q[p,'Jan'] == x[p,'Jan'] - y[p,'Jan'] for p in products), 'initial inventory')    

for i in range(1,len(months)):
    t = months[i]
    t1 = months[i-1]
    model.addConstrs((q[p,t] == q[p,t1] + x[p,t] - y[p,t] for p in products),'inventory')
    
model.addConstrs((q[p,'Jun'] == scalars['FinalStorage'] for p in products), 'final inventory')

# 4.	Storage capacity (upper bound constraint)

# 5.	Each machine should enter maintenance once (grinder enter 2 at once) in the six months

model.addConstrs((z.sum(m,'*') == n[m] for m in machines[machines != 'Grinding']))
model.addConstr(z.sum('Grinding','*') == 2)

#-------------- Model Execution

model.setParam('OutputFlag',0)    # Turns off the Optimization Details sheet print after the model.optimize() call
# Setting timer
begin = datetime.now()

model.optimize()

# Stopping timer
elapsed = datetime.now() - begin

#%% Results Report

# Reporting variables values
for t in months:
    print(f'---------------------------\n{t}\n---------------------------')
    print('Production:--------')
    for p in products:
        val = x[p,t].x
        if val > 0:
            print(f'\t{p} -> {val}')
    print('Sales:-------------')
    for p in products:
        val = y[p,t].x
        if val > 0:
            print(f'\t{p} -> {val} (£{c[p]} each)')
    print('Inventory:---------')
    for p in products:
        val = q[p,t].x
        if val > 0:
            print(f'\t{p} -> {val}')       
    print('Maintenance:-------')      
    for m in machines:
        val = z[m,t].x
        if val > 0:
            print(f'\t{m}')  
        

# Reporting objective function value
print('***************************************')
print(f'Objective function value: £{round(model.objval)}')
print(f'Time elapsed: {round(elapsed.microseconds/1000000,2)} seconds')
print('***************************************')


#%% Reporting constraints slacks
'''
for t in months:
    print(f'---------------------------\n{t}\n---------------------------')
    for m in machines:
        print(f'{m} - Available = {scalars["ProductiveHours"]*n[m,t]}')
        total = 0
        for p in products:
            use = a[p,m]*x[p,t].x 
            total += use
            print(f'\t{p} -> {use}')
        print(f'\t~~~~Total: {total}')
'''

#%% End of file