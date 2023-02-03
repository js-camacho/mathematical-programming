# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 10:47:22 2020

@author: johan
"""
#*******************************
# FACTORY PLANNING I
#*******************************

#%% Importing libraries

from ortools.linear_solver import pywraplp
import pandas as pd
from datetime import datetime

#%% Importing parameters

# Importing data from excel file
data = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = None)

products = data['profit']['PRODUCT'].unique()
machines = data['machinery']['MACHINE'].unique()
months = data['machinery']['MONTH'].unique()

profits = data['profit']
profits = profits.set_index('PRODUCT')

production = data['production']
production = production.set_index(['PRODUCT','MACHINE'])

machinery = data['machinery']
machinery = machinery.set_index(['MACHINE','MONTH'])

demands = data['demand']
demands = demands.set_index(['PRODUCT', 'MONTH'])

scalars = data['scalars']
scalars = scalars.set_index('NAME')

#%% Model Formulation

# Instantiate a Glop solver
solver = pywraplp.Solver('Factory Planning I', pywraplp.Solver.GLOP_LINEAR_PROGRAMMING)

# ------------------------ Declaring objective
objective = solver.Objective()
objective.SetMaximization()

# ------------------------ Creating variables
use = {}
sell = {}
store = {}
for p in products:
    profit = int(profits.loc[p])
    for t in months:
        demand = int(demands.loc[p,t])
        # Creating <use> variables
        var = solver.NumVar(0, solver.infinity(), f'Use[{p},{t}]')
        use[(p,t)] = var
        objective.SetCoefficient(var, 0)
        # Creating <sell> variables with simple upper bound
        var = solver.NumVar(0, demand, f'Sell[{p},{t}]')
        sell[(p,t)] = var
        objective.SetCoefficient(var, profit)
        # Creating <store> variables with simple upper bound
        var = solver.NumVar(0, int(scalars.loc['StorageCapacity']), f'Store[{p},{t}]')
        store[(p,t)] = var
        objective.SetCoefficient(var, -float(scalars.loc['StorageCost']))

# ------------------------ Creating constrraints

# 1.	The production policy of a month cannot surpass the production hours availability of any machine
c1 = {}
for m in machines:
    for t in months:
        # sum(p in products) production[p,m]*use[p,m] <= ProductiveHours*machinery[m,t]
        upper_b = int(machinery.loc[m,t])*int(scalars.loc['ProductiveHours'])
        constr = solver.Constraint(0, upper_b)
        for p in products:
            constr.SetCoefficient(use[p,t], float(production.loc[p,m]))
        c1[m,t] = constr
            
# 2.	The units sold in any month must be less or equal to the units demanded (upper bound constraint)

# 3.	Relationship between units produced, sold and stored:
for p in products:
    # 3.1 Initial inventory: store[p,'Jan'] - use[p,'Jan'] + sell[o, 'Jan'] = 0
    constr = solver.Constraint(0, 0)
    constr.SetCoefficient(store[p,months[0]], 1)
    constr.SetCoefficient(use[p,months[0]], -1)
    constr.SetCoefficient(sell[p,months[0]], 1)
    # 3.2 Final inventory: inv[p,'Jun'] = 50
    final = int(scalars.loc['FinalStorage'])
    constr = solver.Constraint(final, final)
    constr.SetCoefficient(store[p,months[-1]], 1)
    
    for j in range(1, len(months)):
        t = months[j]
        t1 = months[j-1]
        # store[p,t] - store[p,t-1] - use[p,t] + sell[p,t] = 0
        constr = solver.Constraint(0,0)
        constr.SetCoefficient(store[p,t], 1)
        constr.SetCoefficient(store[p,t1], -1)
        constr.SetCoefficient(use[p,t], -1)
        constr.SetCoefficient(sell[p,t], 1)
    
# 4.	Storage capacity (upper bound constraint)

# ------------------------ Model Execution
# Setting timer
begin = datetime.now()

status = solver.Solve()

# Stopping timer
elapsed = datetime.now() - begin

#%% Results report

# Reporting variables values
for t in months:
    print(f'---------------------------\n{t}\n---------------------------')
    print('Production:')
    for p in products:
        val = round(use[p,t].solution_value(),2)
        if val > 0:
            print(f'\t{p} -> {val}')
    print('Sales:')
    for p in products:
        val = round(sell[p,t].solution_value(),2)
        if val > 0:
            print(f'\t{p} -> {val}')
    print('Inventory:')
    for p in products:
        val = round(store[p,t].solution_value(),2)
        if val > 0:
            print(f'\t{p} -> {val}')
    
# Reporting objective function value  

print('***************************************')
print(f'Objective function value: ${round(objective.Value(),1)}')
print(f'Time elapsed: {round(elapsed.microseconds/1000000,2)} seconds')
print('***************************************')


#%% Sensitivity Analysis

def SensitivityAnalysis():
    # Reduced Costs
    print('\n*******************************************')
    print('Reduced Costs: Recommended price increases')
    print('*******************************************')
    for t in months:
        for p in products:
            rc = round(sell[p,t].ReducedCost(),1)
            val = round(sell[p,t].solution_value(),1)
            ub = sell[p,t].ub()
            # Evade the cases where reduced cost is 0 or where the value take the upper bound value (bound shadow price)
            if rc != 0 and val != ub:
                print(f'{"-"*10}\n{t}\n{"-"*10}')
                print(f'{p}: {rc}')
            
    # Bound sensitivities of variables
    print('\n*******************************************')
    print('Shadow prices: Value of an extra machine hour')
    print('*******************************************')
    for t in months:
        for m in machines:
            shadow = round(c1[m,t].DualValue(),2)
            if shadow != 0:
                print(f'{"-"*10}\n{t}\n{"-"*10}')
                print(f'{m}: {shadow}')
            
SA = input('Print Sensitivity Analysis? [y/n]\n')
if SA == 'y':
    SensitivityAnalysis()
            
#%% End of file

