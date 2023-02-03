# -*- coding: utf-8 -*-
"""
Created on Mon Aug  3 18:01:47 2020

@author: johan
"""
#*******************************
# REFINERY OPTIMISATION
#*******************************

#%% Importing Gurobi Shell and other libraries

import gurobipy as gb
import pandas as pd
from datetime import datetime


#%% Model Data

# Importing data from excel file
data = pd.read_excel('Parameters.xlsx', index_col = None, header = 0, sheet_name = None)

# Creating dictionaries with the data

dic_octane = {}
for index, row in data['octane'].iterrows(): 
    dic_octane[row[0]] = row[1]
    
dic_fractions = {}
for index, row in data['fractions'].iterrows(): 
    dic_fractions[row[0],row[1]] = row[2]
    
dic_yield_reform = {}
for index, row in data['yield_reform'].iterrows(): 
    dic_yield_reform[row[0]] = row[1]
    
dic_yield_crack_oil = {}
for index, row in data['yield_crack_oil'].iterrows(): 
    dic_yield_crack_oil[row[0]] = row[1]
    
dic_yield_crack_gas = {}
for index, row in data['yield_crack_gas'].iterrows(): 
    dic_yield_crack_gas[row[0]] = row[1]

dic_octane_petrols = {}
for index, row in data['octane_petrols'].iterrows(): 
    dic_octane_petrols[row[0]] = row[1]
    
dic_vapor_fuel = {}
for index, row in data['vapor_fuel'].iterrows(): 
    dic_vapor_fuel[row[0]] = (row[1],row[2])
    
dic_availability = {}
for index, row in data['availability'].iterrows(): 
    dic_availability[row[0]] = row[1]
    
dic_profit = {}
for index, row in data['profit'].iterrows(): 
    dic_profit[row[0]] = row[1]
    
scalars = data['scalars'].set_index('NAME')

# Creating model parameters
naphtha_gas, o = gb.multidict(dic_octane) 
_, f = gb.multidict(dic_fractions)  
naphthas, r = gb.multidict(dic_yield_reform)  
standard, co = gb.multidict(dic_yield_crack_oil) 
standard, cg = gb.multidict(dic_yield_crack_gas) 
petrols, m = gb.multidict(dic_octane_petrols) 
oils, s, q = gb.multidict(dic_vapor_fuel) 
crudes, a = gb.multidict(dic_availability)
products, profits = gb.multidict(dic_profit)

naphtha_gas_petrol = []
for j in naphtha_gas:
    for p in petrols:
        naphtha_gas_petrol.append((j,p))
        
naphtha_gas_petrol = gb.tuplelist(naphtha_gas_petrol)

del data, dic_octane, dic_fractions, dic_yield_reform, dic_yield_crack_oil, dic_yield_crack_gas, dic_octane_petrols
del dic_vapor_fuel, dic_availability, dic_profit, index, row

#%% Model Formulation

#-------------- Model Creation

model = gb.Model('Refinery Optimisation')

#-------------- Variables Creation
distille = model.addVars(crudes, name='distille', ub = a)
reform = model.addVars(naphthas, name='reform')
crack = model.addVars(standard, name='crack')
blendp = model.addVars(naphtha_gas_petrol, name='blendp')
blendj = model.addVars(oils, name='blendj')
sell = model.addVars(products, obj = profits,  name='sell')
sell['Lube oil'].lb = int(scalars.loc['ll'])
sell['Lube oil'].ub = int(scalars.loc['lu'])
        
model.ModelSense = gb.GRB.MAXIMIZE

#------------- Constraints Creation

# 1.	The barrels of naphtha available for reforming and blending petrol depend on the fraction of distilled 
#       crude barrels that produce that naphtha

model.addConstrs((reform[n] + blendp.sum(n,'*') == gb.quicksum(f[c,n]*distille[c] for c in crudes) for n in naphthas))

# 2.	The barrels of oils available for cracking and blending jet fuel and fuel oil depend on the distilled 
#       crude barrels

model.addConstrs((crack[o] + blendj[o] + sell['Fuel oil']*q[o]/sum(q.values()) == gb.quicksum(f[c,o] * distille[c] for c in crudes)
                  for o in standard))

# 3.	The barrels of reformed gasoline available for blending petrols depend on the naphtha reformed

model.addConstr(gb.quicksum(blendp['Reformed Gasoline',p] for p in petrols) == gb.quicksum(r[n]*reform[n] for n in naphthas))

# 4.	The barrels of cracked oil available for blending jet fuel and fuel oil depend on the barrels cracked

model.addConstr(blendj['Cracked Oil'] + sell['Fuel oil']*q['Cracked Oil']/sum(q.values()) == crack.prod(co))

# 5.	The barrels of cracked gasoline available for blending petrols depend on the barrels cracked

model.addConstr(blendp.sum('Cracked Gasoline','*') == crack.prod(cg))

# 6.	The barrels of lube oil available for selling depend on the residuum barrels cracked

l = float(scalars.loc['l'])
model.addConstr(sell['Lube oil'] == l*crack['Residuum'])

# 7.	The barrels of each petrol available for sale depend on the blended naphtha and gasolines

model.addConstrs((sell[p] == blendp.sum('*',p) for p in petrols))

# 8.	Also, the barrels of petrol have each a minimum octane number according to the blended materials

model.addConstrs((sell[p]*m[p] <= gb.quicksum(o[j]*blendp[j,p] for j in naphtha_gas) for p in petrols))

# 9.	The barrels of jet fuel available for sale depends on the blended oils

model.addConstr(sell['Jet fuel'] == blendj.sum('*'))

# 10.	The barrels of jet fuel have a maximum vapor pressure, which depends on blended materials’ pressures

S = int(scalars.loc['S'])
model.addConstr(blendj.prod(s) <= S*sell['Jet fuel'])

# 12.	Machines capacities

model.addConstr(distille.sum('*') <= scalars.loc['D'])
model.addConstr(reform.sum('*') <= scalars.loc['R'])
model.addConstr(crack.sum('*') <= scalars.loc['C'])

# 14.	Premium and regular petrol relationship

model.addConstr(0.4*sell['Regular petrol'] <= sell['Premium petrol'])

#-------------- Model Execution

model.setParam('OutputFlag',1)    # Turns off the Optimization Details sheet print after the model.optimize() call
# Setting timer
begin = datetime.now()

model.optimize()

# Stopping timer
elapsed = datetime.now() - begin

#%% Results Report

# Reporting variables values

print('---------------------------\nDistille\n---------------------------')
for c in crudes:
    val = distille[c].x
    if val > 0:
        print(f'{c} -> {val}')
     
print('---------------------------\nReform\n---------------------------')
for n in naphthas:
    val = reform[n].x
    if val > 0:
        print(f'{n} -> {val}')

print('---------------------------\nCrack\n---------------------------')
for j in standard:
    val = crack[j].x
    if val > 0:
        print(f'{j} -> {val}')

print('---------------------------\nBLend\n---------------------------')
for p in petrols:
    print(f'\n{p}:')
    for j in naphtha_gas:
        val = blendp[j,p].x
        if val > 0:
            print(f'\t{j} -> {val}')
            
print('\nJet fuel:')
for j in oils:
    val = blendj[j].x
    if val > 0:
        print(f'\t{j} -> {val}')
        
print('\nFuel oil:')
for j in oils:
    val = sell['Fuel oil'].x
    if val > 0:
        print(f'\t{j} -> {val*q[j]}')            

print('---------------------------\nSales\n---------------------------')
for j in products:
    val = sell[j].x
    if val > 0:
        print(f'{j} -> {val}')
    

# Reporting objective function value
print('\n***************************************')
print(f'Objective function value: £{round(model.objval)}')
print(f'Time elapsed: {round(elapsed.microseconds/1000000,2)} seconds')
print('***************************************')


print('\nDifference with respect to answer in the \nbook (£211.365) is due to the approximations \nused in the book')
#%% End of file
PMF = 6818
RMF = 17044
JF = 15156
FO = 0
LBO = 500


