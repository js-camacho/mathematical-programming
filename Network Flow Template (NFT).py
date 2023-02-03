# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 16:53:54 2019
Last Edited on Thu Feb 25 11:00:30 2021

@author: Johan

*************************************
 Network Flow Optimization Template
*************************************

Problem: 
    The Network Flow Problems can address a variety of problems applied
    to different areas. Although, these problems share a common structure that
    will be developed in this model. The parameters of the problem and some details
    regarding the type of problem that can be represented with different groups
    of parameters will be displayed in the auxiliar parameters file: 
    'Parameters NFT.xslx'
    The great advantage of the Network Structure consist in the fact that the solution
    for a relaxed problem (continous decision variables) will have an integral solution
    if the parameters are integers, without having to appeal to Branch & Bound. Also,
    this type of problems escalate well for huge instances.

Formulation:
    
    min  ∑(i,j)∈A c_ij*x_ij
    
    subject to
    
     ∑{i|(i,j)∈A} x_ij - ∑{i|(j,i)∈A} x_ji = b_i ∀i∈N
    
    l_ij <= x_ij <= u_ij
    
    where
    
    N: set of nodes
    A: set of arcs
    
    c_ij: cost of the arc (i,j)
    l_ij: lower bound of the arc (i,j)
    u_ij: upper bound of the arc (i,j)
    b_i: demand/supply/transfer parameter of node i
    
    x_ij: flow through the arc (i,j)

"""

#%% Importing Gurobi Shell and other libraries

import gurobipy as gb
import pandas as pd
import openpyxl  # This module is used by pandas.read_excel() for newer versions of Excel
import os

#%% Using directory where the file is located
abspath = os.path.abspath(__file__)
dname = os.path.dirname(abspath)
os.chdir(dname)

#%% Model Data

# Importing data from excel file
df_nodes = pd.read_excel('Parameters NFT.xlsx', index_col = None, header = 0, sheet_name = 'nodes')
df_edges = pd.read_excel('Parameters NFT.xlsx', index_col = None, header = 0, sheet_name = 'edges')

# Creating dictionaries
dic_nodes = {row[0]:row[1] for index,row in df_nodes.iterrows()}    
dic_edges = {(row[0],row[1]):(row[2],row[3],row[4]) for index, row in df_edges.iterrows()}
    
    
# Creating model parameters
nodes, b = gb.multidict(dic_nodes)
edges, c, l, u = gb.multidict(dic_edges)


#%% Model Formulation

#-------------- Model Creation

nf = gb.Model('Network Flow')

#-------------- Variables Creation

x = nf.addVars(edges, name = 'flow', obj = c, vtype = gb.GRB.CONTINUOUS, lb = l, ub = u)
nf.ModelSense = gb.GRB.MINIMIZE

#------------- Constraints Creation

# Balance Constraints

nf.addConstrs((x.sum(i,'*') - x.sum('*',i) == b[i] for i in nodes),'balance')

#-------------- Model Execution

nf.setParam('OutputFlag',0)    # Turns off the Optimization Details sheet print after the tsp.optimize() call
nf.optimize()

#%% Results Report

print('---------------------------------\nFlow Variables:\n---------------------------------')

for i,j in edges:
    if x[i,j].x > 0:
        print(i,'->',j,':',x[i,j].x,'($'+str(c[i,j]*x[i,j].x)+')')
        
print('---------------------------------\nObjective Funtion Value: $',nf.objval,'\n---------------------------------')    