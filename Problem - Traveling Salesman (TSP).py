"""
Created on Wed Jul 17 15:42:01 2019

@author: Johan

-----------------------------------------
Traveling Salesman Problem (TSP)
-----------------------------------------

Problem:
    Back in the days when salesmen traveled door-to-door hawking vacuums and encyclopedias, they had to plan 
    their routes, from house to house or city to city. The shorter the route, the better. Finding the shortest 
    route that visits a set of locations is an exponentially difficult problem: finding the shortest path for 
    20 locations is much more than twice as hard as 10 locations.
    
    An exhaustive search of all possible paths would be guaranteed to find the shortest, but is computationally 
    intractable for all but small sets of locations. For larger problems, optimization techniques are needed to 
    intelligently search the solution space and find near-optimal solutions.

    Mathematically, traveling salesman problems can be represented as a graph, where the locations are the nodes 
    and the edges (or arcs) represent direct travel between the locations. The weight of each edge is the distance 
    between the nodes. The goal is to find the path with the shortest sum of weights.
    
    Here are some variants of the TSP:
        - Assymetric TSP
        - Price Collecting TSP
        - TSP with Time Windows
        
    In general, the traveling salesman problem is hard to solve. If there is a way to break this problem into 
    smaller component problems, the components will be at least as complex as the original one. This is what 
    computer scientists call NP-hard problems.

Example: 
    In this example, we consider a salesman traveling in the US. The salesman starts in New York and has 
    to visit a set of cities on a business trip before returning home. The problem then consists of  
    finding the shortest tour which visits every city on the itinerary.

Formulation:

minimize    ∑(i,j)∈E c_ij*x_ij

subject to

∑j∈V x_ij = 2,  ∀i∈V

∑i,j∈S,i≠j x_ij ≤ |S| − 1,  ∀ S⊂V, S ≠ {}    # These constraints are called Capacity Cut Constraints, they 
                                                # remove the subtours formed in a solution. Since there are a lot
                                                # of constraints, they are introduced as needed. (lazy constraints)
x_ij∈{0,1}

where

V = set of vertices (cities)
E = set of edges (pairs of vertices)
S = set of subtours

d_ij = distance between the pair (i,j) of cities

x_ij = 1: The edge (i,j) is in the tour
       0: Otherwise
"""
#%% Importing Gurobi Shell and other libraries

from gurobipy import *
from pandas import *
from matplotlib.pyplot import *
        
#%% Model Data

# Importing data from excel file

df_distance_matrix = read_excel('Parameters TSP.xlsx', index_col = None, header = 0, sheet_name = 'distance')
df_coord = read_excel('Parameters TSP.xlsx', index_col = 'CITY', header = 0, sheet_name = 'coordinates')

# Transforming the distance matrix df_distance into a row-wise distance DataFrame

df_distance = melt(df_distance_matrix, id_vars = 'NODE I', var_name = 'NODE J', value_name = 'DISTANCE')

# Eliminating the (i,i) edges and the duplicate (i,j) = (j,i) edges

for index, row in df_distance.iterrows():
    if int(row[0][5:]) >= int(row[1][5:]):    # Filters both, (i,i) and duplicate edges
        df_distance = df_distance.drop(index)
        
# Creating dictionaries with the data

dic_distance = {}

for index, row in df_distance.iterrows():
    dic_distance[(row[0],row[1])] = row[2]

# Creating model parameters

vertices = tuplelist(df_distance_matrix.iloc[:,0])
edges, distance = multidict(dic_distance)

#%% Results Report

def report_results():
    # Reporting variables values    
    print('----------------------------------------\nSelected edges:\n----------------------------------------')
    
    for i,j in edges:
        if x[i,j].x > 0:
            print(i,'->',j,':',x[i,j].obj)
            
    # Reporting objective function value    
    print('****************************************\nThe Total Distance Traveled is: ', round(tsp.objVal),'\n****************************************')


#%% Create plot
'''
Using <%matplotlib inline> or <%matplotlib qt> is possible 
to define wheter a plot is shown in the console or a window
'''
def create_plot():
    close()  # Erase previous plot

    xlabel("X Coordinate")
    ylabel("Y Coordinate")
    title("Travelling Salesman Problem")
    scatter(df_coord['X'],df_coord['Y'],50,'Blue','o')
    for index, row in df_coord.iterrows():
        text(row[0]-5,row[1]+4,index,fontsize = 10)
    for i,j in edges:
        if x[i,j].x > 0:
            plot([df_coord.loc[i,'X'],df_coord.loc[j,'X']], [df_coord.loc[i,'Y'],df_coord.loc[j,'Y']], c = 'red', alpha = 0.8, linestyle = '--',linewidth = 1)
    show()


#%% Model Formulation

#-------------- Model Creation

tsp = Model('Traveling Salesman Problem')

#-------------- Variables Creation

x = tsp.addVars(edges, name = 'include', obj = distance, vtype = GRB.BINARY)
tsp.ModelSense = GRB.MINIMIZE

#------------- Constraints Creation

# 1. Each vertex must be visited twice, once entering and one leaving

tsp.addConstrs((x.sum(i,'*') + x.sum('*',i) == 2 for i in vertices), 'visited')

#-------------- Model Execution

tsp.setParam('OutputFlag',0)    # Turns off the Optimization Details sheet print after the tsp.optimize() call
tsp.optimize()

#%% Adding lazy constraints as needed

# This function creates and populates the subtour and checked lists making use of a recursive function call

def search_next(i):
    
    for j in vertices:
        
        # The if and elif statementes make sure that the current j vertice exists, isn't checked and takes value
        
        if (i,j) in edges and j not in checked:
            if x[i,j].x > 0:
                subtour.append(j)
                checked.append(j)
                search_next(j)
                
        elif (j,i) in edges and j not in checked:
            if x[j,i].x > 0:
                subtour.append(j)
                checked.append(j)
                search_next(j)

# The while reoptimizes the model such that in each iteration the subtours are identified and the lazy constraints are added

iterate = True

while iterate:
    create_plot()
    
    subtours = []
    checked = []
    
    for i in vertices:
        subtour = []
        
        if i not in checked:
            subtour.append(i)
            checked.append(i)
            search_next(i)
            subtours.append(subtour)
    
    # This nested fors create the linear expression object that is the sum of the variable in each identified subtour
    
    if len(subtours) > 1:
        for s in subtours:
            expr = LinExpr()
            for i in s:
                for j in s:
                    if (i,j) in edges:
                        expr += x[i,j] 
                        
            # 2. Lazy Constraints Addition
            tsp.addConstr(expr <= len(s) - 1,'subtour')
            
        # Run the model again
        tsp.optimize()  
        
    # This else end the loop when there is only one subtour in the subtour list, which means, the solution is feasible
    else:
        report_results()
        iterate = False
