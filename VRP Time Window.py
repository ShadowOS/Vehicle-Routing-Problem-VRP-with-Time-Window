# -*- coding: utf-8 -*-
"""
Created on Mon May 13 23:40:01 2019

@author: Karnika
"""

from gurobipy import*
import os
import xlrd

book = xlrd.open_workbook(os.path.join("time window.xlsx"))



Node=[]
Demand={}           #Demand in Thousands
ServiceTime={}      #Service time in minutes
Distance={}         #Distance in kms
TravelTime={}       #Travel time in minutes
VehicleNum=[]       #Vehicle number
Cap = 20
ai={}
bi={}
Arc = {}

sh = book.sheet_by_name("demand")

i = 1
while True:
    try:
        sp = sh.cell_value(i,0)
        Node.append(sp)
        Demand[sp]=sh.cell_value(i,1)
        ServiceTime[sp]=sh.cell_value(i,2)
        ai[sp]=sh.cell_value(i,3)
        bi[sp]=sh.cell_value(i,4)
        i = i + 1   
    except IndexError:
        break
sh = book.sheet_by_name("VehicleNum")

i = 1
while True:
    try:
        sp = sh.cell_value(i,0)
        VehicleNum.append(sp)
        i = i + 1   
    except IndexError:
        break
cost={}
sh = book.sheet_by_name("Cost")
i = 1
for P in Node:
    j = 1
    for Q in Node:
        cost[P,Q] = sh.cell_value(i,j)
        j += 1
    i += 1
sh = book.sheet_by_name("Distance")
i = 1
for P in Node:
    j = 1
    for Q in Node:
        Distance[P,Q] = sh.cell_value(i,j)
        j += 1
    i += 1 
sh = book.sheet_by_name("TravelTime")
i = 1
for P in Node:
    j = 1
    for Q in Node:
        TravelTime[P,Q] = sh.cell_value(i,j)
        j += 1
    i += 1    
Aij = {}
sh = book.sheet_by_name("Aij")
i = 1
for P in Node:
    j = 1
    for Q in Node:
        Aij[P,Q] = sh.cell_value(i,j)
        j += 1
    i += 1
    
numberOfVechile=2

K=numberOfVechile

m=Model("Time window 1")


m.modelSense=GRB.MINIMIZE

xijk=m.addVars(Node,Node,VehicleNum,vtype=GRB.BINARY,name='X_ijk')

Tik = m.addVars(Node,VehicleNum,vtype=GRB.CONTINUOUS,name='T_ik')

m.setObjective(sum((cost[i,j]*xijk[i,j,k] for i in Node for j in Node for k in VehicleNum if  Aij[i,j] == 1)))

for i in Node :
    if i!='DepotStart' and i != 'DepotEnd' :
        m.addConstr(sum(xijk[i,j,k] for j in Node for k in VehicleNum if  Aij[i,j] == 1)==1)  
    
for k in VehicleNum :
    m.addConstr(sum(xijk['DepotStart',j,k] for j in Node if  Aij['DepotStart',j] == 1)==1)
    
for j in Node:
    for k in VehicleNum:
        if j!='DepotStart' and j != 'DepotEnd':
            m.addConstr(sum(xijk[i,j,k] for i in Node if  Aij[i,j] == 1)-sum(xijk[j,i,k] for i in Node if  Aij[j,i] == 1)==0)
        
for k in VehicleNum:
    m.addConstr(sum(xijk[i,'DepotEnd',k] for i in Node   if  Aij[i,'DepotEnd'] == 1)==1)

for i in Node:
    for j in Node:
        for k in VehicleNum:
            if Aij[i,j] == 1:
                m.addConstr(Tik[i,k]+ServiceTime[i]+TravelTime[i,j] -Tik[j,k]<= (1-xijk[i,j,k])*5000 ) #subtour elimination constraint

for i in Node:
    for k in VehicleNum:
        for j in Node:
            if Aij[i,j]==1:
                m.addConstr(Tik[i,k] >= ai[i]) and m.addConstr(Tik[i,k] <= bi[i])
            
                
for k in VehicleNum:
    m.addConstr((sum(Demand[i]*xijk[i,j,k]*Aij[i, j]  for i in Node for j in Node ))<=Cap)
    
m.optimize()
                
m.write('Timewindow.lp')
                            
for v in m.getVars():
    if v.x > 0.01:
        print(v.varName, v.x)
print('Objective:',m.objVal)