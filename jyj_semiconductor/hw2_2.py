import time
start = time.time()

import pyodbc
from docplex.mp.model import Model
conn = pyodbc.connect((r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=C:\Users\User\Desktop\Micron.accdb;'))
cursor = conn.cursor()

lot_size = 25

#importing WS names and costs
cursor.execute("select * from table1")
rows = cursor.fetchall()
WS=[]
cost = {}
for row in rows:
    WS.append(row[1])
    cost.update({row[1]: row[2]})

#importing table 5 data
cursor.execute("select * from table5")
rows = cursor.fetchall()
E = {} #creating dictionary
c=[] #creating list of column names so that dictionary will find keys
for row in cursor.columns(table='table5'):
    c.append(row.column_name) #adding column names
for row in rows:
    k = row[0] #dictionary key
    tempdict = {} #creating an inner dictionary
    for i in range(1, len(row)): 
        cdict = c[i] #key for the inner dictionary
        v = row[i] #value for the inner dictionary
        tempdict.update({cdict: v}) #adding key-value to inner dictionary
    E.update({k: tempdict}) #adding inner dictionary to dictionary
    
#importing demand data
cursor.execute("select * from table3")
rows = cursor.fetchall()
c=[]
for row in cursor.columns(table='table3'):
    c.append(row.column_name)
Product = c[2:]
Demand = {}
for row in rows:
    for i in range(2,5):
        Demand.update({c[i] : (row[i] / lot_size)})

#importing data 2 table
cursor.execute("select * from table2")
rows = cursor.fetchall()
A= {}
f=[]
for row in cursor.columns(table='table2'):
    f.append(row.column_name)
for row in rows:
    k = row[0]
    tempdict = {}
    for i in range(1,4):
        fdict = f[i]
        v = row[i]
        tempdict.update({fdict: v})
    A.update({k : tempdict})

step = list(A.keys())
#print(cost[WS[0]]-1)
with Model() as md1:
    x = md1.integer_var_list(len(WS) ,name='WS')
    s = md1.continuous_var_list(step,name='steps')
    t = md1.continuous_var_matrix(step, WS, name='t')
    q=0
    sol = md1.sum(cost[WS[i]]*x[i] for i in range(len(WS)))

    md1.minimize(sol)
    q=0
    for l in step:
        md1.add_constraint((md1.sum(Demand[m]*A[l][m] for m in Product) <= s[l-1]))
    for l in step:
         md1.add_constraint(md1.sum(t[l,i] for i in WS if not E[l][i] == 0) >= s[l-1])
    q=0
    for i in WS:
        for l in step:
            md1.add_constraint(E[l][i]*t[l,i] <= 7*24*60*x[q])
        q=q+1
    md1.solve_details
    solution = md1.solve()
    print(solution)

end = time.time()
print(end - start)