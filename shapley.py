import xlwings as xw
import more_itertools
import numpy as np
import pandas as pd
import math

def get_coalitions(divs):
    #Number of possible coalitions 
    n = 2**divs - 1
    #Create a vector of divisions from 0 to #divs - 1
    powerlist = [i for i in range(divs)]
    #Create all possible powersets
    coals = list(more_itertools.powerset(powerlist))
    #Put all possible powersets into list form
    for i in range(len(coals)):
        coals[i] = list(coals[i])
    #Delete empty set (???)
    del coals[0]
    #Generate the binary coalition matrix
    coal_final = np.zeros([n,divs])
    for i in range(n):
        coal_final[i, coals[i]] = 1
    return(coal_final, coals)  


def shapleycalc(characteristic_function):  
    #Change input list to array form
    cf = np.array(characteristic_function)
    
    #Total characteristic values in array
    n = len(cf)
    
    #Number of divisions
    divs = int(math.log2(n + 1))
    
    #Build binary coalition matrix and subset matrix 
    all_coalitions = get_coalitions(divs)
    coalitions = all_coalitions[0]
    subsets = all_coalitions[1]
    fac = math.factorial
    copy_coalitions = coalitions.copy()
    
    #Build ordering matrix
    for k in range(n):
        for i in range(divs):
            s = np.count_nonzero(copy_coalitions[k,])
            if coalitions[k, i] == 0:
                coalitions[k, i] = -1*(fac(s) * fac(divs - s - 1))/fac(divs) * cf[k]
            else:
                coalitions[k, i] = (fac(s - 1) * fac(divs - s))/fac(divs) * cf[k]
                    
    #Calculate shapley values 
    shapley = coalitions.sum(0)
    
    #What divisions should be inactive? Only used in first stage
    inactive = shapley <= 0.0001
    
    return(shapley, inactive, subsets)

#Function to find intersections of lists, to be used later
def intersection(a, b):
    c = [value for value in a if value in b]
    return c

# Build the row of divisions in excel sheet
@xw.func
def divbuild(divs):
    divs = int(divs)
    div = ["Division "+ str(i + 1) for i in range(divs)]
    
    return(div)

#Build the row of coalitions in excel sheet
@xw.func
def coals(divs):
    divs = int(divs)
    powerlist = [(i+1) for i in range(divs)]
    coals = list(more_itertools.powerset(powerlist))
    del coals[0]
    strcoals = []
    for i in range(len(coals)):
       strcoals.append(str(coals[i]))
    return(strcoals)


#Double shapley function
@xw.func(async_mode = 'threading')
def twostageshapley(revenue, cost, andcost):
    
    #Change inputs to arrays
    revenue = np.array(revenue)
    cost = np.array(cost)
    andcost = np.array(andcost)
    #number of divisions
    divs = len(revenue)
    #number of coalitions
    n = len(cost)
        
    #Create all possible powersets to be used as dictionary keys
    powerlist = [i for i in range(divs)]
    coals = list(more_itertools.powerset(powerlist))
    #removing empty set
    del coals[0]
    
    #Putting costs into dictionary for specific coalitions
    #One dictionary for and costs, one or or/direct costs
    costdict = {}
    andcostdict = {}
    for i in range(n):
        costdict[coals[i]] = cost[i]
        if i >= divs: 
            andcostdict[coals[i]] = andcost[i - divs]
    
    #Create a vector of all included or/direct costs for a given coalition
    coalitioncost = np.zeros([n])
    for j in range(n):
         for i in costdict.keys():
             jkeys=list(costdict.keys())[j]
             ikeys=list(i)
             if len(intersection(jkeys, ikeys)) > 0:
                 coalitioncost[j] += costdict[i]
    
    #Create a vector of all and costs attributed to a certain coalition
    andcostcoals = np.zeros([n - divs]) 
    for j in range(n - divs):
        for i in andcostdict.keys():
            jkeys = list(andcostdict.keys())[j]
            ikeys = list(i)
            if set(ikeys).issubset(set(jkeys)):
                andcostcoals[j] += andcostdict[i]
    
    #Total cost for a coalition
    totcost = np.zeros([n])
    totcost[:divs] = coalitioncost[:divs]
    totcost[divs:] = coalitioncost[divs:] + andcostcoals
       
    #Find coalition revenue and profits
    coalitionrev = np.zeros([n])
    for i in range(n):
        coalitionrev[i] = sum(revenue[[coals[i]]])
    coalitionprofits = coalitionrev - totcost
    profdict = {}
    for i in range(n):
        profdict[coals[i]] = coalitionprofits[i]
       
    #First stage characteristic function
    V = np.zeros([n])    
    for j in range(n):
        maxval = 0
        for i in profdict.keys():
            jkeys = list(profdict.keys())[j]
            ikeys = list(i)
            if set(ikeys).issubset(set(jkeys)):
                temp = profdict[i]
                if temp > maxval:
                    maxval = temp
                V[j] = maxval
    
    #First stage shapley values and inactive divisions
    firststage = shapleycalc(V)
    fsvals = firststage[0]
    inactive = firststage[1]
    
    ssrevenue = revenue * inactive
    
    #Get list of lists to index in stage 2
    coalist = firststage[2]
    
    #Binary dictionary of inactive coalitions
    ssdict = {}
    for i in range(n):
        ssdict[coals[i]] = min(inactive[coalist[i]])
        
    andcoalist = coalist[divs:]
    
    #Finding coalition costs conditional on active/inactive or/direct costs
    sscost = [ssdict[key] * costdict[key] for key in ssdict]
    
    
    #Add "and" costs to coalition costs
    for i in range(n - divs):
        if all(inactive[andcoalist[i]] == 1):
            sscost[i + divs] += andcostcoals[i]
    
    #Put coalition costs and total coalition revenue into a dictionary
    ssrevdict = {}
    sscostdict = {}
    for i in range(n):
       ssrevdict[coals[i]] = sum(ssrevenue[[coals[i]]])
       sscostdict[coals[i]] = sscost[i]
    
    #Total coalition cost
    sstot = np.zeros([n])
    for j in range(n):
         for i in sscostdict.keys():
             jkeys=list(sscostdict.keys())[j]
             ikeys=list(i)
             if len(intersection(jkeys, ikeys)) > 0:
                 sstot[j] += sscostdict[i]
    
    #Second stage characteristic function
    ssV = np.array(list(ssrevdict.values())) - sstot
    
    #Second stage shapley values and other elements of solution table
    ssvals = shapleycalc(ssV)[0]
    profit = revenue - cost[:divs]
    profitshares = fsvals + ssvals
    overheadshares = profit - profitshares
    
    #Construct and return solution table
    tabledict = {"1st Stage Vals": np.round(fsvals,3),
                 "2nd Stage Vals": np.round(ssvals,3),
                 "Overhead Shares": overheadshares,
                 "Profit Shares": profitshares}
    
    divlist = ["Division " + str(i+1) for i in range(divs)]
    table = pd.DataFrame(data = tabledict, index = divlist)
    total = table.sum()
    total.name = 'Total'
    table = table.append(total.transpose()).transpose()
    return(table)

