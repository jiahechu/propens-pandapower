# -*- coding: utf-8 -*-
"""
Created on Wed Dec 28 16:45:28 2022

@author: marti
"""
# import pandas as pd
#from openpyxl import Workbook
# from openpyxl import load_workbook
from tqdm import tqdm
from src.analysis.parameters import preallocate_table

#%%
def sort_results(net, number, time_steps, element_results, column, parameters, element, scenario_name): 

    print(' > of the '+ element+'s')

    # extract the values from the netowrk topology and power flow results
    table = preallocate_table(element, column, number)
    i = 0
    if net[element].shape[0] == 0:
        table['scenario'][0] = scenario_name
        return table
    
    for index in net[element].index:
        # from time step, here only the first iteration is done, so '0'
        table.loc[i,'step'] = 0
        # from network topology
        table.loc[i,'index'] = index 
        table.loc[i,parameters['net_' + element]] = net[element].loc[index, parameters['net_' + element]]                
        #from power flow results: without time series, directly from net.res_##; with TS, from results (dict/dataframe)
        if time_steps == 1:
            table.loc[i,parameters['res_' + element]] = net['res_'+element].loc[index,parameters['res_' + element]]  
        else:
            for parameter in parameters['res_' + element]:
                table.loc[0,parameter] = element_results[parameter].loc[0,index]
        i = i + 1    
    # results: if there is only '1' time step, directly from the pandapower network and the 'while' loop condiction won't start
    # otherwise the results were called from the temporary excels, and now they will be written from step 1 on
    # first step was 0 (zero)
    if time_steps > 1:
        for step in tqdm(range(1,time_steps,1)):
            offset = step*number['load'] # offset because all the values from all the loads and time steps are in the same table, 
            i = 0                                #so we write each time step results under the previous one
            for index in net[element].index:
                # time step
                table.loc[offset + i,'step'] = step
                # from network topology, is repeated for all the time steps
                table.loc[offset + i,'index'] = index 
                table.loc[offset + i,parameters['net_' + element]] = net[element].loc[index, parameters['net_' + element]] 
                # from results, each value is taken from a different dataframe within results dict
                for parameter in parameters['res_' + element]:
                    table.loc[offset + i,parameter] = element_results[parameter].loc[step,index]
                i = i + 1
            #we go to the next time step  
            step = step + 1
            
    return table
















