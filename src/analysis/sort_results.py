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
def sort_load_results(net, number, time_steps, load, column, parameters): 
    # extract the values from the netowrk topology and power flow results
    load_table = preallocate_table('load', column, number)
    element = 'load'
    i = 0
    for load_index in net.load.index:
        # from time step, here only the first iteration is donde, so '0'
        load_table.loc[i,'step'] = 0
        # from network topology
        load_table.loc[i,'index'] = load_index 
        load_table.loc[i,parameters['net_' + element]] = net.load.loc[load_index, parameters['net_' + element]]                
        #from power flow results: without time series, directly from net.res_##; with TS, from results (dict/dataframe)
        if time_steps == 1:
            load_table.loc[i,'p_mw'] = net.res_load.loc[load_index,'p_mw']  
            load_table.loc[i,'q_mvar'] = net.res_load.loc[load_index,'q_mvar']
        else:
            load_table.loc[i,'p_mw'] = load['p_mw'].loc[0,load_index] 
            load_table.loc[i,'q_mvar'] = load['q_mvar'].loc[0,load_index] 
        i = i + 1
         
    # results: if there is only '1' time step, directly from the pandapower network and the 'while' loop condiction won't start
    # otherwise the results were called from the temporary excels, and now they will be written from step 1 on
    # first step was 0 (zero)
    print('\n Sorting the results: loads - 1/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['loads'] # offset because all the values from all the loads and time steps are in the same table, 
        i = 0                                #so we write each time step results under the previous one
        for load_index in net.load.index:
            # time step
            load_table.loc[offset + i,'step'] = step
            # from network topology, is repeated for all the time steps
            load_table.loc[offset + i,'index'] = load_index  
            load_table.loc[offset + i,parameters['net_' + element]] = net.load.loc[load_index, parameters['res_' + element]] 
            # from results, each value is taken from a different dataframe within results dict
            load_table.loc[offset + i,'p_mw'] = load['p_mw'].loc[step,load_index]
            load_table.loc[offset + i,'q_mvar'] = load['q_mvar'].loc[step,load_index]
            i = i + 1
        #we go to the next time step  
        step = step + 1
            
    return load_table

def sort_gen_results(net, number, time_steps, gen, gen_fuel_tech, column, parameters): 
    gen_table = preallocate_table('gen', column, number)
    element = 'gen'
    i = 0
    for i in range(number['gen']): # extract the values from power flow results
        gen_table.loc[i,'step'] = 0 
        
        gen_table.loc[i,'index'] = net.gen.index[i]
        gen_table.loc[i,parameters['net_' + element]] = net.gen.loc[i,parameters['net_' + element]]
         
        if time_steps == 1:
            gen_table.loc[i,'p_mw'] = net.res_gen.loc[i,'p_mw']
            gen_table.loc[i,'q_mvar'] = net.res_gen.loc[i,'q_mvar']
        else: 
            gen_table.loc[i,'p_mw'] = gen['p_mw'].loc[0,i] 
            gen_table.loc[i,'q_mvar'] = gen['q_mvar'].loc[0,i]
                    
             
    print('\n Sorting the results: generators - 2/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['gen']
                
        for i in range(number['gen']): 
            gen_table.loc[offset + i,'step'] = step
            gen_table.loc[offset + i,'index'] = net.gen.index[i]
             
            gen_table.loc[offset + i, parameters['net_' + element]] = net.gen.loc[i,parameters['net_' + element]]
             
            gen_table.loc[offset + i,'p_mw'] = gen['p_mw'].loc[step,i] 
            gen_table.loc[offset + i,'q_mvar'] = gen['q_mvar'].loc[step,i]
                
        step = step + 1

    return gen_table


def sort_line_results(net, number, time_steps, line, column, parameters): 
    line_table = preallocate_table('line', column, number)
    element = 'line'
    i = 0                                
    for line_index in net.line.index: # extract the values from power flow results
        line_table.loc[i,'zone'] = 0

        line_table.loc[i,'index'] = line_index
        line_table.loc[i,parameters['net_' + element]] = net.line.loc[line_index, parameters['net_' + element]]
         
        if time_steps == 1:
            line_table.loc[i, parameters['res_' + element]] = net.res_line.loc[line_index ,parameters['res_' + element]]

        else: 
            line_table.loc[i,'p_from_mw'] = line['p_from_mw'].loc[0,line_index ]
            line_table.loc[i,'q_from_mvar'] =  line['q_from_mvar'].loc[0,line_index ]
            line_table.loc[i,'p_to_mw'] =  line['p_to_mw'].loc[0,line_index ]
            line_table.loc[i,'q_to_mvar'] =  line['q_to_mvar'].loc[0,line_index ]
            line_table.loc[i,'pl_mw'] =  line['pl_mw'].loc[0,line_index ]
            line_table.loc[i,'ql_mvar'] =  line['ql_mvar'].loc[0,line_index ]
            line_table.loc[i,'loading_percent'] =  line['loading_percent'].loc[0,line_index ]
        i = i + 1    

    print('\n Sorting the results: lines - 3/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['line']
        i = 0        
        for line_index in net.line.index: 
            line_table.loc[offset + i,'step'] = step
            
            line_table.loc[offset + i,'index'] = line_index
            line_table.loc[offset + i, parameters['res_' + element]] = net.res_line.loc[line_index ,parameters['res_' + element]]
            
            line_table.loc[offset + i,'p_from_mw'] = line['p_from_mw'].loc[step,line_index ]
            line_table.loc[offset + i,'q_from_mvar'] =  line['q_from_mvar'].loc[step,line_index ]
            line_table.loc[offset + i,'p_to_mw'] =  line['p_to_mw'].loc[step,line_index ]
            line_table.loc[offset + i,'q_to_mvar'] =  line['q_to_mvar'].loc[step,line_index ]
            line_table.loc[offset + i,'pl_mw'] =  line['pl_mw'].loc[step,line_index ]
            line_table.loc[offset + i,'ql_mvar'] =  line['ql_mvar'].loc[step,line_index ]
            line_table.loc[offset + i,'loading_percent'] =  line['loading_percent'].loc[step,line_index ]
            i = i + 1     
        step = step + 1

    return line_table 


def sort_trafo_results(net, number, time_steps, trafo, column, parameters): 
    trafo_table = preallocate_table('trafo', column, number)        
    element = 'trafo'    
    i = 0 
    for trafo_index in net.trafo.index: # extract the values from power flow results
        trafo_table.loc[i,'step'] = 0 
        
        trafo_table.loc[i,'index'] = trafo_index
        trafo_table.loc[i,parameters['net_' + element]] = net.trafo.loc[trafo_index, parameters['net_' + element]]
        
        if time_steps == 1:                
            trafo_table.loc[i,parameters['res_' + element]] = net.res_trafo.loc[trafo_index, parameters['res_' + element]]
        else: 
            trafo_table.loc[i,'p_hv_mw'] = trafo['p_hv_mw'].loc[0,trafo_index] 
            trafo_table.loc[i,'q_hv_mvar'] = trafo['q_hv_mvar'].loc[0,trafo_index] 
            trafo_table.loc[i,'p_lv_mw'] = trafo['p_lv_mw'].loc[0,trafo_index] 
            trafo_table.loc[i,'q_lv_mvar'] = trafo['q_lv_mvar'].loc[0,trafo_index] 
            trafo_table.loc[i,'pl_mw'] = trafo['pl_mw'].loc[0,trafo_index] 
            trafo_table.loc[i,'ql_mvar'] = trafo['ql_mvar'].loc[0,trafo_index] 
            trafo_table.loc[i,'loading_percent'] = trafo['loading_percent'].loc[0,trafo_index] 
        i = i + 1
            
    print('\n Sorting the results: trafos - 4/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['trafos']
        i = 0        
        for trafo_index in net.trafo.index: 
            trafo_table.loc[offset + i,'step'] = step
            
            trafo_table.loc[offset + i,'index'] = trafo_index
            trafo_table.loc[offset + i,parameters['net_' + element]] = net.trafo.loc[trafo_index, parameters['net_' + element]]
            
            trafo_table.loc[offset + i,'p_hv_mw'] = trafo['p_hv_mw'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'q_hv_mvar'] = trafo['q_hv_mvar'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'p_lv_mw'] = trafo['p_lv_mw'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'q_lv_mvar'] = trafo['q_lv_mvar'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'pl_mw'] = trafo['pl_mw'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'ql_mvar'] = trafo['ql_mvar'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'loading_percent'] = trafo['loading_percent'].loc[step,trafo_index] 
            i = i + 1
        step = step + 1

    return trafo_table


def sort_bus_results(net, number, time_steps, bus, column, parameters):    
    bus_table = preallocate_table('bus', column, number)
    element = 'bus'
    i = 0     
    for bus_index in net.bus.index:
        bus_table.loc[i,'step'] = 0
        
        bus_table.loc[i,'index'] = bus_index
        bus_table.loc[i,parameters['net_' + element]] = net.bus.loc[bus_index, parameters['net_' + element]]
        
        if time_steps == 1:
            bus_table.loc[i,parameters['res_' + element]] = net.res_bus.loc[bus_index, parameters['res_' + element]]
        else: 
            bus_table.loc[i,'vm_pu'] = bus['vm_pu'].loc[0,bus_index]
            bus_table.loc[i,'va_degree'] = bus['va_degree'].loc[0,bus_index]
            bus_table.loc[i,'p_mw'] = bus['p_mw'].loc[0,bus_index]
            bus_table.loc[i,'q_mvar'] = bus['q_mvar'].loc[0,bus_index]
        i = i + 1
    
    print('\n Sorting the results: buses - 5/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['buses']
        i = 0
        for bus_index in net.bus.index: 
            bus_table.loc[offset + i,'step'] = step
            
            bus_table.loc[offset + i,'index'] = bus_index
            bus_table.loc[offset + i,parameters['net_' + element]] = net.bus.loc[bus_index, parameters['net_' + element]]
            
            bus_table.loc[offset + i,'vm_pu'] = bus['vm_pu'].loc[step,bus_index]
            bus_table.loc[offset + i,'va_degree'] = bus['va_degree'].loc[step,bus_index]
            bus_table.loc[offset + i,'p_mw'] = bus['p_mw'].loc[step,bus_index]
            bus_table.loc[offset + i,'q_mvar'] = bus['q_mvar'].loc[step,bus_index]      
            i = i + 1
        step = step + 1
        
    return bus_table















