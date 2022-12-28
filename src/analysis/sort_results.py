# -*- coding: utf-8 -*-
"""
Created on Wed Dec 28 16:45:28 2022

@author: marti
"""
# import pandas as pd
#from openpyxl import Workbook
# from openpyxl import load_workbook
from tqdm import tqdm

def sort_load_results(net, number, load_table, time_steps, load): 
    # extract the values from the netowrk topology and power flow results
    i = 0
    for load_index in net.load.index:
        # from time step, here only the first iteration is donde, so '0'
        load_table.loc[i,'step'] = 0
        # from network topology
        bus_index = net.load.loc[load_index,'bus']
        bus_row = net.bus.index.to_list().index(bus_index) 
        load_table.loc[i,'zone'] = net.bus.iloc[bus_row]['zone']    
        load_table.loc[i,'load_index'] = load_index                
        load_table.loc[i,'bus_index'] = net.load.loc[load_index,'bus']     
        load_table.loc[i,'in_service'] = net.load.loc[load_index,'in_service']    
        load_table.loc[i,'load_voltage'] = net.bus.iloc[bus_row]['vn_kv']             
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
            bus_index = net.load.loc[i,'bus']
            bus_row = net.bus.index.to_list().index(bus_index) 
            load_table.loc[offset + i,'zone'] = net.bus.iloc[bus_row]['zone']    
            load_table.loc[offset + i,'load_index'] = load_index                 
            load_table.loc[offset + i,'bus_index'] = net.load.loc[load_index,'bus']     
            load_table.loc[offset + i,'in_service'] = net.load.loc[load_index,'in_service']    
            load_table.loc[offset + i,'load_voltage'] = net.bus.iloc[bus_row]['vn_kv'] 
            # from results, each value is taken from a different dataframe within results dict
            load_table.loc[offset + i,'p_mw'] = load['p_mw'].loc[step,load_index]
            load_table.loc[offset + i,'q_mvar'] = load['q_mvar'].loc[step,load_index]
            i = i + 1
        #we go to the next time step  
        step = step + 1
            
    return load_table

def sort_gen_results(net, number, gen_table, time_steps, gen, gen_fuel_tech): 
                    
    for i in range(number['generators']): # extract the values from power flow results
        gen_table.loc[i,'step'] = 0 
        
        bus_index = net.gen.loc[i,'bus']
        bus_row = net.bus.index.to_list().index(bus_index) 
        gen_table.loc[i,'zone'] = net.bus.iloc[bus_row]['zone']
        gen_table.loc[i,'bus_index'] = net.gen.loc[i,'bus']
        gen_table.loc[i,'gen_index'] = net.gen.index[i]
        
        gen_table.loc[i,'fuel'] = gen_fuel_tech.loc[i,'fuel']
        gen_table.loc[i,'tech'] = gen_fuel_tech.loc[i,'tech']
        
        gen_table.loc[i,'voltage'] = net.bus.iloc[bus_row]['vn_kv']
         
        gen_table.loc[i,'name'] = net.gen.loc[i,'name']
        gen_table.loc[i,'in_service'] =  net.gen.loc[i,'in_service']
        gen_table.loc[i,'vm_pu'] =  net.gen.loc[i,'vm_pu']
        gen_table.loc[i,'max_p_mw'] =  net.gen.loc[i,'max_p_mw']
        gen_table.loc[i,'max_q_mvar'] =  net.gen.loc[i,'max_q_mvar']
        gen_table.loc[i,'min_p_mw'] =  net.gen.loc[i,'min_p_mw']
        gen_table.loc[i,'min_q_mvar'] =  net.gen.loc[i,'min_q_mvar']
         
        if time_steps == 1:
            gen_table.loc[i,'p_mw'] = net.res_gen.loc[i,'p_mw']
            gen_table.loc[i,'q_mvar'] = net.res_gen.loc[i,'q_mvar']
        else: 
            gen_table.loc[i,'p_mw'] = gen['p_mw'].loc[0,i] 
            gen_table.loc[i,'q_mvar'] = gen['q_mvar'].loc[0,i]
                    
             
    print('\n Sorting the results: generators - 2/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['generators']
                
        for i in range(number['generators']): 
            gen_table.loc[offset + i,'step'] = step
            
            bus_index = net.gen.loc[i,'bus']
            bus_row = net.bus.index.to_list().index(bus_index) 
            gen_table.loc[offset + i,'zone' ] = net.bus.iloc[bus_row]['zone']
            gen_table.loc[offset + i,'bus_index'] = net.gen.loc[i,'bus']
            gen_table.loc[offset + i,'gen_index'] = net.gen.index[i]
             
            gen_table.loc[offset + i,'fuel'] = gen_fuel_tech.loc[i,'fuel']
            gen_table.loc[offset + i,'tech'] = gen_fuel_tech.loc[i,'tech']
             
            gen_table.loc[offset + i,'voltage'] = net.bus.iloc[bus_row]['vn_kv']
             
            gen_table.loc[offset + i,'name'] = net.gen.loc[i,'name']
            gen_table.loc[offset + i,'in_service'] =  net.gen.loc[i,'in_service']
            gen_table.loc[offset + i,'vm_pu'] =  net.gen.loc[i,'vm_pu']
            gen_table.loc[offset + i,'max_p_mw'] =  net.gen.loc[i,'max_p_mw']
            gen_table.loc[offset + i,'max_q_mvar'] =  net.gen.loc[i,'max_q_mvar']
            gen_table.loc[offset + i,'min_p_mw'] =  net.gen.loc[i,'min_p_mw']
            gen_table.loc[offset + i,'min_q_mvar'] =  net.gen.loc[i,'min_q_mvar']
             
            gen_table.loc[offset + i,'p_mw'] = gen['p_mw'].loc[step,i] 
            gen_table.loc[offset + i,'q_mvar'] = gen['q_mvar'].loc[step,i]
                
        step = step + 1

    return gen_table


def sort_line_results(net, number, line_table, time_steps, line): 
    i = 0                                
    for line_index in net.line.index: # extract the values from power flow results
        line_table.loc[i,'zone'] = 0
    
        bus_index = net.line.loc[line_index ,'from_bus']
        bus_row = net.bus.index.to_list().index(bus_index) 
        line_table.loc[i,'zone'] = net.bus.iloc[bus_row]['zone']
        line_table.loc[i,'line_index'] = line_index
        line_table.loc[i,'voltage'] = net.bus.iloc[bus_row]['vn_kv']
        line_table.loc[i,'name'] = net.line.loc[line_index ,'name']
        line_table.loc[i,'from_bus'] = net.line.loc[line_index ,'from_bus']
        line_table.loc[i,'to_bus'] = net.line.loc[line_index ,'to_bus']
        line_table.loc[i,'in_service'] = net.line.loc[line_index ,'in_service']
        line_table.loc[i,'length_km'] = net.line.loc[line_index ,'length_km']
        line_table.loc[i,'max_i_ka'] = net.line.loc[line_index ,'max_i_ka']
        line_table.loc[i,'max_loading_percent'] = net.line.loc[line_index ,'max_loading_percent']
        line_table.loc[i,'parallel'] = net.line.loc[line_index ,'parallel']
        line_table.loc[i,'std_type'] = net.line.loc[line_index ,'std_type']
         
        if time_steps == 1:
            line_table.loc[i,'p_from_mw'] = net.res_line.loc[line_index ,'p_from_mw']
            line_table.loc[i,'q_from_mvar'] = net.res_line.loc[line_index ,'q_from_mvar']
            line_table.loc[i,'p_to_mw'] = net.res_line.loc[line_index ,'p_to_mw']
            line_table.loc[i,'q_to_mvar'] = net.res_line.loc[line_index ,'q_to_mvar']
            line_table.loc[i,'pl_mw'] = net.res_line.loc[line_index ,'pl_mw']
            line_table.loc[i,'ql_mvar'] = net.res_line.loc[line_index ,'ql_mvar']
            line_table.loc[i,'loading_percent'] = net.res_line.loc[line_index ,'loading_percent']
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
        offset = step*number['lines']
        i = 0        
        for line_index in net.line.index: 
            line_table.loc[offset + i,'step'] = step
            
            bus_index = net.line.loc[line_index,'from_bus']
            bus_row = net.bus.index.to_list().index(bus_index) 
            line_table.loc[offset + i,'zone'] = net.bus.iloc[bus_row]['zone']
            line_table.loc[offset + i,'line_index'] = line_index
            line_table.loc[offset + i,'voltage'] = net.bus.iloc[bus_row]['vn_kv']
            line_table.loc[offset + i,'name'] = net.line.loc[line_index ,'name']
            line_table.loc[offset + i,'from_bus'] = net.line.loc[line_index ,'from_bus']
            line_table.loc[offset + i,'to_bus'] = net.line.loc[line_index ,'to_bus']
            line_table.loc[offset + i,'in_service'] = net.line.loc[line_index ,'in_service']
            line_table.loc[offset + i,'length_km'] = net.line.loc[line_index ,'length_km']
            line_table.loc[offset + i,'max_i_ka'] = net.line.loc[line_index ,'max_i_ka']
            line_table.loc[offset + i,'max_loading_percent'] = net.line.loc[line_index ,'max_loading_percent']
            line_table.loc[offset + i,'parallel'] = net.line.loc[line_index ,'parallel']
            line_table.loc[offset + i,'std_type'] = net.line.loc[line_index ,'std_type']
  
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


def sort_trafo_results(net, number, trafo_table, time_steps, trafo):             
    i = 0 
    for trafo_index in net.trafo.index: # extract the values from power flow results
        trafo_table.loc[i,'step'] = 0 
        
        bus_index = net.trafo.loc[trafo_index,'hv_bus']
        bus_row = net.bus.index.to_list().index(bus_index) 
        trafo_table.loc[i,'zone'] = net.bus.iloc[bus_row]['zone']
        trafo_table.loc[i,'trafo_index'] = trafo_index
        trafo_table.loc[i,'name'] = net.trafo.loc[trafo_index,'name']
        trafo_table.loc[i,'std_type'] = net.trafo.loc[trafo_index,'std_type']
        trafo_table.loc[i,'hv_bus'] = net.trafo.loc[trafo_index,'hv_bus']
        trafo_table.loc[i,'lv_bus'] = net.trafo.loc[trafo_index,'lv_bus']
        trafo_table.loc[i,'vn_hv_kv'] = net.trafo.loc[trafo_index,'vn_hv_kv']
        trafo_table.loc[i,'vn_lv_kv'] = net.trafo.loc[trafo_index,'vn_lv_kv']
        trafo_table.loc[i,'pfe_kw'] = net.trafo.loc[trafo_index,'pfe_kw']
        trafo_table.loc[i,'shift_degree'] = net.trafo.loc[trafo_index,'shift_degree']
        trafo_table.loc[i,'tap_pos'] = net.trafo.loc[trafo_index,'tap_pos']
        trafo_table.loc[i,'parallel'] = net.trafo.loc[trafo_index,'parallel']
        trafo_table.loc[i,'in_service'] = net.trafo.loc[trafo_index,'in_service']
        
        if time_steps == 1:                
            trafo_table.loc[i,'p_hv_mw'] = net.res_trafo.loc[trafo_index,'p_hv_mw']
            trafo_table.loc[i,'q_hv_mvar'] = net.res_trafo.loc[trafo_index,'q_hv_mvar']
            trafo_table.loc[i,'p_lv_mw'] = net.res_trafo.loc[trafo_index,'p_lv_mw']
            trafo_table.loc[i,'q_lv_mvar'] = net.res_trafo.loc[trafo_index,'q_lv_mvar']
            trafo_table.loc[i,'pl_mw'] = net.res_trafo.loc[trafo_index,'pl_mw']
            trafo_table.loc[i,'ql_mvar'] =net.res_trafo.loc[trafo_index,'ql_mvar']
            trafo_table.loc[i,'loading_percent'] = net.res_trafo.loc[trafo_index,'loading_percent']
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
            
            bus_index = net.trafo.loc[trafo_index,'hv_bus']
            bus_row = net.bus.index.to_list().index(bus_index) 
            trafo_table.loc[offset + i,'zone'] = net.bus.iloc[bus_row]['zone']
            trafo_table.loc[offset + i,'trafo_index'] = trafo_index
            trafo_table.loc[offset + i,'name'] = net.trafo.loc[trafo_index,'name']
            trafo_table.loc[offset + i,'std_type'] = net.trafo.loc[trafo_index,'std_type']
            trafo_table.loc[offset + i,'hv_bus'] = net.trafo.loc[trafo_index,'hv_bus']
            trafo_table.loc[offset + i,'lv_bus'] = net.trafo.loc[trafo_index,'lv_bus']
            trafo_table.loc[offset + i,'vn_hv_kv'] = net.trafo.loc[trafo_index,'vn_hv_kv']
            trafo_table.loc[offset + i,'vn_lv_kv'] = net.trafo.loc[trafo_index,'vn_lv_kv']
            trafo_table.loc[offset + i,'pfe_kw'] = net.trafo.loc[trafo_index,'pfe_kw']
            trafo_table.loc[offset + i,'shift_degree'] = net.trafo.loc[trafo_index,'shift_degree']
            trafo_table.loc[offset + i,'tap_pos'] = net.trafo.loc[trafo_index,'tap_pos']
            trafo_table.loc[offset + i,'parallel'] = net.trafo.loc[trafo_index,'parallel']
            trafo_table.loc[offset + i,'in_service'] = net.trafo.loc[trafo_index,'in_service']
               
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


def sort_bus_results(net, number, bus_table, time_steps, bus):    
    i = 0     
    for bus_index in net.bus.index:
        bus_table.loc[i,'step'] = 0
        
        bus_table.loc[i,'index'] = bus_index
        bus_table.loc[i,'zone'] = net.bus.loc[bus_index,'zone']
        bus_table.loc[i,'name'] = net.bus.loc[bus_index,'name']
        bus_table.loc[i,'vn_kv'] = net.bus.loc[bus_index,'vn_kv']
        bus_table.loc[i,'in_service'] = net.bus.loc[bus_index,'in_service']
        
        if time_steps == 1:
            bus_table.loc[i,'vm_pu'] = net.res_bus.loc[bus_index,'vm_pu']
            bus_table.loc[i,'va_degree'] = net.res_bus.loc[bus_index,'va_degree']
            bus_table.loc[i,'p_mw'] = net.res_bus.loc[bus_index,'p_mw']
            bus_table.loc[i,'q_mvar'] = net.res_bus.loc[bus_index,'q_mvar']
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
            bus_table.loc[offset + i,'zone'] = net.bus.loc[bus_index,'zone']
            bus_table.loc[offset + i,'name'] = net.bus.loc[bus_index,'name']
            bus_table.loc[offset + i,'vn_kv'] = net.bus.loc[bus_index,'vn_kv']
            bus_table.loc[offset + i,'in_service'] = net.bus.loc[bus_index,'in_service']
            
            bus_table.loc[offset + i,'vm_pu'] = bus['vm_pu'].loc[step,bus_index]
            bus_table.loc[offset + i,'va_degree'] = bus['va_degree'].loc[step,bus_index]
            bus_table.loc[offset + i,'p_mw'] = bus['p_mw'].loc[step,bus_index]
            bus_table.loc[offset + i,'q_mvar'] = bus['q_mvar'].loc[step,bus_index]      
            i = i + 1
        step = step + 1
        
    return bus_table















