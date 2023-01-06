# -*- coding: utf-8 -*-
"""
Created on Thu Jan  5 22:17:16 2023

@author: marti
"""


import pandas as pd
import pandapower as pp

def output_parameters(net):#, Sum_Bus_Vol_Under_Data):
       
    # Preallocate values: number of loads/gen/buses/trafos/lines, and columns in the excel template and their names
    # based on the network topology, the quentity of elements are counted
    number = { 'load' : len(net.load),
               'gen' : len(net.gen),
               'line' : len(net.line),
               'bus' : len(net.bus),
               'trafo' : len(net.trafo)}#,
               #'summary' : len(Sum_Bus_Vol_Under_Data) }
    
    # according to the desired excel output, the corresponding columns are assigning to each element
    column = {'letter' : {},
              'parameter' : {} }
    column['letter'] = { 'load' : ['C','D','E','F','G','H','I'],
                        'gen' : ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q'],
                        'line' : ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U'],
                        'trafo' : ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W'],
                        'bus' : ['C','D','E','F','G','H','I','J','K','L'],
                        'summary' : ['C', 'D' ,'E', 'F', 'G'] }
    
    
    #  according to the desired excel output, the corresponding output_parameters by element are selected
    parameters = {'net_bus' : ['zone','name','vn_kv','in_service'],
                  'net_load' : ['zone','bus','in_service'],
                  'net_gen' : ['zone','bus','name','in_service','vm_pu', 'max_p_mw', 'max_q_mvar','min_p_mw','min_q_mvar'],
                  'net_line' : ['zone','name','from_bus','to_bus', 'in_service','length_km', 'max_i_ka','max_loading_percent', 'parallel', 'std_type'],
                  'net_trafo' : ['zone','name', 'std_type', 'hv_bus', 'lv_bus', 'vn_hv_kv', 'vn_lv_kv', 'pfe_kw', 'shift_degree','tap_pos', 'parallel', 'in_service' ],
                  
                  'res_bus' : ['vm_pu','va_degree','p_mw','q_mvar'],           
                  'res_load' : ['p_mw','q_mvar'],
                  'res_gen' : ['p_mw','q_mvar'],
                  'res_line' : ['p_from_mw', 'q_from_mvar', 'p_to_mw', 'q_to_mvar', 'pl_mw', 'ql_mvar', 'loading_percent'],
                  'res_trafo' : ['p_hv_mw','q_hv_mvar', 'p_lv_mw', 'q_lv_mvar', 'pl_mw', 'ql_mvar', 'loading_percent'],
                  'summary' : ['component','Percentage','Extra Info'] }
    
    for element in number:
        if number[element] > 0:
            pp.add_zones_to_elements(net, replace=True, elements = element) 
            column['parameter'][element] = ['step','index'] + parameters['net_' + element] + parameters['res_' + element]


    return number, column, parameters

def preallocate_table(element, column, number):
    # preallocating the columns names with the parameters from load, this is our desired ordered excel output
    table = {} 
    for column_t in column['parameter'][element]:
        table[column_t] = [0]*number[element]
        
    table_df = pd.DataFrame(data = table) 

    return table_df

def check_load(net, time_steps, parameters, number):
    for i in range(len(parameters['net_load'])): 
        if not parameters['net_load'][i] in net.load.keys():  # check the column of parameters
        # if the columm of parameter is empty, it creat that columm and put value as none
            net.load[parameters['net_load'][i]] = [None]*number['load'] 
    if time_steps == 1: # only check if there is no time series,
    # as for time series the selection of output is done by the functions: temp_files_to_excel_input, create_output_writer
        for i in range(len(parameters['res_load'])): # check the column of parameters
             if not parameters['res_load'][i] in net.res_load.keys():  # check the colum of parameter
             # if the columm of parameter is empty, it creates that columm and put value as none 
                 net.res_load[parameters['res_load'][i]] = [None]*number['load']
                 
    return net

def check_bus(net, time_steps, parameters, number):
    for i in range(len(parameters['net_bus'])):
         if not parameters['net_bus'][i] in net.bus.keys():
             net.bus[parameters['net_bus'][i]] = [None]*number['bus']
    if time_steps ==1:
        for i in range(len(parameters['res_bus'])):
            if not parameters['res_bus'][i] in net.res_bus.keys():
                net.res_bus[parameters['res_bus'][i]] = [None]*number['bus']      
    
    return net

def check_gen(net, time_steps, parameters, number, gen_fuel_tech):
    for i in range(len(parameters['net_gen'])):
         if not parameters['net_gen'][i] in net.gen.keys():
             net.gen[parameters['net_gen'][i]] = [None]*number['gen']
    if time_steps == 1:
        for i in range(len(parameters['res_gen'])):
             if not parameters['res_gen'][i] in net.res_gen.keys():
                 net.res_gen[parameters['res_gen'][i]] = [None]*number['gen']         
    if not len(gen_fuel_tech) > 0:
         d = {'fuel': ['---'], 'tech': ['---']}
         gen_fuel_tech = pd.DataFrame(data = d)
         df_aux = pd.DataFrame(data = d)
         for i in range(number['gen']):
             gen_fuel_tech = pd.concat([gen_fuel_tech, df_aux],ignore_index=True)  
    
    return net

def check_line(net, time_steps, parameters, number):
    for i in range(len(parameters['net_line'])):
         if not parameters['net_line'][i] in net.line.keys():
             net.line[parameters['net_line'][i]] = [None]*number['line']
    if time_steps == 1:
         for i in range(len(parameters['res_line'])):
             if not parameters['res_line'][i] in net.res_line.keys():
                 net.res_line[parameters['res_line'][i]] = [None]*number['line']     

    return net

def check_trafo(net, time_steps, parameters, number):
    for i in range(len(parameters['net_trafo'])):
         if not parameters['net_trafo'][i] in net.trafo.keys():
             net.trafo[parameters['net_trafo'][i]] = [None]*number['trafo']
    if time_steps == 1:
        for i in range(len(parameters['res_trafo'])):
             if not parameters['res_trafo'][i] in net.res_trafo.keys():
                 net.res_trafo[parameters['res_trafo'][i]] = [None]*number['trafo']      
    
    return net

