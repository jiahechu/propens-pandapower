# -*- coding: utf-8 -*-
"""
Created on Thu Jan  5 22:17:16 2023

@author: marti
"""


import pandas as pd
import pandapower as pp

def output_parameters(net, gen_fuel_tech):#, Sum_Bus_Vol_Under_Data):
       
    # Preallocate values: number of loads/gen/buses/trafos/lines, and columns in the excel template and their names
    # based on the network topology, the quentity of elements are counted
    elements = ['bus','load','gen','sgen','line','trafo']
    generation = ['gen', 'sgen']
    number ={}
    
    for element in elements:  
        number[element] = len(net[element]) # count the number in each type of element
        if element in generation: net[element]['type'] = element # add the type of generation
    #'summary' : len(Sum_Bus_Vol_Under_Data) }
    
    net = add_fuel(net, gen_fuel_tech) # add fuel to the generators
    
    # according to the desired excel output, the corresponding columns are assigning to each element
    column = {'letter' : {},
              'parameter' : {} }
    column['letter'] = { 'load' : ['C','D','E','F','G','H','I'],
                        'gen' : ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R'],
                        'sgen' : ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R'],
                        'line' : ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U'],
                        'trafo' : ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W'],
                        'bus' : ['C','D','E','F','G','H','I','J','K','L'],
                        'summary' : ['C', 'D' ,'E', 'F', 'G'] }
    
    
    #  according to the desired excel output, the corresponding output_parameters by element are selected
    parameters = {'net_bus' : ['zone','name','vn_kv','in_service'],
                  'net_load' : ['zone','bus','in_service'],
                  'net_gen' : ['type','zone','bus','name','in_service','fuel','vm_pu', 'max_p_mw', 'max_q_mvar','min_p_mw','min_q_mvar'],
                  'net_line' : ['zone','name','from_bus','to_bus', 'in_service','length_km', 'max_i_ka','max_loading_percent', 'parallel', 'std_type'],
                  'net_trafo' : ['zone','name', 'std_type', 'hv_bus', 'lv_bus', 'vn_hv_kv', 'vn_lv_kv', 'pfe_kw', 'shift_degree','tap_pos', 'parallel', 'in_service' ],
                  
                  'res_bus' : ['vm_pu','va_degree','p_mw','q_mvar'],           
                  'res_load' : ['p_mw','q_mvar'],
                  'res_gen' : ['p_mw','q_mvar'],
                  'res_line' : ['p_from_mw', 'q_from_mvar', 'p_to_mw', 'q_to_mvar', 'pl_mw', 'ql_mvar', 'loading_percent'],
                  'res_trafo' : ['p_hv_mw','q_hv_mvar', 'p_lv_mw', 'q_lv_mvar', 'pl_mw', 'ql_mvar', 'loading_percent'],
                  'summary' : ['component','Percentage','Extra Info'] }
    #parameters that just include the same of other already written, 
    #but might not have all of them, in that case '---' will be written
    parameters['net_sgen'] = parameters['net_gen'] 
    parameters['res_sgen'] = parameters['res_gen']
    
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

def check_parameter(net, time_steps, parameters, number, element):
    for i in range(len(parameters['net_' + element])): 
        if not parameters['net_'+ element][i] in net[element].keys():  # check the column of parameters
        # if the columm of parameter is empty, it creat that columm and put value as none
            net[element][parameters['net_'+ element][i]] = [None]*number[element] 
    if time_steps == 1: # only check if there is no time series,
    # as for time series the selection of output is done by the functions: temp_files_to_excel_input, create_output_writer
        for i in range(len(parameters['res_'+ element])): # check the column of parameters
             if not parameters['res_'+ element][i] in net['res_'+element].keys():  # check the colum of parameter
             # if the columm of parameter is empty, it creates that columm and put value as none 
                 net['res_' + element][parameters['res_'+ element][i]] = [None]*number[element] 
    return net        

def add_fuel(net,gen_fuel_tech):
    if len(gen_fuel_tech) > 0:
        for i in gen_fuel_tech.index:
            element = gen_fuel_tech['gen_type'][i]
            index = gen_fuel_tech['gen_type'][i]
            fuel = gen_fuel_tech['fuel'][i]
            if not 'fuel' in net[element].keys(): net[element]['fuel'] = 0
            net[element].loc[index,'fuel'] = fuel
            
    return net
            