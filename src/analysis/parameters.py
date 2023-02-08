# -*- coding: utf-8 -*-
"""
Created on Thu Jan  5 22:17:16 2023

@author: marti
"""


import pandas as pd
import pandapower as pp

def sheets_parameters():
    # dict that returns the different types of elements at each sheet
    # so the keys() is the the name of the sheets
    elements_by_type = {'Buses' :['bus'],
                        'Demand' : ['load'],
                        'Generation' : ['gen', 'sgen', 'ext_grid', 'storage'],
                        'Lines' : ['line'],
                        'Trafos' : ['trafo']}
    # each output sheet contains the same type of elements, so we can analyze the whole group 
    # e.g. generation contains all the elements that act as supply in the network
    # if a new element want to be added, it is enough to add the name in the list, and the parameters
    # for the output would be the same for all the element in the same type of element
    sheets_names = elements_by_type.keys()
    cell = {} # dictionary with the different cell values for excel
    cell['initial'] = 'C4' # cell in the template were all the titles start
    # add the initial cell value to each sheet
    sheet = add_initial_cell (sheets_names, cell, elements_by_type)            
    return sheet, cell, elements_by_type

def add_initial_cell(sheets_names, cell, elements_by_type):
    sheet = {} # dict that returns the sheet of each element
    for sheet_ in sheets_names:
        cell[sheet_] = cell['initial']
        for element in elements_by_type[sheet_]:
            sheet[element] = sheet_ 
    return sheet 

def select_parameters():
    # according to the desired excel output, the corresponding output_parameters by element are selected
    # only built it parameters can be added directly, in order to add parameters like vn_kv to the loads, or gen, they should also
    # be added in the 'ouput_parameters' function, similiar how it is done with 'zone', 'fuel'
    # if a element doesn't contain a parameter added here, it will appear '---' in the excel output
    parameters = {'net_bus' : ['scenario','zone','name','vn_kv','in_service'],
                  'net_load' : ['scenario','zone','name','bus','in_service'],
                  'net_gen' : ['scenario','zone','name','type','bus','in_service','fuel','vm_pu', 'max_p_mw', 'max_q_mvar','min_p_mw','min_q_mvar'],
                  'net_line' : ['scenario','zone','name','from_bus','to_bus', 'in_service','length_km', 'max_i_ka','max_loading_percent', 'parallel', 'std_type'],
                  'net_trafo' : ['scenario','zone','name', 'std_type', 'hv_bus', 'lv_bus', 'vn_hv_kv', 'vn_lv_kv', 'pfe_kw', 'shift_degree','tap_pos', 'parallel', 'in_service' ],
                  
                  'res_bus' : ['vm_pu','va_degree','p_mw','q_mvar'],           
                  'res_load' : ['p_mw','q_mvar'],
                  'res_gen' : ['p_mw','q_mvar'],
                  'res_line' : ['p_from_mw', 'q_from_mvar', 'p_to_mw', 'q_to_mvar', 'pl_mw', 'ql_mvar', 'loading_percent'],
                  'res_trafo' : ['p_hv_mw','q_hv_mvar', 'p_lv_mw', 'q_lv_mvar', 'pl_mw', 'ql_mvar', 'loading_percent']}      
    return parameters

def output_parameters(net, gen_fuel_tech, scenario_name):#, Sum_Bus_Vol_Under_Data):      
    # Preallocate values: number of loads/gen/buses/trafos/lines, and columns in the excel template and their names
    # based on the network topology, the quentity of elements are counted
    sheet, cell, elements_by_type = sheets_parameters()
    
    elements = []
    for element_type in elements_by_type:
        elements = elements + elements_by_type[element_type]
    
    number ={}
    for element in elements:  
        number[element] = len(net[element]) # count the number in each type of element
        if len(scenario_name) > 0:
            net[element]['scenario'] = scenario_name # add the scenario name to the net that is being calculated now 
            if element in elements_by_type['Generation']: net[element]['type'] = element # add the type of generation
    # add fuel to the generators
    net = add_fuel(net, gen_fuel_tech, elements_by_type) 
    # according to the desired excel output, the corresponding output_parameters by element are selected
    parameters = select_parameters()

    #parameters that just include the same of other already written, 
    #but might not have all of them, in that case '---' will be written
    for element_type in elements_by_type:
        if len(elements_by_type[element_type]) > 1:
            for element in elements_by_type[element_type]:      
                parameters['net_'+element] = parameters['net_'+elements_by_type[element_type][0]] 
                parameters['res_'+element] = parameters['res_'+elements_by_type[element_type][0]]
    # according to the desired excel output, the corresponding columns are assigning to each element
    column = {'parameter' : {} }
    for element in number:
        # if number[element] > 0:
        if len(scenario_name) > 0: pp.add_zones_to_elements(net, replace=True, elements = element) 
        column['parameter'][element] = ['step','index'] + parameters['net_' + element] + parameters['res_' + element]            

    return number, column, parameters  

def preallocate_table(element, column, number):
    # preallocating the columns names with the parameters from load, this is our desired ordered excel output
    table = {} 
    for column_t in column['parameter'][element]:
        # print(column_t)
        n = number[element]
        if n == 0: n=1
        table[column_t] = [0]*n   
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

def add_fuel(net,gen_fuel_tech, elements_by_type):    
    if len(gen_fuel_tech) > 0:
        for i in gen_fuel_tech.index:
            element = gen_fuel_tech['gen_type'][i]
            index = gen_fuel_tech['index'][i]
            fuel = gen_fuel_tech['fuel'][i]
            if not 'fuel' in net[element].keys(): net[element]['fuel'] = 0
            net[element].loc[index,'fuel'] = fuel
        try:
            fuel = net['ext_grid']['fuel']
        except:
            n = len(net['ext_grid'].index)
            ext_grid = ['External Grid']*n
            net['ext_grid']['fuel'] = ext_grid   
    else:
        for element in elements_by_type['Generation']:
            net[element]['fuel'] = '---'
    return net

def letters():
    letter = ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    return letter
    
def preallocate_tables(input_setup):
    tables = {}
    for scenario_name, scenario_path, pd_scenario, pd_para in input_setup['scenario_setup']:
        tables[scenario_name] = {}    
    return tables
    
    
    
    
    