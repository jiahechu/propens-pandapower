# -*- coding: utf-8 -*-
"""
Created on Sat Jan 14 11:53:24 2023

@author: marti
"""
from src.analysis.parameters import check_parameter
from src.analysis.parameters import output_parameters
from tqdm import tqdm
from src.analysis.parameters import preallocate_table

#%%
def save_results(net, gen_fuel_tech, scenario_name, time_steps, results):
    print('\n Sorting the results')
    tables = {}
    number, column, parameters = output_parameters(net, gen_fuel_tech, scenario_name)
    
    for element in number:
        if number[element] > 0:       
            # checking values of the parameters, and adding columns and/or formatting them    
            net = check_parameter(net, time_steps, parameters, number, element)  
            print(' > of the '+ element+'s')
        else:
            # if there is no data i.e. no generator just an external grid 
            # it isnt necesity to check the results  
            # results[element] = pd.DataFrame()
            print(' > no '+ element +'s in the net')
        # read the values from results, and sort them for the output
        try:
            tables[element] = sort_results(net, number, time_steps, results[element], column, parameters, element, scenario_name)
        except:
            print('\n ----    Error in the sort_results function input -------')
            print('>>>>>>>>> Element : ' + element)
            print('>>>>>>>>> Table content: ' + str(tables.keys()))
            raise
    return tables

def sort_results(net, number, time_steps, element_results, column, parameters, element, scenario_name): 
    # extract the values from the netowrk topology and power flow results
    table = preallocate_table(element, column, number)
    if net[element].shape[0] == 0:
        table['scenario'][0] = scenario_name
        if 'type' in table.keys():  table['type'][0] = element
        return table
    i = 0
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
