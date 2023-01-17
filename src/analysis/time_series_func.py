# -*- coding: utf-8 -*-
"""
Created on Wed Nov 23 09:04:03 2022

@author: thoug
"""
#https://github.com/e2nIEE/pandapower/blob/develop/tutorials/time_series.ipynb

import os
import pandas as pd
import tempfile

from pandapower.timeseries import OutputWriter #To write the outputs to HDD
from pandapower.timeseries.run_time_series import run_timeseries #main time series function, calls the controller function to update P,Q value and run pp

from src.analysis.parameters import output_parameters
#%%
'''
read the temporary files created in the running of the time series
analysis, and convert them into a dataframe that will be sent to 
create_excel function
'''

def temp_files_to_excel_input(output_dir, parameters):
    # output_dir : directory path for the temporal data
    # res = res_bus / res_lines / res_trafos / etc
    # myvar : dictionary with the output_parameters result 
    #       for each element e.g. line : { p_mw : [100 , 101, 99, ..]
    #                                      q_mvar : [50 , 50, 20, ..]
    # output_parameter = p_mw / q_mw from the lines for example
    # results : dictionary with all the elements results
    results = {}
    # results_v02 = pd.DataFrame()
    # element = pd.DataFrame()
    
    for net_res_element in parameters:
        res = net_res_element[0:len('res')] # considering only the result parameters, not network topology (net)
        
        if res == 'res' :
            res_element = net_res_element
            network_element_str = res_element[(len('res')+1):] # creating the variable as string
            local_vars = locals() # converting the string to a variable   
            local_vars[network_element_str] =  {}# creating the var as dict
            
            for output_parameter in parameters[res_element]: # calling all the parameters within each element e.g. generator : p_mw, q__mvar, etc
                path = os.path.join(output_dir, res_element, output_parameter + '.xlsx')
                local_vars[network_element_str][output_parameter] = pd.read_excel(path, index_col=0) # reading the values, imported as pd.datafram 
                #parameter = pd.read_excel(path, index_col=0).T
                #parameter['parameter'] = output_parameter  
                #element = pd.concat([element, parameter])
                #paramter = pd.DataFrame()
            #element['element'] = network_element_str
            results.update({network_element_str :local_vars[network_element_str]}) # updating the dictionary that contain all the results
            #results_v02 = pd.concat([results_v02, element])
            #element = pd.DataFrame()
           
    return results #,results_v02
#%%
def run_time_series(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps):
 #%% 
    #. creating the temporary file path
    output_dir = os.path.join(tempfile.gettempdir(), "propens_pandapower_time_series")
    # print("Result can be found in your locan temp folder : {}".format(output_dir))
    
    #. number of columns, column lettere and name of the parameters to extracted from the results
    [number, column, parameters] = output_parameters(net, gen_fuel_tech, scenario_name) 
    
    #. the output writer with the desired result to be stored in the temporary files   
    create_output_writer(net, time_steps, output_dir = output_dir, parameters = parameters)  

    #. the main time series function
    run_timeseries(net)
  
    #. extract the results from the temporary file
    results = temp_files_to_excel_input(output_dir, parameters)

#%%   
    return results, net
#%%
"""
We create the output writer instead of saving the whole net that takes time.
We extract only pre defiend outputs
save the results to "../timeseries/tests/outputs"
write the results to ".xls" Excel files. (Possible are: .json, .p, .csv)
log the variables "p_mw" from "res_load", "vm_pu" from "res_bus" and two res_line values.
"""

def create_output_writer(net, time_steps, output_dir, parameters):
    ow = OutputWriter(net, time_steps, output_path = output_dir, output_file_type ='.xlsx', log_variables=list())
    # these variables are saved to the HDD after/during the time loop  
    for net_res_element in parameters:
        res = net_res_element[0:len('res')] # considering only the result parameters, not network topology (net)
        if res == 'res' :
            for output_parameter in parameters[net_res_element]: # calling all the parameters within each element e.g. generator : p_mw, q__mvar, etc
                ow.log_variable(net_res_element, output_parameter)
    return ow







    



