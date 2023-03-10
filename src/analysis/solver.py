# -*- coding: utf-8 -*-
"""
Created on Tue Jan 10 22:33:32 2023

@author: marti
"""

import pandapower as pp
from src.analysis.time_series_func import run_time_series
from src.analysis.parameters import output_parameters
from src.analysis.plot import plot_interactive

#%%
def solve(network_name, scenario_name, gen_fuel_tech,output_setup , net, time_steps, general):
    output_path = output_setup['output_path']
    for parameter in general:
        if parameter == 'ts_path':
            continue
        else:
            if general[parameter][0]:
                continue
            elif not general[parameter][0]:
                continue
            else: 
                print('\n ------ Error in General configuration, parameters should be either True or False---------')
                raise
    
    if general['use_ts'][0] == False:# for one iteration, 'results' is empty, as everything is in net.res_
        results, net = run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, general, output_setup)  
    elif general['use_opf'][0] == False:
        results, net = run_time_series(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps)
        
    else:
        print('\n ------ Error in General configuration, might be because you are trying OPF with Time Series Analysis, in this version this features can be used at the same time ---------')
        raise
    
    print('\n >>>>>  Solve Network - Done')   
    return results, net

def run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, general, output_setup):
    
    [number, column, parameters] = output_parameters(net, gen_fuel_tech, scenario_name)
    results = {}
    for element in number:
        results[element] = []  
    
    if general['use_opf'][0] == True:
        if general['use_dc'][0] == True:
            print('\n Running a Linearized Optimal Power Flow ')
            pp.rundcopp(net)
        else:
            print('\n Running a Non-linear Optimal Power Flow ')
            pp.runopp(net)
    else: # run power flow
        if general['use_dc'][0] == True:
            print('\n Running a Linearized Power Flow ')
            pp.rundcpp(net)
        else:
            print('\n Running a Non-linear Power Flow ')
            pp.runpp(net)
    
    plot_interactive(output_setup, net, network_name, scenario_name)

    
    return results,  net
