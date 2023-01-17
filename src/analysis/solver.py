# -*- coding: utf-8 -*-
"""
Created on Tue Jan 10 22:33:32 2023

@author: marti
"""

import pandapower as pp
from src.analysis.time_series_func import run_time_series
from src.analysis.parameters import output_parameters
#%%
def solve(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps):
    print('\n Running Power Flow')   
    if time_steps > 1:
        results, net = run_time_series(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps)
    else: # for one iteration, 'results' is empty, as everything is in net.res_
        results, net = run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps)  
    print(' > Done')   
    return results, net

def run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps = 1):
    
    [number, column, parameters] = output_parameters(net, gen_fuel_tech, scenario_name)
    results = {}
    for element in number:
        results[element] = []  
    pp.runpp(net)
    return results,  net
