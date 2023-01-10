# -*- coding: utf-8 -*-
"""
Created on Tue Jan 10 22:33:32 2023

@author: marti
"""

import pandapower as pp
from src.analysis.excel_output import create_excel 
# import pandas as pd

from src.analysis.parameters import output_parameters

def run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps = 1):
    
    [number, column, parameters] = output_parameters(net, gen_fuel_tech)
    results = {}
    for element in number:
        results[element] = []    
    pp.runpp(net)
    create_excel(network_name, scenario_name, net, results, number, column, parameters, output_path, time_steps)
    return 0