# -*- coding: utf-8 -*-
"""
Created on Sat Jan 14 11:53:24 2023

@author: marti
"""
from src.analysis.sort_results import sort_results 
from src.analysis.parameters import check_parameter
from src.analysis.parameters import output_parameters

def save_results(net, gen_fuel_tech, scenario_name, time_steps, results):
    print('\n Sorting the results')
    tables = {}
    number, column, parameters = output_parameters(net, gen_fuel_tech, scenario_name)
    for element in number:
        if number[element] > 0:  # if there is no load, jump to the next element            
            # checking values of the parameters, and adding columns and/or formatting them    
            net = check_parameter(net, time_steps, parameters, number, element) 
            # read the values from results, and sort them for the output
            tables[element] = sort_results(net, number, time_steps, results[element], column, parameters, element) 
            # delete the results to free memory
            del results[element]
        else:
            print(' > no '+ element +'s in the net')
            
    return tables