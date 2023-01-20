# -*- coding: utf-8 -*-
"""
Created on Sat Jan 14 11:53:24 2023

@author: marti
"""
from src.analysis.sort_results import sort_results 
from src.analysis.parameters import check_parameter
from src.analysis.parameters import output_parameters
import pandas as pd

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