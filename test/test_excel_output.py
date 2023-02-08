# -*- coding: utf-8 -*-
"""
Test excel_output
"""
import numpy as np
import pandapower as pp
import pandas as pd
import unittest
import src.analysis.excel_output as eo
import src.analysis.save as save
import src.analysis.parameters as par
import os
# import parameters as par

def load_tables(): 
    net = pp.create_empty_network(name = 'test')
    time_steps = 1
    scenario_name = 'basic'
    gen_fuel_tech = []
    [number, column, parameters] = par.output_parameters(net, gen_fuel_tech, scenario_name)
    results = {}
    for element in number:
        results[element] = []  
    tables = save.save_results(net, gen_fuel_tech, scenario_name, time_steps, results)
    
    load = {  'step' : [0, 0,	0,	0,	0,	0,	0,	0],
        'index':[0,	1,	2,	3,	4,	5,	6,	7],
        'scenario' :['basic', 'basic', 'basic', 'basic', 'basic', 'basic', 'basic', 'basic'],
        'zone'	: [np.nan]*8,
        'name':	['households_#']*8,
        'bus' : 	[2,	3,	4,	5,	6,	7,	8,	9],
        'in_service' :	[True]*8,
        'p_mw' :	[0.0072]*8,
        'q_mvar': [0]*8 }
    load_df = pd.DataFrame(data = load)
    tables['load'] = load_df
    return tables

class TestCreate_excel(unittest.TestCase):   
    # parameters = par.select_parameters()
    # sheet, cell, elements_by_type = par.sheets_parameters()
    
    def test_create_excel(self):
        topology_name = 'test'
        folder_name = 'test_data'
        output_path = './'+ folder_name
        tables = {}
        tables['basic'] = load_tables()
        if not os.path.isdir(folder_name): os.makedirs(folder_name)
        flag = eo.create_excel(topology_name, output_path, tables)
        tables['basic']['load'].to_excel(output_path+'/'+'test_create_excel_data.xlsx')  
        self.assertEqual(flag, True)

if __name__ == '__main__':
    # begin the unittest.main()
    unittest.main()

