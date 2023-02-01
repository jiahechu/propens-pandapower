"""
Test parameters.py.
"""
import unittest
import src.analysis.time_series_func as tsf

import numpy as np
import pandas as pd

import pandapower.control as control
import pandapower as pp
import pandapower.networks as nw
from pandapower.timeseries.run_time_series import run_timeseries
from pandapower.timeseries.data_sources.frame_data import DFData
import src.analysis.parameters as par
import src.analysis.save as save



def create_ts_net():
    # load a pandapower network
    net = nw.mv_oberrhein(scenario='generation')
    # number of time steps
    n_ts = 2
    # load your timeseries from a file (here csv file)
    # df = pd.read_csv("sgen_timeseries.csv")
    # or create a DataFrame with some random time series as an example
    df = pd.DataFrame(np.random.normal(1., 0.1, size=(n_ts, len(net.sgen.index))),
                  index=list(range(n_ts)), columns=net.sgen.index) * net.sgen.p_mw.values
    # create the data source from it
    ds = DFData(df)
    
    # initialising ConstControl controller to update values of the regenerative generators ("sgen" elements)
    # the element_index specifies which elements to update (here all sgens in the net since net.sgen.index is passed)
    # the controlled variable is "p_mw"
    # the profile_name are the columns in the csv file (here this is also equal to the sgen indices 0-N )
    control.ConstControl(net, element='sgen', element_index=net.sgen.index,
                                  variable='p_mw', data_source=ds, profile_name=net.sgen.index)
    
    # do the same for loads
    # df = pd.read_csv("load_timeseries.csv")
    # create a DataFrame with some random time series as an example
    df = pd.DataFrame(np.random.normal(1., 0.1, size=(n_ts, len(net.load.index))),
                  index=list(range(n_ts)), columns=net.load.index) * net.load.p_mw.values
    ds = DFData(df)
    control.ConstControl(net, element='load', element_index=net.load.index,
                                  variable='p_mw', data_source=ds, profile_name=net.load.index)
    return net

class TestSave(unittest.TestCase):
    
    output_dir = './test_ts_data' 
    def test_ts_sort_results(self):
        net = create_ts_net()
        scenario_name = 'test_ts_sort_results'
        time_steps = 2
        gen_fuel_tech=[]
        output_dir = self.output_dir 
        number, column, parameters = par.output_parameters(net, gen_fuel_tech, scenario_name)        
        #. the output writer with the desired result to be stored in the temporary files  
        tsf.create_output_writer(net, time_steps, output_dir = output_dir, parameters = parameters)
        #. the main time series function
        run_timeseries(net)
        #. extract the results from the temporary file
        results = tsf.temp_files_to_excel_input(output_dir, parameters)
        # formatting the results into the excel output table form
        table = save.sort_results(net, number, time_steps, results['bus'], column, parameters, 'bus', scenario_name)
        #checking the new table columns,
        self.assertEqual(column['parameter']['bus'], list(table.keys()))
             
    def test_one_iteration_sort_results(self):
        net = nw.case5()
        scenario_name = 'test_one_iteration_sort_results'
        time_steps = 1
        gen_fuel_tech=[]
        number, column, parameters = par.output_parameters(net, gen_fuel_tech, scenario_name)        
        #. run pfn
        pp.runpp(net)
        results = []
        # formatting the results into the excel output table form
        table = save.sort_results(net, number, time_steps, results, column, parameters, 'bus', scenario_name)
        #checking the new table columns,
        self.assertEqual(column['parameter']['bus'], list(table.keys()))
         
    def test_save_results(self):
        net = create_ts_net()
        scenario_name = 'test_save_results'
        time_steps = 2
        gen_fuel_tech=[]
        output_dir = self.output_dir  
        number, column, parameters = par.output_parameters(net, gen_fuel_tech, scenario_name)        
        #. the output writer with the desired result to be stored in the temporary files  
        tsf.create_output_writer(net, time_steps, output_dir = output_dir, parameters = parameters)
        #. the main time series function
        run_timeseries(net)
        #. extract the results from the temporary file
        results = tsf.temp_files_to_excel_input(output_dir, parameters)
        # formatting and saving the results into the excel output table form
        tables = save.save_results(net, gen_fuel_tech, scenario_name, time_steps, results)
        #checking the new tables keys
        self.assertEqual(number.keys(),tables.keys())
        
if __name__ == '__main__':
    # begin the unittest.main()
    unittest.main()
