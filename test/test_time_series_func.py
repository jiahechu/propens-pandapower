"""
Test parameters.py.
"""
import unittest
import src.analysis.time_series_func as tsf

import numpy as np
import pandas as pd

import pandapower.control as control
import pandapower.networks as nw
from pandapower.timeseries.run_time_series import run_timeseries
from pandapower.timeseries.data_sources.frame_data import DFData



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

class TestTime_series_func(unittest.TestCase):
    parameters = {'net_bus' : ['scenario','zone','name','vn_kv','in_service'],
                  'net_load' : ['scenario','zone','name','bus','in_service'],
                  'net_gen' : ['scenario','zone','name','type','bus','in_service','fuel','vm_pu', 'max_p_mw', 'max_q_mvar','min_p_mw','min_q_mvar'],
                  'net_line' : ['scenario','zone','name','from_bus','to_bus', 'in_service','length_km', 'max_i_ka','max_loading_percent', 'parallel', 'std_type'],
                  'net_trafo' : ['scenario','zone','name', 'std_type', 'hv_bus', 'lv_bus', 'vn_hv_kv', 'vn_lv_kv', 'pfe_kw', 'shift_degree','tap_pos', 'parallel', 'in_service' ],
                  'res_bus' : ['vm_pu'],           
                  'res_line' : ['loading_percent']} 
    output_dir = './test_data'  
    def test_create_output_writer(self):
        net = create_ts_net()
        # network_name = 'pandapower_time_series'
        # scenario_name = 'test_create_output_writer'
        time_steps = 10
        # gen_fuel_tech=[]
             
        output_dir = self.output_dir
        #. number of columns, column lettere and name of the parameters to extracted from the results

        #. the output writer with the desired result to be stored in the temporary files 
        # print('\n Selecting the output variables')  
        ow = tsf.create_output_writer(net, time_steps, output_dir = output_dir, parameters = self.parameters)
        
        self.assertIn('vm_pu', ow.log_variables[0][1])
        self.assertIn('loading_percent',ow.log_variables[1][1])
        
    def test_run_time_series(self):
        net = create_ts_net()
        # network_name = 'pandapower_time_series'
        # scenario_name = 'test_create_output_writer'
        time_steps = 10
        # gen_fuel_tech=[]
        output_dir = self.output_dir        
        #. the output writer with the desired result to be stored in the temporary files 
        # print('\n Selecting the output variables')  
        tsf.create_output_writer(net, time_steps, output_dir = output_dir, parameters = self.parameters)
        #. the main time series function
        # print('\n Running time series')
        run_timeseries(net)
        #. extract the results from the temporary file
        # print('\n Reading the results from the temporary files')
        results = tsf.temp_files_to_excel_input(output_dir, self.parameters)
        self.assertIn('bus',results)
        self.assertIn('vm_pu',results['bus'])
        self.assertIn('line',results)
        self.assertIn('loading_percent',results['line'])
    
if __name__ == '__main__':
    # begin the unittest.main()
    unittest.main()
