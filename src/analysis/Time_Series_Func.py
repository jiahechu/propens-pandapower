# -*- coding: utf-8 -*-
"""
Created on Wed Nov 23 09:04:03 2022

@author: thoug
"""
#https://github.com/e2nIEE/pandapower/blob/develop/tutorials/time_series.ipynb

import os
import numpy as np
import pandas as pd
import tempfile

import pandapower as pp
from pandapower.timeseries import DFData # Dataframe datasoure that holds time series to be calculated
from pandapower.timeseries import OutputWriter #To write the outputs to HDD
from pandapower.timeseries.run_time_series import run_timeseries #main time series function, calls the controller function to update P,Q value and run pp
from pandapower.control import ConstControl # constant controllers, change P and Q values of sgens and loads


import Adv_network_only as addnet

"""
1. Create a simple test net
2. Create the data source, which contains the time series P value
3. create the controllers to update the P values of the load and the sgen
4. define the output writer and desired variables to be saved
5. call the main time series function to calculate the desired results
"""

time_steps = 9000

def timeseries_example(output_dir):
    #1. Create a simple test net
    net = addnet.net  #going to be defined below
    #Since Adavanced_NEtwork already have the net and for later, we get the net from front-end
    
    
    #2. Create (random) data source
    n_timesteps = 9000
    profiles, ds = create_data_source(n_timesteps)
    
    #3. Create controller
    create_controllers(net, ds)
    
    #4. the output writer with the desired result to be sotred to files
    ow = create_output_writer(net, time_steps, output_dir = output_dir)
    
    #5. the main time series function
    run_timeseries(net, time_steps)
    
    a = pd.DataFrame(run_timeseries(net, time_steps))
    
    #6 Return value in case of need
    return a, net
    
    """
    
    since we will get pre-defined network from front-end, this is not needed
    
def simple_test_net():

    
    net = pp.create_empty_network()
    pp.set_user_pf_options(net, init_vm_pu = 'flat', init_va_degree='dc', calculate_voltage_angles = True)

    b0 = pp.create_bus(net, 110)
    b1 = pp.create_bus(net, 110)
    b2 = pp.create_bus(net, 20)
    b3 = pp.create_bus(net, 20)
    b4 = pp.create_bus(net, 20)
    
    pp.create_ext_grid(net, b0)
    pp.create_line(net, b0, b1, 10, "149-AL1/24-ST1A 110.0", name='0-1')
    pp.create_transformer(net, b1, b2, "25 MVA 110/20 kV", name='tr1')
    pp.create_line(net, b2, b3, 10, "184-AL1/30-ST1A 20.0", name='2-3')
    pp.create_line(net, b2, b4, 10, "184-AL1/30-ST1A 20.0", name='2-4')
    
    pp.create_load(net, b3, p_mw=20., q_mavar=10., name='load1')
    pp.create_sgen(net, b4, p_mw=20., q_mavar=0.15, name = 'sgen1'j)
    
    return net
"""
    
def create_data_source(n_timesteps=9000):
    profiles = pd.DataFrame() #making datapase frame
    profiles['load1_p'] = np.random.random(n_timesteps)*20
    profiles['load2_p'] = np.random.random(n_timesteps)*20
    profiles['load3_p'] = np.random.random(n_timesteps)*20
    profiles['load4_p'] = np.random.random(n_timesteps)*20
    profiles['load5_p'] = np.random.random(n_timesteps)*20
    profiles['load6_p'] = np.random.random(n_timesteps)*20
    profiles['load7_p'] = np.random.random(n_timesteps)*20
    profiles['load8_p'] = np.random.random(n_timesteps)*20
    profiles['load9_p'] = np.random.random(n_timesteps)*20
    profiles['load10_p'] = np.random.random(n_timesteps)*20
    profiles['load11_p'] = np.random.random(n_timesteps)*20
    profiles['load12_p'] = np.random.random(n_timesteps)*20
    profiles['load13_p'] = np.random.random(n_timesteps)*20
    profiles['load14_p'] = np.random.random(n_timesteps)*20
    profiles['load15_p'] = np.random.random(n_timesteps)*20
    profiles['load16_p'] = np.random.random(n_timesteps)*20
    profiles['load17_p'] = np.random.random(n_timesteps)*20
    profiles['load18_p'] = np.random.random(n_timesteps)*20
    profiles['load19_p'] = np.random.random(n_timesteps)*20
    profiles['load20_p'] = np.random.random(n_timesteps)*20
    profiles['load21_p'] = np.random.random(n_timesteps)*20
    profiles['load22_p'] = np.random.random(n_timesteps)*20
    profiles['load23_p'] = np.random.random(n_timesteps)*20
    profiles['load24_p'] = np.random.random(n_timesteps)*20
    
    profiles['sgen1_p'] = np.random.random(n_timesteps)*20
    profiles['sgen2_p'] = np.random.random(n_timesteps)*20
    profiles['sgen3_p'] = np.random.random(n_timesteps)*20
    profiles['sgen4_p'] = np.random.random(n_timesteps)*20
    profiles['sgen5_p'] = np.random.random(n_timesteps)*20
    profiles['sgen6_p'] = np.random.random(n_timesteps)*20
    profiles['sgen7_p'] = np.random.random(n_timesteps)*20
    profiles['sgen8_p'] = np.random.random(n_timesteps)*20
    profiles['sgen9_p'] = np.random.random(n_timesteps)*20
    profiles['sgen10_p'] = np.random.random(n_timesteps)*20
   

                                # add data in database frame
    ds = DFData(profiles)
    
    return profiles, ds  #return the modified database frame


def create_controllers(net, ds):
    ConstControl(net, element='load', variable='p_mw', element_index=[1], data_source = ds, profile_name = ["load1_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[3], data_source = ds, profile_name = ["load2_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[4], data_source = ds, profile_name = ["load3_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[6], data_source = ds, profile_name = ["load4_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[1], data_source = ds, profile_name = ["load5_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[3], data_source = ds, profile_name = ["load6_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[4], data_source = ds, profile_name = ["load7_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[6], data_source = ds, profile_name = ["load8_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[1], data_source = ds, profile_name = ["load9_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[3], data_source = ds, profile_name = ["load10_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[4], data_source = ds, profile_name = ["load11_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[6], data_source = ds, profile_name = ["load12_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[1], data_source = ds, profile_name = ["load13_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[3], data_source = ds, profile_name = ["load14_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[4], data_source = ds, profile_name = ["load15_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[6], data_source = ds, profile_name = ["load16_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[1], data_source = ds, profile_name = ["load17_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[3], data_source = ds, profile_name = ["load18_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[4], data_source = ds, profile_name = ["load19_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[6], data_source = ds, profile_name = ["load20_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[1], data_source = ds, profile_name = ["load21_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[3], data_source = ds, profile_name = ["load22_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[4], data_source = ds, profile_name = ["load23_p"])
    ConstControl(net, element='load', variable='p_mw', element_index=[6], data_source = ds, profile_name = ["load24_p"])
    
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen1_p"])
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen2_p"])
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen3_p"])
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen4_p"])
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen5_p"])
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen6_p"])
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen7_p"])
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen8_p"])
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen9_p"])
    ConstControl(net, element='sgen', variable='p_mw', element_index=[0], data_source = ds, profile_name = ["sgen10_p"])



"""
We create the output writer instead of saving the whole net that takes time.
We extract only pre defiend outputs
save the results to "../timeseries/tests/outputs"
write the results to ".xls" Excel files. (Possible are: .json, .p, .csv)
log the variables "p_mw" from "res_load", "vm_pu" from "res_bus" and two res_line values.
"""

def create_output_writer(net, time_steps, output_dir):
    ow = OutputWriter(net, time_steps, output_path = output_dir, output_file_type ='.xlsx', log_variables=list())
    # these variables are saved to the HDD after/during the time loop
    ow.log_variable('res_load', 'p_mw')
    ow.log_variable('res_bus', 'vm_pu')
    ow.log_variable('res_line', 'loading_percent')
    ow.log_variable('res_line', 'i_ka')
    ow.log_variable('res_trafo', 'loading_percent')
    return ow


output_dir = os.path.join(tempfile.gettempdir(), "time_series_example")
print("Result can be found in your locan temp folder : {}".format(output_dir))
if not os.path.exists(output_dir):
    os.mkdir(output_dir)
timeseries_example(output_dir)




#%%
#plot the result
"""
Basically, from create_output_writer, it export the data
and we read that file here again, to plot the data
"""
import matplotlib.pyplot as plt
#%matplotlib inline  

# voltage results
vm_pu_file = os.path.join(output_dir, "res_bus", "vm_pu.xlsx")
vm_pu = pd.read_excel(vm_pu_file, index_col=0)
vm_pu.plot(label="vm_pu")
plt.xlabel("time step")
plt.ylabel("voltage mag. [p.u.]")
plt.title("Voltage Magnitude")
plt.grid()
plt.show()

# line loading results
ll_file = os.path.join(output_dir, "res_line", "loading_percent.xlsx")
line_loading = pd.read_excel(ll_file, index_col=0)
line_loading.plot(label="line_loading")
plt.xlabel("time step")
plt.ylabel("line loading [%]")
plt.title("Line Loading")
plt.grid()
plt.show()

# load results
load_file = os.path.join(output_dir, "res_load", "p_mw.xlsx")
load = pd.read_excel(load_file, index_col=0)
load.plot(label="load")
plt.xlabel("time step")
plt.ylabel("P [MW]")
plt.grid()
plt.show()

# trafo loading
trafo_file = os.path.join(output_dir, "res_trafo", "loading_percent.xlsx")
trafo_loading = pd.read_excel(trafo_file, index_col=0)
trafo_loading.plot(label="trafo")
plt.xlabel("time step")
plt.ylabel("Transformer Loading [%]")
plt.grid()
plt.show()

    
    



