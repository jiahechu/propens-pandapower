# -*- coding: utf-8 -*-
"""
Created on Tue Dec 13 11:31:28 2022

@author: marti
"""



output_dir = os.path.join(tempfile.gettempdir(), "propens_pandapower_time_series")
print("Result can be found in your locan temp folder : {}".format(output_dir))
if not os.path.exists(output_dir):
    os.mkdir(output_dir)
run_time_series(output_dir, network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps)


# import Adv_network_only as addnet
# from excel_output_vTS import  output_parameters
# import os
# import tempfile
# from excel_output_vTS import create_excel

# network_name = 'topology'
# scenario_name = 'scenario'
# net = addnet.net 
# gen_fuel_tech = []
# [number, column, parameters] = output_parameters(net)
# output_path = []
# time_steps = 1
# output_dir = os.path.join(tempfile.gettempdir(), "time_series_example")


# create_excel(network_name, scenario_name, net, gen_fuel_tech, number, column, parameters, output_path, time_steps, output_dir)