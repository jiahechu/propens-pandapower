# -*- coding: utf-8 -*-
"""

"""

from pandapower.plotting import simple_plot 
from pandapower.plotting import simple_plotly
from pandapower.plotting import pf_res_plotly

def plot_topology(output_setup, net):
    if output_setup['plot']['topology'] == True: 
        if net['bus_geodata'].shape[0] != net['bus'].shape[0]: 
            print('\n Missmatch in the number of coordinates of the buses')
            print(' Plotting will be abborted')
        else: 
            simple_plot(net)            
    return

def plot_interactive(output_setup, net):
    if output_setup['plot']['interactive network'] == True or output_setup['plot']['interactive heat map network'] == True:
        if net['bus_geodata'].shape[0] != net['bus'].shape[0]:
            print('\n Missmatch in the number of coordinates of the buses')
            print(' Plotting will be abborted')
        else:
            try:
                if output_setup['plot']['interactive network'] == True: simple_plotly(net)
                if output_setup['plot']['interactive heat map network'] == True: pf_res_plotly(net)
            except: 
                print('Error while doing the interactive plotting')
                print('Please change the interactive plotting to False to continue')
                print('Program stops.')
                raise               
    return