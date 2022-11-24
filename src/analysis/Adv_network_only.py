# -*- coding: utf-8 -*-
"""
Created on Sat Oct 29 11:56:15 2022

@author: thoug
"""

import pandapower as pp
import pandas as pd
from pandapower.plotting import simple_plot as plot
from pandapower.plotting.plotly import simple_plotly
from pandapower.networks import mv_oberrhein
from pandapower.plotting.plotly import vlevel_plotly
#%matplotlib inline


#create an empty network
net = pp.create_empty_network()

# Double busbar
pp.create_bus(net, name='Double Busbar 1', vn_kv=380, type='b')
pp.create_bus(net, vn_kv=380, name='Double Busbar 2', type='b')
for i in range(10):
    pp.create_bus(net, name='Bus DB T%s' % i, vn_kv=380, type='n')
for i in range(1,5):
    pp.create_bus(net, name='Bus DB %s' %i, vn_kv=380, type='n')
    
#Single busbar
pp.create_bus(net, name='Single Busbar', vn_kv=110, type='b')
for i in range(1,6):
    pp.create_bus(net, name='Bus SB %s' %i, vn_kv=110, type='n')
for i in range(1,6):
    for j in [1,2]:
        pp.create_bus(net, name = 'Bus SB T%s.%s' %(i,j), vn_kv=110, type='n')
        
#Remaining buses
for i in range(1,5):
    pp.create_bus(net, name='Bus HV%s'%i, vn_kv=110, type='n')
    
#show bustable
net.bus




#create hv lines
hv_lines=pd.read_csv("csv\hv_lines.csv", sep=';', header =0, decimal=',')
hv_lines

#create all lines
for _, hv_line in hv_lines.iterrows():
    from_bus = pp.get_element_index(net, 'bus', hv_line.from_bus)
    to_bus = pp.get_element_index(net, 'bus', hv_line.to_bus)
    pp.create_line(net, from_bus, to_bus, length_km=hv_line.length, std_type=hv_line.std_type, name=hv_line.line_name, parallel=hv_line.parallel)
#show line table
net.line

#create Transformer
hv_bus = pp.get_element_index(net, 'bus', 'Bus DB 2')
lv_bus = pp.get_element_index(net, 'bus', 'Bus SB 1')
pp.create_transformer_from_parameters(net, hv_bus, lv_bus, sn_mva=300, vn_hv_kv=380, vn_lv_kv=110, vkr_percent=0.06, vk_percent=8, pfe_kw=0, i0_percent=0, tp_pos=0, shift_degree=0, name='EHV-HV-Trafo')
#show transformer info
net.trafo





#create switches
hv_bus_sw = pd.read_csv('csv/hv_bus_sw.csv', sep=';', header=0, decimal=',')
hv_bus_sw

#Bus-bus switches
for _, switch in hv_bus_sw.iterrows():
    from_bus = pp.get_element_index(net, 'bus', switch.from_bus)
    to_bus = pp.get_element_index(net, 'bus', switch.to_bus)
    pp.create_switch(net, from_bus, to_bus, et=switch.et, closed=switch.closed, type=switch.type, name=switch.bus_name)
    
    
# Bus-line switches
hv_buses = net.bus[(net.bus.vn_kv == 380) | (net.bus.vn_kv == 110)].index
hv_ls = net.line[(net.line.from_bus.isin(hv_buses)) & (net.line.to_bus.isin(hv_buses))]
for _, line in hv_ls.iterrows():
        pp.create_switch(net, line.from_bus, line.name, et='l', closed=True, type='LBS', name='Switch %s - %s' % (net.bus.name.at[line.from_bus], line['name']))
        pp.create_switch(net, line.to_bus, line.name, et='l', closed=True, type='LBS', name='Switch %s - %s' % (net.bus.name.at[line.to_bus], line['name']))

# Trafo-line switches
pp.create_switch(net, pp.get_element_index(net, "bus", 'Bus DB 2'), pp.get_element_index(net, "trafo", 'EHV-HV-Trafo'), et='t', closed=True, type='LBS', name='Switch DB2 - EHV-HV-Trafo')
pp.create_switch(net, pp.get_element_index(net, "bus", 'Bus SB 1'), pp.get_element_index(net, "trafo", 'EHV-HV-Trafo'), et='t', closed=True, type='LBS', name='Switch SB1 - EHV-HV-Trafo')

# show switch table
net.switch




#create external grid
pp.create_ext_grid(net, pp.get_element_index(net, 'bus', 'Double Busbar 1'), vm_pu=1.03, va_degree=0, name='External grid', s_sc_max_mva=10000, rx_max=0.1, rx_min=0.1)

net.ext_grid




#create loads

#HV_Loads
hv_loads = pd.read_csv('csv/hv_loads.csv', sep=';', header =0, decimal=',')
hv_loads

for _, load in hv_loads.iterrows():
    bus_idx = pp.get_element_index(net, 'bus', load.bus)
    pp.create_load(net, bus_idx, p_mw=load.p, q_mvar=load.q, name=load.load_name)
    
net.load




#create Generator
pp.create_gen(net, pp.get_element_index(net, 'bus', 'Bus HV4'), vm_pu=1.03, p_mw=100, name='Gas turbine')

net.gen

#create Static Generator
pp.create_sgen(net, pp.get_element_index(net, 'bus', 'Bus SB 5'), p_mw=20, q_mvar=4, sn_mva=45, type='WP', name='Wind Park')

net.sgen




#create Shunt
pp.create_shunt(net, pp.get_element_index(net, 'bus', 'Bus HV1'), p_mw=0, q_mvar=0.960, name='Shunt')

net.shunt




#create Impedance(extended ward equivalents)
pp.create_impedance(net, pp.get_element_index(net, 'bus', 'Bus HV3'), pp.get_element_index(net, 'bus', 'Bus HV1'), rft_pu=0.074873, xft_pu=0.198872, sn_mva=100, name='Impedance')

net.impedance

#XWards
pp.create_xward(net, pp.get_element_index(net, "bus", 'Bus HV3'), ps_mw=23.942, qs_mvar=-12.24187, pz_mw=2.814571, 
                qz_mvar=0, r_ohm=0, x_ohm=12.18951, vm_pu=1.02616, name='XWard 1')
pp.create_xward(net, pp.get_element_index(net, "bus", 'Bus HV1'), ps_mw=3.776, qs_mvar=-7.769979, pz_mw=9.174917, 
                qz_mvar=0, r_ohm=0, x_ohm=50.56217, vm_pu=1.024001, name='XWard 2')

# show xward table
net.xward

#simple_plot(net)



#HV_LEVEL done-------------------------------------------------------





#Medium Voltage Level
pp.create_bus(net, name='Bus MV0 20kV', vn_kv=20, type='n')
for i in range(8):
    pp.create_bus(net, name='Bus MV%s' % i, vn_kv=10, type='n')

#show only medium voltage bus table
mv_buses = net.bus[(net.bus.vn_kv == 10) | (net.bus.vn_kv == 20)]
mv_buses


#Create MV Lines
mv_lines = pd.read_csv('csv/mv_lines.csv', sep=';', header=0, decimal=',')
for _, mv_line in mv_lines.iterrows():
    from_bus = pp.get_element_index(net, "bus", mv_line.from_bus)
    to_bus = pp.get_element_index(net, "bus", mv_line.to_bus)
    pp.create_line(net, from_bus, to_bus, length_km=mv_line.length, std_type=mv_line.std_type, name=mv_line.line_name)

# show only medium voltage lines
net.line[net.line.from_bus.isin(mv_buses.index)]



#3-Ph Transformer
hv_bus = pp.get_element_index(net, "bus", "Bus HV2")
mv_bus = pp.get_element_index(net, "bus", "Bus MV0 20kV")
lv_bus = pp.get_element_index(net, "bus", "Bus MV0")
pp.create_transformer3w_from_parameters(net, hv_bus, mv_bus, lv_bus, vn_hv_kv=110, vn_mv_kv=20, vn_lv_kv=10, 
                                        sn_hv_mva=40, sn_mv_mva=15, sn_lv_mva=25, vk_hv_percent=10.1, 
                                        vk_mv_percent=10.1, vk_lv_percent=10.1, vkr_hv_percent=0.266667, 
                                        vkr_mv_percent=0.033333, vkr_lv_percent=0.04, pfe_kw=0, i0_percent=0, 
                                        shift_mv_degree=30, shift_lv_degree=30, tap_side="hv", tap_neutral=0, tap_min=-8, 
                                        tap_max=8, tap_step_percent=1.25, tap_pos=0, name='HV-MV-MV-Trafo')

# show transformer3w table
net.trafo3w



#MV Switches
# Bus-line switches
mv_buses = net.bus[(net.bus.vn_kv == 10) | (net.bus.vn_kv == 20)].index
mv_ls = net.line[(net.line.from_bus.isin(mv_buses)) & (net.line.to_bus.isin(mv_buses))]
for _, line in mv_ls.iterrows():
        pp.create_switch(net, line.from_bus, line.name, et='l', closed=True, type='LBS', name='Switch %s - %s' % (net.bus.name.at[line.from_bus], line['name']))
        pp.create_switch(net, line.to_bus, line.name, et='l', closed=True, type='LBS', name='Switch %s - %s' % (net.bus.name.at[line.to_bus], line['name']))

# open switch
open_switch_id = net.switch[(net.switch.name == 'Switch Bus MV5 - MV Line5')].index
net.switch.closed.loc[open_switch_id] = False

#show only medium voltage switch table
net.switch[net.switch.bus.isin(mv_buses)]


#MV Loads
mv_loads = pd.read_csv('csv/mv_loads.csv', sep=';', header=0, decimal=',')
for _, load in mv_loads.iterrows():
    bus_idx = pp.get_element_index(net, "bus", load.bus)
    pp.create_load(net, bus_idx, p_mw=load.p, q_mvar=load.q, name=load.load_name)

# show only medium voltage loads
net.load[net.load.bus.isin(mv_buses)]




#MV Static Gens
mv_sgens = pd.read_csv('csv/mv_sgens.csv', sep=';', header=0, decimal=',')
for _, sgen in mv_sgens.iterrows():
    bus_idx = pp.get_element_index(net, "bus", sgen.bus)
    pp.create_sgen(net, bus_idx, p_mw=sgen.p, q_mvar=sgen.q, sn_mva=sgen.sn, type=sgen.type, name=sgen.sgen_name)

# show only medium voltage static generators
net.sgen[net.sgen.bus.isin(mv_buses)]








#MV_LEVEL-------------------------------------------------------



#LV Level

#LV Bus
pp.create_bus(net, name='Bus LV0', vn_kv=0.4, type='n')
for i in range(1, 6):
    pp.create_bus(net, name='Bus LV1.%s' % i, vn_kv=0.4, type='m')
for i in range(1, 5):
    pp.create_bus(net, name='Bus LV2.%s' % i, vn_kv=0.4, type='m')
pp.create_bus(net, name='Bus LV2.2.1', vn_kv=0.4, type='m')
pp.create_bus(net, name='Bus LV2.2.2', vn_kv=0.4, type='m')

# show only low voltage buses
lv_buses = net.bus[net.bus.vn_kv == 0.4]
lv_buses



# create LV lines
lv_lines = pd.read_csv('csv/lv_lines.csv', sep=';', header=0, decimal=',')
for _, lv_line in lv_lines.iterrows():
    from_bus = pp.get_element_index(net, "bus", lv_line.from_bus)
    to_bus = pp.get_element_index(net, "bus", lv_line.to_bus)
    pp.create_line(net, from_bus, to_bus, length_km=lv_line.length, std_type=lv_line.std_type, name=lv_line.line_name)

# show only low voltage lines
net.line[net.line.from_bus.isin(lv_buses.index)]


#LV Transformer
hv_bus = pp.get_element_index(net, "bus", "Bus MV4")
lv_bus = pp.get_element_index(net, "bus","Bus LV0")
pp.create_transformer_from_parameters(net, hv_bus, lv_bus, sn_mva=.4, vn_hv_kv=10, vn_lv_kv=0.4, vkr_percent=1.325, vk_percent=4, pfe_kw=0.95, i0_percent=0.2375, tap_side="hv", tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tp_pos=0, shift_degree=150, name='MV-LV-Trafo')

#show only low voltage transformer
net.trafo[net.trafo.lv_bus.isin(lv_buses.index)]



#LV Switches
lv_buses
# Bus-line switches
lv_ls = net.line[(net.line.from_bus.isin(lv_buses.index)) & (net.line.to_bus.isin(lv_buses.index))]
for _, line in lv_ls.iterrows():
        pp.create_switch(net, line.from_bus, line.name, et='l', closed=True, type='LBS', name='Switch %s - %s' % (net.bus.name.at[line.from_bus], line['name']))
        pp.create_switch(net, line.to_bus, line.name, et='l', closed=True, type='LBS', name='Switch %s - %s' % (net.bus.name.at[line.to_bus], line['name']))

# Trafo-line switches
pp.create_switch(net, pp.get_element_index(net, "bus", 'Bus MV4'), pp.get_element_index(net, "trafo", 'MV-LV-Trafo'), et='t', closed=True, type='LBS', name='Switch MV4 - MV-LV-Trafo')
pp.create_switch(net, pp.get_element_index(net, "bus", 'Bus LV0'), pp.get_element_index(net, "trafo", 'MV-LV-Trafo'), et='t', closed=True, type='LBS', name='Switch LV0 - MV-LV-Trafo')

# show only low vvoltage switches
net.switch[net.switch.bus.isin(lv_buses.index)]



#LV Loads
lv_loads = pd.read_csv('csv/lv_loads.csv', sep=';', header=0, decimal=',')
for _, load in lv_loads.iterrows():
    bus_idx = pp.get_element_index(net, "bus", load.bus)
    pp.create_load(net, bus_idx, p_mw=load.p, q_mvar=load.q, name=load.load_name)
    
# show only low voltage loads
net.load[net.load.bus.isin(lv_buses.index)]




#LV Static Gens
lv_sgens = pd.read_csv('csv/lv_sgens.csv', sep=';', header=0, decimal=',')
for _, sgen in lv_sgens.iterrows():
    bus_idx = pp.get_element_index(net, "bus", sgen.bus)
    pp.create_sgen(net, bus_idx, p_mw=sgen.p, q_mvar=sgen.q, sn_mva=sgen.sn, type=sgen.type, name=sgen.sgen_name)

# show only low voltage static generators
net.sgen[net.sgen.bus.isin(lv_buses.index)]


#LV_LEVEL-------------------------------------------------------


#------------------------------------
#Finish creaiting nets
#-------------------------------------------

"""
#run powerflow
pp.runpp(net, calculate_voltage_angles=True, init="dc")




#simple_plotly(net);
#vlevel_plotly(net);


#Plot Graph
plot(net)

#Command to export result to Excel
#pp.to_excel(net, 'excel_read_out.xlsx', include_results=True)

######################################################################

import Analysis_function as af

af.Anal_Bus_Voltage(net.res_bus)
af.Anal_Line_Loading(net.res_line)
af.Anal_Trafo_Loading(net.res_trafo)
af.Anal_Trafo3w_Loading(net.res_trafo3w)
af.Anal_Line_Loading_Better(net.res_line)

#final goal is to call one function to call all sub_functions?
#or just call all the other sub_functions individually
#af.Anal()

#####################################################3
#####################################################3
#####################################################3

import Time_Series_Func as TS
TS.timeseries_example()
"""