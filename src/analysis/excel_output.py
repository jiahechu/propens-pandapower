# -*- coding: utf-8 -*-
"""
Created on Tue Nov 15 21:40:43 2022

@author: marti
"""

# import os
# import tempfile

# import pandapower.networks as pn
import pandapower as pp
import pandas as pd
#from openpyxl import Workbook
from openpyxl import load_workbook
from tqdm import tqdm
#from openpyxl.worksheet.table import Table, TableStyleInfo
#%%
# import Adv_network_only as addnet
# import Time_Series_Func as tsf
# import Analysis_Func as anal


# """
# For testing ADV_Network_Only compatibility, change in line 15~17 and 24
# """
# # cases to develop/test the code
# #net = pn.create_cigre_network_mv(with_der="all")
# #net = pn.case5()
# net = addnet.net
# #net = pn.panda_four_load_branch()
# pp.runpp(net)

# network_name = 'Network'
# scenario_name = 'Scenario'

# # input from fron-end
# gen_fuel_tech =[]

# output_dir = os.path.join(tempfile.gettempdir(), "time_series_example")
# Time_res = tsf.timeseries_example(output_dir)

# Sum_Bus_Vol_Under_Data = anal.Anal_Bus_Under(net.res_bus)
# Sum_Bus_Vol_Over_Data = anal.Anal_Bus_Over(net.res_bus)
# Sum_Trafo_Over_Data = anal.Anal_Trafo_Loading(net.res_trafo)
# Sum_Trafo3w_Over_Data = anal.Anal_Trafo3w_Loading(net.res_trafo3w)

def output_parameters(net):#, Sum_Bus_Vol_Under_Data):
    # Preallocate values: number of loads/gen/buses/trafos/lines, and columns in the excel template and their names
    
    # based on the network topology, the quentity of elements are counted
    number = { 'loads' : len(net.load),
               'generators' : len(net.gen),
               'lines' : len(net.line),
               'buses' : len(net.bus),
               'trafos' : len(net.trafo)}#,
               #'summary' : len(Sum_Bus_Vol_Under_Data) }
    
    # according to the desired excel output, the corresponding columns are assigning to each element
    column = { 'load' : ['B','C','D','E','F','G','H','I'],
              'gen' : ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q'],
              'line' : ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U'],
              'trafo' : ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V'],
              'bus' : ['B','C','D','E','F','G','H','I','J','K'],
              'summary' : ['B', 'C', 'D' ,'E', 'F'] }
    
    #  according to the desired excel output, the corresponding output_parameters by element are selected
    parameters = {'net_load' : ['zone','bus','vn_kv','in_service'],
                  'net_gen' : ['bus','name','in_service','vm_pu', 'max_p_mw', 'max_q_mvar','min_p_mw','min_q_mvar'],
                  'net_line' : ['name','from_bus','to_bus', 'in_service','length_km', 'max_i_ka','max_loading_percent', 'parallel', 'std_type'],
                  'net_trafo' : ['name', 'std_type', 'hv_bus', 'lv_bus', 'vn_hv_kv', 'vn_lv_kv', 'pfe_kw', 'shift_degree','tap_pos', 'parallel', 'in_service' ],
                  'net_bus' : ['zone','name','vn_kv','in_service'],
                  'res_load' : ['p_mw','q_mvar'],
                  'res_gen' : ['p_mw','q_mvar'],
                  'res_line' : ['p_from_mw', 'q_from_mvar', 'p_to_mw', 'q_to_mvar', 'pl_mw', 'ql_mvar', 'loading_percent'],
                  'res_trafo' : ['p_hv_mw','q_hv_mvar', 'p_lv_mw', 'q_lv_mvar', 'pl_mw', 'ql_mvar', 'loading_percent'],
                  'res_bus' : ['vm_pu','va_degree','p_mw','q_mvar'],
                  'summary' : ['component','Percentage','Extra Info'] }
    
    return number, column, parameters


   
def create_excel(network_name, scenario_name, net, results, gen_fuel_tech, number, column, parameters, output_path, time_steps):    
    # read the template, write the results and add important parameters from the network topology
    # finally save into a new excel workbook, according to the network topology and scenario name 
    #%%
    
    # cell in the template were all the tables start
    initial_cell = 'B3' # row 3: all the parameters names
    initial_line = 4 # row 4: where the fisrt line of parameter values are written
    
    
    # read the template and retrieve the sheets
    output_template = 'src/analysis/output_templates/output_template.xlsm'
    wb = load_workbook(filename = output_template, read_only = False, keep_vba = True)

    summary_sheet = wb["Summary"]
    demand_sheet = wb["Demand"]
    generators_sheet = wb["Generation"]
    buses_sheet = wb["Buses"]
    lines_sheet = wb["Lines"]
    trafos_sheet = wb["Trafos"]

    
    
    # %% Demand Sheet
    if number['loads'] > 0:           
        
        for i in range(len(parameters['net_load'])):
            if not parameters['net_load'][i] in net.load.keys():  # check the column of parameters
                net.load[parameters['net_load'][i]] = [None]*number['loads'] # if the columm of parameter is empty, it creat that columm and put value as none
        if time_steps == 1: # only check if there is no time series, as for time series the selection of output is done by the functions: temp_files_to_excel_input, create_output_writer
            for i in range(len(parameters['res_load'])): # check the column of parameters
                 if not parameters['res_load'][i] in net.res_load.keys():  # check the colum of parameter
                     net.res_load[parameters['res_load'][i]] = [None]*number['loads'] # if the columm of parameter is empty, it creat that columm and put value as none        
      
        # preallocating the columns names with the parameters from load, this is our desired ordered excel output
        lo_columns = ['step','zone','load_index','bus_index','in_service','load_voltage','p_mw','q_mvar']
        lo_table = {}
        for column_t in lo_columns:
            lo_table[column_t] = [0]*number['loads']
        load_table = pd.DataFrame(data = lo_table)
        

        # extract the values from the netowrk topology and power flow results
        i = 0
        for load_index in net.load.index:
            # from time step, here only the first iteration is donde, so '0'
            load_table.loc[i,'step'] = 0
            # from network topology
            bus_index = net.load.loc[load_index,'bus']
            bus_row = net.bus.index.to_list().index(bus_index) 
            load_table.loc[i,'zone'] = net.bus.iloc[bus_row]['zone']    
            load_table.loc[i,'load_index'] = load_index                
            load_table.loc[i,'bus_index'] = net.load.loc[load_index,'bus']     
            load_table.loc[i,'in_service'] = net.load.loc[load_index,'in_service']    
            load_table.loc[i,'load_voltage'] = net.bus.iloc[bus_row]['vn_kv']             
            #from power flow results: without time series, directly from net.res_##; with TS, from results (dict/dataframe)
            if time_steps == 1:
                load_table.loc[i,'p_mw'] = net.res_load.loc[load_index,'p_mw']  
                load_table.loc[i,'q_mvar'] = net.res_load.loc[load_index,'q_mvar']
            else:
                load_table.loc[i,'p_mw'] = results['load']['p_mw'].loc[0,load_index] 
                load_table.loc[i,'q_mvar'] = results['load']['q_mvar'].loc[0,load_index] 
            i = i + 1
             
        # results: if there is only '1' time step, directly from the pandapower network and the 'while' loop condiction won't start
        # otherwise the results were called from the temporary excels, and now they will be written from step 1 on
        # first step was 0 (zero)
        print('\n Writing the loads - 1/5')
        for step in tqdm(range(1,time_steps,1)):
            offset = step*number['loads'] # offset because all the values from all the loads and time steps are in the same table, 
            i = 0                                #so we write each time step results under the previous one
            for load_index in net.load.index:
                # time step
                load_table.loc[offset + i,'step'] = step
                # from network topology, is repeated for all the time steps
                bus_index = net.load.loc[i,'bus']
                bus_row = net.bus.index.to_list().index(bus_index) 
                load_table.loc[offset + i,'zone'] = net.bus.iloc[bus_row]['zone']    
                load_table.loc[offset + i,'load_index'] = load_index                 
                load_table.loc[offset + i,'bus_index'] = net.load.loc[load_index,'bus']     
                load_table.loc[offset + i,'in_service'] = net.load.loc[load_index,'in_service']    
                load_table.loc[offset + i,'load_voltage'] = net.bus.iloc[bus_row]['vn_kv'] 
                # from results, each value is taken from a different dataframe within results dict
                load_table.loc[offset + i,'p_mw'] = results['load']['p_mw'].loc[step,load_index]
                load_table.loc[offset + i,'q_mvar'] = results['load']['q_mvar'].loc[step,load_index]
                i = i + 1
            #we go to the next time step  
            step = step + 1
            
        # write the values in excel table
        for i in range(load_table.shape[0]): # going through all the rows
            for j in range(load_table.shape[1]): # going though all the columns
                load_row = initial_line + i #values start in the 'initial line' (row 4 of the excel sheet)               
                load_cell = column['load'][j] + str(load_row) #update cell reference, from left to right                
                if load_table.iloc[i,j] == None: # add --- if the value in the cell is None
                    load_value = '---'
                else:
                    load_value = load_table.iloc[i,j]  # add the value from the dataframe, auxiliar variable just to write the cell
                demand_sheet[load_cell] = load_value # update cell value
     
     
     # %% Generation's sheet
     
     #check if there is any generators, could be that there is an external grid only
     
    if number['generators'] > 0:  
     
        for i in range(len(parameters['net_gen'])):
             if not parameters['net_gen'][i] in net.gen.keys():
                 net.gen[parameters['net_gen'][i]] = [None]*number['generators']
        if time_steps == 1:
            for i in range(len(parameters['res_gen'])):
                 if not parameters['res_gen'][i] in net.res_gen.keys():
                     net.res_gen[parameters['res_gen'][i]] = [None]*number['generators']   
            
        if not len(gen_fuel_tech) > 0:
             d = {'fuel': ['---'], 'tech': ['---']}
             gen_fuel_tech = pd.DataFrame(data = d)
             df_aux = pd.DataFrame(data = d)
             for i in range(number['generators']):
                 gen_fuel_tech = pd.concat([gen_fuel_tech, df_aux],ignore_index=True)         
                 
        gen_columns = ['step','zone', 'bus_index', 'gen_index','fuel', 'tech','voltage','name','in_service','vm_pu',
                         'max_p_mw', 'max_q_mvar','min_p_mw','min_q_mvar','p_mw','q_mvar']
    
        ge_table = {}
        for column_t in gen_columns:
            ge_table[column_t] = [0]*number['generators']
        gen_table = pd.DataFrame(data = ge_table)
   
                     
        for i in range(number['generators']): # extract the values from power flow results
            gen_table.loc[i,'step'] = 0 
            
            bus_index = net.gen.loc[i,'bus']
            bus_row = net.bus.index.to_list().index(bus_index) 
            gen_table.loc[i,'zone'] = net.bus.iloc[bus_row]['zone']
            gen_table.loc[i,'bus_index'] = net.gen.loc[i,'bus']
            gen_table.loc[i,'gen_index'] = net.gen.index[i]
            
            gen_table.loc[i,'fuel'] = gen_fuel_tech.loc[i,'fuel']
            gen_table.loc[i,'tech'] = gen_fuel_tech.loc[i,'tech']
            
            gen_table.loc[i,'voltage'] = net.bus.iloc[bus_row]['vn_kv']
             
            gen_table.loc[i,'name'] = net.gen.loc[i,'name']
            gen_table.loc[i,'in_service'] =  net.gen.loc[i,'in_service']
            gen_table.loc[i,'vm_pu'] =  net.gen.loc[i,'vm_pu']
            gen_table.loc[i,'max_p_mw'] =  net.gen.loc[i,'max_p_mw']
            gen_table.loc[i,'max_q_mvar'] =  net.gen.loc[i,'max_q_mvar']
            gen_table.loc[i,'min_p_mw'] =  net.gen.loc[i,'min_p_mw']
            gen_table.loc[i,'min_q_mvar'] =  net.gen.loc[i,'min_q_mvar']
             
            if time_steps == 1:
                gen_table.loc[i,'p_mw'] = net.res_gen.loc[i,'p_mw']
                gen_table.loc[i,'q_mvar'] = net.res_gen.loc[i,'q_mvar']
            else: 
                gen_table.loc[i,'p_mw'] = results['gen']['p_mw'].loc[0,i] 
                gen_table.loc[i,'q_mvar'] = results['gen']['q_mvar'].loc[0,i]
                        
                 
        print('\n Writing the generators - 2/5')
        for step in tqdm(range(1,time_steps,1)):
            offset = step*number['generators']
                    
            for i in range(number['generators']): 
                gen_table.loc[offset + i,'step'] = step
                
                bus_index = net.gen.loc[i,'bus']
                bus_row = net.bus.index.to_list().index(bus_index) 
                gen_table.loc[offset + i,'zone' ] = net.bus.iloc[bus_row]['zone']
                gen_table.loc[offset + i,'bus_index'] = net.gen.loc[i,'bus']
                gen_table.loc[offset + i,'gen_index'] = net.gen.index[i]
                 
                gen_table.loc[offset + i,'fuel'] = gen_fuel_tech.loc[i,'fuel']
                gen_table.loc[offset + i,'tech'] = gen_fuel_tech.loc[i,'tech']
                 
                gen_table.loc[offset + i,'voltage'] = net.bus.iloc[bus_row]['vn_kv']
                 
                gen_table.loc[offset + i,'name'] = net.gen.loc[i,'name']
                gen_table.loc[offset + i,'in_service'] =  net.gen.loc[i,'in_service']
                gen_table.loc[offset + i,'vm_pu'] =  net.gen.loc[i,'vm_pu']
                gen_table.loc[offset + i,'max_p_mw'] =  net.gen.loc[i,'max_p_mw']
                gen_table.loc[offset + i,'max_q_mvar'] =  net.gen.loc[i,'max_q_mvar']
                gen_table.loc[offset + i,'min_p_mw'] =  net.gen.loc[i,'min_p_mw']
                gen_table.loc[offset + i,'min_q_mvar'] =  net.gen.loc[i,'min_q_mvar']
                 
                gen_table.loc[offset + i,'p_mw'] = results['gen']['p_mw'].loc[step,i] 
                gen_table.loc[offset + i,'q_mvar'] = results['gen']['q_mvar'].loc[step,i]
                    
            step = step + 1
                    
 
        for i in range(gen_table.shape[0]): # going to all the values in the same line/row = 
            for j in range(gen_table.shape[1]): # jump to next row 
                
                gen_row = initial_line + i #values start in the row 4 of the sheet
                gen_cell = column['gen'][j] + str(gen_row) #update cell reference, from left to right 
                
                if gen_table.iloc[i,j] == None:
                    gen_value = '---'
                else:
                    gen_value = gen_table.iloc[i,j] 
                
                generators_sheet[gen_cell] = gen_value # update cell value
    else:
        print('\n No generators in the net - 2/5')
 
    
     
     # %% Lines' sheet
     
    if number['lines'] > 0:
        
        for i in range(len(parameters['net_line'])):
             if not parameters['net_line'][i] in net.line.keys():
                 net.line[parameters['net_line'][i]] = [None]*number['lines']
        if time_steps == 1:
             for i in range(len(parameters['res_line'])):
                 if not parameters['res_line'][i] in net.res_line.keys():
                     net.res_line[parameters['res_line'][i]] = [None]*number['lines']        
                                   
        line_columns = ['step','zone','line_index','voltage','name','from_bus','to_bus','in_service', 
                        'length_km','max_i_ka','max_loading_percent','parallel','std_type','p_from_mw', 
                        'q_from_mvar','p_to_mw','q_to_mvar','pl_mw','ql_mvar','loading_percent']
        li_table = {}
        for column_t in line_columns:
            li_table[column_t] = [0]*number['lines']
        line_table = pd.DataFrame(data = li_table)
        
        
        i = 0                                
        for line_index in net.line.index: # extract the values from power flow results
            line_table.loc[i,'zone'] = 0
        
            bus_index = net.line.loc[line_index ,'from_bus']
            bus_row = net.bus.index.to_list().index(bus_index) 
            line_table.loc[i,'zone'] = net.bus.iloc[bus_row]['zone']
            line_table.loc[i,'line_index'] = line_index
            line_table.loc[i,'voltage'] = net.bus.iloc[bus_row]['vn_kv']
            line_table.loc[i,'name'] = net.line.loc[line_index ,'name']
            line_table.loc[i,'from_bus'] = net.line.loc[line_index ,'from_bus']
            line_table.loc[i,'to_bus'] = net.line.loc[line_index ,'to_bus']
            line_table.loc[i,'in_service'] = net.line.loc[line_index ,'in_service']
            line_table.loc[i,'length_km'] = net.line.loc[line_index ,'length_km']
            line_table.loc[i,'max_i_ka'] = net.line.loc[line_index ,'max_i_ka']
            line_table.loc[i,'max_loading_percent'] = net.line.loc[line_index ,'max_loading_percent']
            line_table.loc[i,'parallel'] = net.line.loc[line_index ,'parallel']
            line_table.loc[i,'std_type'] = net.line.loc[line_index ,'std_type']
             
            if time_steps == 1:
                line_table.loc[i,'p_from_mw'] = net.res_line.loc[line_index ,'p_from_mw']
                line_table.loc[i,'q_from_mvar'] = net.res_line.loc[line_index ,'q_from_mvar']
                line_table.loc[i,'p_to_mw'] = net.res_line.loc[line_index ,'p_to_mw']
                line_table.loc[i,'q_to_mvar'] = net.res_line.loc[line_index ,'q_to_mvar']
                line_table.loc[i,'pl_mw'] = net.res_line.loc[line_index ,'pl_mw']
                line_table.loc[i,'ql_mvar'] = net.res_line.loc[line_index ,'ql_mvar']
                line_table.loc[i,'loading_percent'] = net.res_line.loc[line_index ,'loading_percent']
            else: 
                line_table.loc[i,'p_from_mw'] = results['line']['p_from_mw'].loc[0,line_index ]
                line_table.loc[i,'q_from_mvar'] =  results['line']['q_from_mvar'].loc[0,line_index ]
                line_table.loc[i,'p_to_mw'] =  results['line']['p_to_mw'].loc[0,line_index ]
                line_table.loc[i,'q_to_mvar'] =  results['line']['q_to_mvar'].loc[0,line_index ]
                line_table.loc[i,'pl_mw'] =  results['line']['pl_mw'].loc[0,line_index ]
                line_table.loc[i,'ql_mvar'] =  results['line']['ql_mvar'].loc[0,line_index ]
                line_table.loc[i,'loading_percent'] =  results['line']['loading_percent'].loc[0,line_index ]
            i = i + 1    
            

        print('\n Writing the lines - 3/5')
        for step in tqdm(range(1,time_steps,1)):
            offset = step*number['lines']
            i = 0        
            for line_index in net.line.index: 
                line_table.loc[offset + i,'step'] = step
                
                bus_index = net.line.loc[line_index,'from_bus']
                bus_row = net.bus.index.to_list().index(bus_index) 
                line_table.loc[offset + i,'zone'] = net.bus.iloc[bus_row]['zone']
                line_table.loc[offset + i,'line_index'] = line_index
                line_table.loc[offset + i,'voltage'] = net.bus.iloc[bus_row]['vn_kv']
                line_table.loc[offset + i,'name'] = net.line.loc[line_index ,'name']
                line_table.loc[offset + i,'from_bus'] = net.line.loc[line_index ,'from_bus']
                line_table.loc[offset + i,'to_bus'] = net.line.loc[line_index ,'to_bus']
                line_table.loc[offset + i,'in_service'] = net.line.loc[line_index ,'in_service']
                line_table.loc[offset + i,'length_km'] = net.line.loc[line_index ,'length_km']
                line_table.loc[offset + i,'max_i_ka'] = net.line.loc[line_index ,'max_i_ka']
                line_table.loc[offset + i,'max_loading_percent'] = net.line.loc[line_index ,'max_loading_percent']
                line_table.loc[offset + i,'parallel'] = net.line.loc[line_index ,'parallel']
                line_table.loc[offset + i,'std_type'] = net.line.loc[line_index ,'std_type']
      
                line_table.loc[offset + i,'p_from_mw'] = results['line']['p_from_mw'].loc[step,line_index ]
                line_table.loc[offset + i,'q_from_mvar'] =  results['line']['q_from_mvar'].loc[step,line_index ]
                line_table.loc[offset + i,'p_to_mw'] =  results['line']['p_to_mw'].loc[step,line_index ]
                line_table.loc[offset + i,'q_to_mvar'] =  results['line']['q_to_mvar'].loc[step,line_index ]
                line_table.loc[offset + i,'pl_mw'] =  results['line']['pl_mw'].loc[step,line_index ]
                line_table.loc[offset + i,'ql_mvar'] =  results['line']['ql_mvar'].loc[step,line_index ]
                line_table.loc[offset + i,'loading_percent'] =  results['line']['loading_percent'].loc[step,line_index ]
                i = i + 1     
            step = step + 1
        
     
        for i in range(line_table.shape[0]): # going to all the values in the same line/row = 
            for j in range(line_table.shape[1]): # jump to next row 
                
                line_row = initial_line + i #values start in the row 4 of the sheet
                line_cell = column['line'][j] + str(line_row) #update cell reference, from left to right 
                
                if line_table.iloc[i,j] == None:
                    line_value = '---'
                else:
                    line_value = line_table.iloc[i,j] 
                
                lines_sheet[line_cell] = line_value # update cell value
     
        
     # %% Trafos' sheet
     
    if number['trafos'] > 0:
         
        for i in range(len(parameters['net_trafo'])):
             if not parameters['net_trafo'][i] in net.trafo.keys():
                 net.trafo[parameters['net_trafo'][i]] = [None]*number['trafos']
        if time_steps == 1:
            for i in range(len(parameters['res_trafo'])):
                 if not parameters['res_trafo'][i] in net.res_trafo.keys():
                     net.res_trafo[parameters['res_trafo'][i]] = [None]*number['trafos']        
                     
        trafo_columns = ['step','zone','trafo_index','name','std_type','hv_bus','lv_bus','vn_hv_kv','vn_lv_kv','pfe_kw',
                         'shift_degree','tap_pos','parallel','in_service','p_hv_mw','q_hv_mvar','p_lv_mw','q_lv_mvar',
                         'pl_mw','ql_mvar','loading_percent']
        
        tr_table = {}
        for column_t in trafo_columns:
            tr_table[column_t] = [0]*number['trafos']
        trafo_table = pd.DataFrame(data = tr_table)
                         
        i = 0 
        for trafo_index in net.trafo.index: # extract the values from power flow results
            trafo_table.loc[i,'step'] = 0 
            
            bus_index = net.trafo.loc[trafo_index,'hv_bus']
            bus_row = net.bus.index.to_list().index(bus_index) 
            trafo_table.loc[i,'zone'] = net.bus.iloc[bus_row]['zone']
            trafo_table.loc[i,'trafo_index'] = trafo_index
            trafo_table.loc[i,'name'] = net.trafo.loc[trafo_index,'name']
            trafo_table.loc[i,'std_type'] = net.trafo.loc[trafo_index,'std_type']
            trafo_table.loc[i,'hv_bus'] = net.trafo.loc[trafo_index,'hv_bus']
            trafo_table.loc[i,'lv_bus'] = net.trafo.loc[trafo_index,'lv_bus']
            trafo_table.loc[i,'vn_hv_kv'] = net.trafo.loc[trafo_index,'vn_hv_kv']
            trafo_table.loc[i,'vn_lv_kv'] = net.trafo.loc[trafo_index,'vn_lv_kv']
            trafo_table.loc[i,'pfe_kw'] = net.trafo.loc[trafo_index,'pfe_kw']
            trafo_table.loc[i,'shift_degree'] = net.trafo.loc[trafo_index,'shift_degree']
            trafo_table.loc[i,'tap_pos'] = net.trafo.loc[trafo_index,'tap_pos']
            trafo_table.loc[i,'parallel'] = net.trafo.loc[trafo_index,'parallel']
            trafo_table.loc[i,'in_service'] = net.trafo.loc[trafo_index,'in_service']
            
            if time_steps == 1:                
                trafo_table.loc[i,'p_hv_mw'] = net.res_trafo.loc[trafo_index,'p_hv_mw']
                trafo_table.loc[i,'q_hv_mvar'] = net.res_trafo.loc[trafo_index,'q_hv_mvar']
                trafo_table.loc[i,'p_lv_mw'] = net.res_trafo.loc[trafo_index,'p_lv_mw']
                trafo_table.loc[i,'q_lv_mvar'] = net.res_trafo.loc[trafo_index,'q_lv_mvar']
                trafo_table.loc[i,'pl_mw'] = net.res_trafo.loc[trafo_index,'pl_mw']
                trafo_table.loc[i,'ql_mvar'] =net.res_trafo.loc[trafo_index,'ql_mvar']
                trafo_table.loc[i,'loading_percent'] = net.res_trafo.loc[trafo_index,'loading_percent']
            else: 
                trafo_table.loc[i,'p_hv_mw'] = results['trafo']['p_hv_mw'].loc[0,trafo_index] 
                trafo_table.loc[i,'q_hv_mvar'] = results['trafo']['q_hv_mvar'].loc[0,trafo_index] 
                trafo_table.loc[i,'p_lv_mw'] = results['trafo']['p_lv_mw'].loc[0,trafo_index] 
                trafo_table.loc[i,'q_lv_mvar'] = results['trafo']['q_lv_mvar'].loc[0,trafo_index] 
                trafo_table.loc[i,'pl_mw'] = results['trafo']['pl_mw'].loc[0,trafo_index] 
                trafo_table.loc[i,'ql_mvar'] = results['trafo']['ql_mvar'].loc[0,trafo_index] 
                trafo_table.loc[i,'loading_percent'] = results['trafo']['loading_percent'].loc[0,trafo_index] 
            i = i + 1
                
        print('\n Writing the trafos - 4/5')
        for step in tqdm(range(1,time_steps,1)):
            offset = step*number['trafos']
            i = 0        
            for trafo_index in net.trafo.index: 
                trafo_table.loc[offset + i,'step'] = step
                
                bus_index = net.trafo.loc[trafo_index,'hv_bus']
                bus_row = net.bus.index.to_list().index(bus_index) 
                trafo_table.loc[offset + i,'zone'] = net.bus.iloc[bus_row]['zone']
                trafo_table.loc[offset + i,'trafo_index'] = trafo_index
                trafo_table.loc[offset + i,'name'] = net.trafo.loc[trafo_index,'name']
                trafo_table.loc[offset + i,'std_type'] = net.trafo.loc[trafo_index,'std_type']
                trafo_table.loc[offset + i,'hv_bus'] = net.trafo.loc[trafo_index,'hv_bus']
                trafo_table.loc[offset + i,'lv_bus'] = net.trafo.loc[trafo_index,'lv_bus']
                trafo_table.loc[offset + i,'vn_hv_kv'] = net.trafo.loc[trafo_index,'vn_hv_kv']
                trafo_table.loc[offset + i,'vn_lv_kv'] = net.trafo.loc[trafo_index,'vn_lv_kv']
                trafo_table.loc[offset + i,'pfe_kw'] = net.trafo.loc[trafo_index,'pfe_kw']
                trafo_table.loc[offset + i,'shift_degree'] = net.trafo.loc[trafo_index,'shift_degree']
                trafo_table.loc[offset + i,'tap_pos'] = net.trafo.loc[trafo_index,'tap_pos']
                trafo_table.loc[offset + i,'parallel'] = net.trafo.loc[trafo_index,'parallel']
                trafo_table.loc[offset + i,'in_service'] = net.trafo.loc[trafo_index,'in_service']
                   
                trafo_table.loc[offset + i,'p_hv_mw'] = results['trafo']['p_hv_mw'].loc[step,trafo_index] 
                trafo_table.loc[offset + i,'q_hv_mvar'] = results['trafo']['q_hv_mvar'].loc[step,trafo_index] 
                trafo_table.loc[offset + i,'p_lv_mw'] = results['trafo']['p_lv_mw'].loc[step,trafo_index] 
                trafo_table.loc[offset + i,'q_lv_mvar'] = results['trafo']['q_lv_mvar'].loc[step,trafo_index] 
                trafo_table.loc[offset + i,'pl_mw'] = results['trafo']['pl_mw'].loc[step,trafo_index] 
                trafo_table.loc[offset + i,'ql_mvar'] = results['trafo']['ql_mvar'].loc[step,trafo_index] 
                trafo_table.loc[offset + i,'loading_percent'] = results['trafo']['loading_percent'].loc[step,trafo_index] 
                i = i + 1
            step = step + 1
 
        for i in range(trafo_table.shape[0]): # going to all the values in the same line/row = 
            for j in range(trafo_table.shape[1]): # jump to next row 
                
                trafo_row = initial_line + i #values start in the row 4 of the sheet
                trafo_cell = column['trafo'][j] + str(trafo_row) #update cell reference, from left to right 
                
                if trafo_table.iloc[i,j] == None:
                    trafo_value = '---'
                else:
                    trafo_value = trafo_table.iloc[i,j] 
                
                trafos_sheet[trafo_cell] = trafo_value # update cell value
 

 
         
     # %% Buses Sheet
     
    if number['buses'] > 0:
         
        for i in range(len(parameters['net_bus'])):
             if not parameters['net_bus'][i] in net.bus.keys():
                 net.bus[parameters['net_bus'][i]] = [None]*number['buses']
        if time_steps ==1:
            for i in range(len(parameters['res_bus'])):
                if not parameters['res_bus'][i] in net.res_bus.keys():
                    net.res_bus[parameters['res_bus'][i]] = [None]*number['buses']   
                    
        bus_columns = ['step','index','zone','name','vn_kv','in_service','vm_pu','va_degree','p_mw','q_mvar']
        bu_table = {}
        for column_t in bus_columns:
            bu_table[column_t] = [0]*number['buses']
        bus_table = pd.DataFrame(data = bu_table )
        
        
        i = 0     
        for bus_index in net.bus.index:
            bus_table.loc[i,'step'] = 0
            
            bus_table.loc[i,'index'] = bus_index
            bus_table.loc[i,'zone'] = net.bus.loc[bus_index,'zone']
            bus_table.loc[i,'name'] = net.bus.loc[bus_index,'name']
            bus_table.loc[i,'vn_kv'] = net.bus.loc[bus_index,'vn_kv']
            bus_table.loc[i,'in_service'] = net.bus.loc[bus_index,'in_service']
            
            if time_steps == 1:
                bus_table.loc[i,'vm_pu'] = net.res_bus.loc[bus_index,'vm_pu']
                bus_table.loc[i,'va_degree'] = net.res_bus.loc[bus_index,'va_degree']
                bus_table.loc[i,'p_mw'] = net.res_bus.loc[bus_index,'p_mw']
                bus_table.loc[i,'q_mvar'] = net.res_bus.loc[bus_index,'q_mvar']
            else: 
                bus_table.loc[i,'vm_pu'] = results['bus']['vm_pu'].loc[0,bus_index]
                bus_table.loc[i,'va_degree'] = results['bus']['va_degree'].loc[0,bus_index]
                bus_table.loc[i,'p_mw'] = results['bus']['p_mw'].loc[0,bus_index]
                bus_table.loc[i,'q_mvar'] = results['bus']['q_mvar'].loc[0,bus_index]
            i = i + 1
        
        print('\n Writing the buses - 5/5')
        for step in tqdm(range(1,time_steps,1)):
            offset = step*number['buses']
            i = 0
            for bus_index in net.bus.index: 
                bus_table.loc[offset + i,'step'] = step
                
                bus_table.loc[offset + i,'index'] = bus_index
                bus_table.loc[offset + i,'zone'] = net.bus.loc[bus_index,'zone']
                bus_table.loc[offset + i,'name'] = net.bus.loc[bus_index,'name']
                bus_table.loc[offset + i,'vn_kv'] = net.bus.loc[bus_index,'vn_kv']
                bus_table.loc[offset + i,'in_service'] = net.bus.loc[bus_index,'in_service']
                
                bus_table.loc[offset + i,'vm_pu'] = results['bus']['vm_pu'].loc[step,bus_index]
                bus_table.loc[offset + i,'va_degree'] = results['bus']['va_degree'].loc[step,bus_index]
                bus_table.loc[offset + i,'p_mw'] = results['bus']['p_mw'].loc[step,bus_index]
                bus_table.loc[offset + i,'q_mvar'] = results['bus']['q_mvar'].loc[step,bus_index]      
                i = i + 1
            step = step + 1
     
        for i in range(bus_table.shape[0]): # going to all the values in the same line/row = 
            for j in range(bus_table.shape[1]): # jump to next row 
                
                bus_row = initial_line + i #values start in the row 4 of the sheet
                bus_cell = column['bus'][j] + str(bus_row) #update cell reference, from left to right 
                
                if bus_table.iloc[i,j] == None:
                    bus_value = '---'
                else:
                    bus_value = bus_table.iloc[i,j] 
                
                buses_sheet[bus_cell] = bus_value # update cell value
                

 
    
    # %% Summary Sheet
    
    
    
    #%%
    # """
    # if there's sth wrong, then Anal function is called.
    # Anal function is returning certain data base of faults
    # That data sheet will be written in here.
    
    
    # 30Nov Add
    # I may have to make the Summary_data based on Panda data frame format
    
    
    # 30Nov Add_2
    # When I add all data into one Summary_data, it becacme tuple. 
    # I need to make different Summary data for each component and add it into one set
    
    
    # 6Dec Add
    
    # Naming Convention
    
    # Sum_Bus_Vol_Under
    # Sum_Bus_Vol_Over
    # Sum_Line_Over
    # Sum_Trafo_Over
    # Sum_Trafo3w_Over
    
    
    
    # """

    # Sum_Bus_Vol_Under_Data = anal.Anal_Bus_Under(net.res_bus)
    
    # if Sum_Bus_Vol_Under_Data is not None :
    #     for i in range(len(parameters['summary'])):
    #         if not parameters['summary'][i] in Sum_Bus_Vol_Under_Data.keys():
    #             Sum_Bus_Vol_Under_Data[parameters['summary'][i]] = [None]*len(Sum_Bus_Vol_Under_Data)
                
    #     for i in range(len(Sum_Bus_Vol_Under_Data)):
    #         sum_index = Sum_Bus_Vol_Under_Data.index[i]
    #         sum_bus_vol_under_index = Sum_Bus_Vol_Under_Data.loc[i,['index']].to_list()
    #         sum_bus_vol_under_value = Sum_Bus_Vol_Under_Data.loc[i,['value']].to_list()
    #         sum_bus_vol_under_row = [step] + [sum_index] + sum_bus_vol_under_index + sum_bus_vol_under_value
            
    #         for j in range(len(sum_bus_vol_under_row)):
    #             sum_bus_vol_under_line = i + initial_line + 10 + step*len(Sum_Bus_Vol_Under_Data)
    #             sum_bus_vol_under_cell = column['summary'][j] + str(sum_bus_vol_under_line)
                
    #             if sum_bus_vol_under_row[j] == None:
    #                 sum_bus_vol_under_value_write = '---'
    #             else:
    #                 sum_bus_vol_under_value_write = sum_bus_vol_under_row[j]
                
    #             summary_sheet[sum_bus_vol_under_cell] = sum_bus_vol_under_value_write
                
    # ###############################################################
                
    # Sum_Bus_Vol_Over_Data = anal.Anal_Bus_Over(net.res_bus)
    
    # if Sum_Bus_Vol_Over_Data is not None :
    #     for i in range(len(parameters['summary'])):
    #         if not parameters['summary'][i] in Sum_Bus_Vol_Over_Data.keys():
    #             Sum_Bus_Vol_Over_Data[parameters['summary'][i]] = [None]*len(Sum_Bus_Vol_Over_Data)
                
    #     for i in range(len(Sum_Bus_Vol_Over_Data)):
    #         sum_index = Sum_Bus_Vol_Over_Data.index[i]
    #         sum_bus_vol_over_index = Sum_Bus_Vol_Over_Data.loc[i,['index']].to_list()
    #         sum_bus_vol_over_value = Sum_Bus_Vol_Over_Data.loc[i,['value']].to_list()
    #         sum_bus_vol_over_row = [step] + [sum_index] + sum_bus_vol_over_index + sum_bus_vol_over_value
            
    #         for j in range(len(sum_bus_vol_over_row)):
    #             sum_bus_vol_over_line = i + initial_line + sum_bus_vol_under_line + step*len(Sum_Bus_Vol_Over_Data)
    #             sum_bus_vol_over_cell = column['summary'][j] + str(sum_bus_vol_over_line)
                
    #             if sum_bus_vol_over_row[j] == None:
    #                 sum_bus_vol_over_value_write = '---'
    #             else:
    #                 sum_bus_vol_over_value_write = sum_bus_vol_over_row[j]
                
    #             summary_sheet[sum_bus_vol_over_cell] = sum_bus_vol_over_value_write
    
    
    # ###############################################################
    
    # Sum_Line_Over_Data = anal.Anal_Line_Loading_Better(net.res_line)
    
    # if Sum_Line_Over_Data is not None :
    #     for i in range(len(parameters['summary'])):
    #         if not parameters['summary'][i] in Sum_Line_Over_Data.keys():
    #             Sum_Line_Over_Data[parameters['summary'][i]] = [None]*len(Sum_Line_Over_Data)
                
    #     for i in range(len(Sum_Line_Over_Data)):
    #         sum_index = Sum_Line_Over_Data.index[i]
    #         sum_line_over_index = Sum_Line_Over_Data.loc[i,['index']].to_list()
    #         sum_line_over_value = Sum_Line_Over_Data.loc[i,['value']].to_list()
    #         sum_line_over_row = [step] + [sum_index] + sum_line_over_index + sum_line_over_value
            
    #         for j in range(len(sum_line_over_row)):
    #             sum_line_over_line = i + initial_line + sum_bus_vol_over_line + step*len(Sum_Line_Over_Data)
    #             sum_line_over_cell = column['summary'][j] + str(sum_line_over_line)
                
    #             if sum_line_over_row[j] == None:
    #                 sum_line_over_value_write = '---'
    #             else:
    #                 sum_line_over_value_write = sum_line_over_row[j]
                
    #             summary_sheet[sum_line_over_cell] = sum_line_over_value_write
                
    
    # ###############################################################
    # Sum_Trafo_Over_Data = anal.Anal_Trafo_Loading(net.res_trafo)
    
    # if Sum_Trafo_Over_Data is not None :
    #     for i in range(len(parameters['summary'])):
    #         if not parameters['summary'][i] in Sum_Trafo_Over_Data.keys():
    #             Sum_Trafo_Over_Data[parameters['summary'][i]] = [None]*len(Sum_Trafo_Over_Data)
                
    #     for i in range(len(Sum_Trafo_Over_Data)):
    #         sum_index = Sum_Trafo_Over_Data.index[i]
    #         sum_trafo_over_index = Sum_Trafo_Over_Data.loc[i,['index']].to_list()
    #         sum_trafo_over_value = Sum_Trafo_Over_Data.loc[i,['value']].to_list()
    #         sum_trafo_over_row = [step] + [sum_index] + sum_trafo_over_index + sum_trafo_over_value
            
    #         for j in range(len(sum_line_over_row)):
    #             sum_trafo_over_line = i + initial_line + sum_line_over_line + step*len(Sum_Trafo_Over_Data)
    #             sum_trafo_over_cell = column['summary'][j] + str(sum_trafo_over_line)
                
    #             if sum_trafo_over_row[j] == None:
    #                 sum_trafo_over_value_write = '---'
    #             else:
    #                 sum_trafo_over_value_write = sum_trafo_over_row[j]
                
    #             summary_sheet[sum_trafo_over_cell] = sum_trafo_over_value_write
    
    
    # ###############################################################
    # Sum_Trafo3w_Over_Data = anal.Anal_Trafo3w_Loading(net.res_trafo3w)
    
    # if Sum_Trafo3w_Over_Data is not None :
    #     for i in range(len(parameters['summary'])):
    #         if not parameters['summary'][i] in Sum_Trafo3w_Over_Data.keys():
    #             Sum_Trafo3w_Over_Data[parameters['summary'][i]] = [None]*len(Sum_Trafo3w_Over_Data)
                
    #     for i in range(len(Sum_Trafo3w_Over_Data)):
    #         sum_index = Sum_Trafo3w_Over_Data.index[i]
    #         sum_trafo3w_over_index = Sum_Trafo3w_Over_Data.loc[i,['index']].to_list()
    #         sum_trafo3w_over_value = Sum_Trafo3w_Over_Data.loc[i,['value']].to_list()
    #         sum_trafo3w_over_row = [step] + [sum_index] + sum_trafo3w_over_index + sum_trafo3w_over_value
            
    #         for j in range(len(sum_line_over_row)):
    #             sum_trafo3w_over_line = i + initial_line + sum_trafo_over_line + step*len(Sum_Trafo3w_Over_Data)
    #             sum_trafo3w_over_cell = column['summary'][j] + str(sum_trafo3w_over_line)
                
    #             if sum_trafo3w_over_row[j] == None:
    #                 sum_trafo3w_over_value_write = '---'
    #             else:
    #                 sum_trafo3w_over_value_write = sum_trafo3w_over_row[j]
                
    #             summary_sheet[sum_trafo3w_over_cell] = sum_trafo3w_over_value_write
    
    
    
    # %% Table reference, and here cell is the last cell added i.e. bottom-right corner of each table
    
    demand_table = demand_sheet.tables["demand_table"]
    generation_table = generators_sheet.tables["generation_table"]
    trafos_table = trafos_sheet.tables["trafos_table"]
    lines_table = lines_sheet.tables["lines_table"]
    bus_table = buses_sheet.tables["bus_table"]
            
    if number['loads'] > 0:
        demand_table.ref = initial_cell + ':' + load_cell
    if number['generators'] > 0:  
        generation_table.ref = initial_cell + ':' + gen_cell
    if number['trafos'] > 0:
        trafos_table.ref = initial_cell + ':' + trafo_cell
    if number['lines'] > 0:
        lines_table.ref = initial_cell + ':' + line_cell
    if number['buses'] > 0:
        bus_table.ref = initial_cell + ':' + bus_cell 
    
    # %% save with the topology and scenarios names
    
    file_name =  'results_' + network_name + '_' + scenario_name + '.xlsm'
    
    wb.save(output_path + '/' + file_name)
    print('\n Done - you can check the results with the path: ' + output_path ) 
    print('\n Done - the results were saved in : ' + file_name) 
    
    """
    Summary_data_Bus_Voltage = anal.Anal_Bus_Under(net.res_bus)
    
    if Summary_data_Bus_Voltage is not None :
            
            for i in range(len(parameters_summary)):
                if not parameters_summary[i] in Summary_data_Bus_Voltage.keys():
                    Summary_data_Bus_Voltage[parameters_summary[i]] = [None]*len(Summary_data_Bus_Voltage)
            
            for i in range(len(Summary_data_Bus_Voltage)):
                summary_index = Summary_data_Bus_Voltage.index[i]
                summary_bus_index = Summary_data_Bus_Voltage.loc[i,['index']].to_list()
                summary_bus_value = Summary_data_Bus_Voltage.loc[i, ['value']].to_list()
                Summary_bus_row = [step] + [summary_index] + summary_bus_index + summary_bus_value
                
                for j in range(len(Summary_bus_row)):
                    bus_sum_line = i + initial_line + 10 + step*len(Summary_data_Bus_Voltage) #due to button in Macro, I need to shift some a bit down
                    bus_sum_cell = summary_columm[j] + str(bus_sum_line)
                    
                    if Summary_bus_row[j] == None:
                        summary_value = '---'
                    else:
                        summary_value = Summary_bus_row[j]
                    
                    summary_sheet[bus_sum_cell] = summary_value # update cell value
                
    """

#%%
''' 
Running powerflow and create_exccel in case of ''1'' iteration
'''

def run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps = 1):
    
    [number, column, parameters] = output_parameters(net)
    results = []
    pp.runpp(net)
    create_excel(network_name, scenario_name, net, results, gen_fuel_tech, number, column, parameters, output_path, time_steps)
    
    return 0
    
    
    