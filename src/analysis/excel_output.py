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
from src.analysis.sort_results import sort_load_results 
from src.analysis.sort_results import sort_gen_results
from src.analysis.sort_results import sort_line_results
from src.analysis.sort_results import sort_trafo_results
from src.analysis.sort_results import sort_bus_results


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
    
    # %% Demand Sheet
    if number['loads'] > 0:  # if there is no load, jump to the next element            
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
        load_table = pd.DataFrame(data = lo_table) # preallocating the dataframe 
        load_table = sort_load_results(net, number, load_table, time_steps,results['load']) # read the values from results, and sort them for the output
        # write the values in excel table
        print('Now writing into the excel file')
        for i in tqdm(range(load_table.shape[0])): # going through all the rows
            for j in range(load_table.shape[1]): # going though all the columns
                load_row = initial_line + i #values start in the 'initial line' (row 4 of the excel sheet)               
                load_cell = column['load'][j] + str(load_row) #update cell reference, from left to right                
                if load_table.iloc[i,j] == None: # add --- if the value in the cell is None
                    load_value = '---'
                else:
                    load_value = load_table.iloc[i,j]  # add the value from the dataframe, auxiliar variable just to write the cell
                wb["Demand"][load_cell] = load_value # update cell value
        del results['load']
        del load_table
     
     # %% Generation's sheet     
    if number['generators'] > 0:#check if there is any generators, could be that there is an external grid only
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
        gen_table = sort_gen_results(net, number, gen_table, time_steps,results['gen'],gen_fuel_tech)                  
        print('Now writing into the excel file')
        for i in tqdm(range(gen_table.shape[0])): # going to all the values in the same line/row = 
            for j in range(gen_table.shape[1]): # jump to next row 
                gen_row = initial_line + i #values start in the row 4 of the sheet
                gen_cell = column['gen'][j] + str(gen_row) #update cell reference, from left to right 
                if gen_table.iloc[i,j] == None:
                    gen_value = '---'
                else:
                    gen_value = gen_table.iloc[i,j] 
                wb["Generation"][gen_cell] = gen_value # update cell value
        del results['gen']
        del gen_table
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
        line_table = sort_line_results(net, number, line_table, time_steps,results['line'])
        print('Now writing into the excel file')
        for i in tqdm(range(line_table.shape[0])): # going to all the values in the same line/row = 
            for j in range(line_table.shape[1]): # jump to next row 
                line_row = initial_line + i #values start in the row 4 of the sheet
                line_cell = column['line'][j] + str(line_row) #update cell reference, from left to right 
                if line_table.iloc[i,j] == None:
                    line_value = '---'
                else:
                    line_value = line_table.iloc[i,j] 
                wb["Lines"][line_cell] = line_value # update cell value
        del results['line']
        del line_table
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
        trafo_table = sort_trafo_results(net, number, trafo_table, time_steps,results['trafo'])              
        print('Now writing into the excel file')
        for i in tqdm(range(trafo_table.shape[0])): # going to all the values in the same line/row = 
            for j in range(trafo_table.shape[1]): # jump to next row               
                trafo_row = initial_line + i #values start in the row 4 of the sheet
                trafo_cell = column['trafo'][j] + str(trafo_row) #update cell reference, from left to right                 
                if trafo_table.iloc[i,j] == None:
                    trafo_value = '---'
                else:
                    trafo_value = trafo_table.iloc[i,j]          
                wb["Trafos"] [trafo_cell] = trafo_value # update cell value
        del results['trafo']
        del trafo_table
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
        bus_table = sort_bus_results(net, number, bus_table, time_steps,results['bus'])
        print('Now writing into the excel file')
        for i in tqdm(range(bus_table.shape[0])): # going to all the values in the same line/row = 
            for j in range(bus_table.shape[1]): # jump to next row          
                bus_row = initial_line + i #values start in the row 4 of the sheet
                bus_cell = column['bus'][j] + str(bus_row) #update cell reference, from left to right 
                if bus_table.iloc[i,j] == None:
                    bus_value = '---'
                else:
                    bus_value = bus_table.iloc[i,j]           
                wb["Buses"][bus_cell] = bus_value # update cell value
        del results['bus']
        del bus_table
    # %% Summary Sheet
    
    summary_sheet = wb["Summary"]

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
    
    
    
    # %% Update Table reference, and here the cell is the last cell added i.e. bottom-right corner of each table  
    print('\n Updating tables references in excel...')
    if number['loads'] > 0:
        wb["Demand"].tables["demand_table"].ref = initial_cell + ':' + load_cell
    if number['generators'] > 0:  
        wb['Generation'].tables["generation_table"].ref = initial_cell + ':' + gen_cell
    if number['trafos'] > 0:
        wb["Trafos"].tables["trafos_table"].ref = initial_cell + ':' + trafo_cell
    if number['lines'] > 0:
        wb["Lines"].tables["lines_table"].ref = initial_cell + ':' + line_cell
    if number['buses'] > 0:
        wb["Buses"].tables["bus_table"].ref = initial_cell + ':' + bus_cell 
    
    # %% save with the topology and scenarios names
    print('\n Closing workbook ...')
    file_name =  'results_' + network_name + '_' + scenario_name + '.xlsm'
    
    wb.save(output_path + '/' + file_name)
    print('\n Done' )
    print('      > You can check the results with the path: ' + output_path ) 
    print('      > and the results were saved in : ' + file_name) 
    
    #%%
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
    
    
    