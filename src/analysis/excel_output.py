# -*- coding: utf-8 -*-
"""
Created on Tue Nov 15 21:40:43 2022

@author: marti
"""

# import os
# import tempfile
# import pandapower.networks as pn
import pandapower as pp
import numpy as np
# import pandas as pd
from openpyxl import load_workbook
from tqdm import tqdm

from src.analysis.sort_results import sort_results 

from src.analysis.parameters import check_parameter

#%%

def write_in_the_excel(table, sheet, element, column, initial_line, number):
    print('Now writing into the excel file')
    
    # j=0
    # for column_title in table.keys():
    #     row = initial_line
    #     cell = column['letter'][element][j] + str(row)
    #     sheet[cell] = column_title
    #     j = j + 1
    if element == 'sgen': initial_line = initial_line + number['gen']
        
    for i in tqdm(range(table.shape[0])): # going through all the rows
        for j in range(table.shape[1]): # going though all the columns
            row = initial_line + i #values start in the 'initial line' (row 5 of the excel sheet)               
            cell = column['letter'][element][j] + str(row) #update cell reference, from left to right                
            if table.iloc[i,j] == None or str(table.iloc[i,j]) == 'nan': # add --- if the value in the cell is None
                value = '---'
            else:
                value = table.iloc[i,j]  # add the value from the dataframe, auxiliar variable just to write the cell
            sheet[cell] = value # update cell value
    return cell

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
   
def create_excel(network_name, scenario_name, net, results, number, column, parameters, output_path, time_steps):    
    # read the template, write the results and add important parameters from the network topology
    # finally save into a new excel workbook, according to the network topology and scenario name 
    #%%  
    # cell in the template were all the tables start
    initial_cell = 'C4' # row 3: all the parameters names
    initial_line = 5 # row 4: where the fisrt line of parameter values are written
    # read the template and retrieve the sheets
    output_template = 'src/analysis/output_templates/output_template.xlsm'
    wb = load_workbook(filename = output_template, read_only = False, keep_vba = True)   
    cell = {} # preallocate dictionary with last cell of tables, for references
    
    sheet = {'load': 'Demand',
             'bus': 'Buses',
             'gen':'Generation',
             'sgen':'Generation',
             'line': 'Lines',
             'trafo': 'Trafos'}
    
    table_name = sheet # in the excel template the tables have the same as the sheet 
    # %% Demand Sheet
    for element in number:
        if number[element] > 0:  # if there is no load, jump to the next element            
            # checking values of the parameters, and adding columns and/or formatting them    
            net = check_parameter(net, time_steps, parameters, number, element) 
            # read the values from results, and sort them for the output
            table = sort_results(net, number, time_steps, results[element], column, parameters, element) 
            # write the values in excel table
            cell[element] = write_in_the_excel(table, wb[sheet[element]], element, column, initial_line, number)
            # delete the results to free memory
            del results[element]
        else:
            print('\n No '+ element +'s in the net')
     
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
    print('\n\n Updating tables references in excel...')
    # the order of the element in the 'number' dict  is important, as it dictate which table in on top when
    # they are written in the same sheet, e.g. generation includes gen & sgen 
    for element in number:
        if number[element] > 0:
            wb[sheet[element]].tables[table_name[element]].ref = initial_cell + ':' + cell[element]
    
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

    
    
    