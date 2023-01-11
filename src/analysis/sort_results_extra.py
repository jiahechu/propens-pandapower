# -*- coding: utf-8 -*-
"""
Created on Tue Jan 10 21:46:15 2023

@author: marti
"""

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

# import os
# import tempfile
# import pandapower.networks as pn
import pandapower as pp
import numpy as np
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






















    # %% Generation's sheet     
   if number['gen'] > 0:#check if there is any generators, could be that there is an external grid only
       # checking values of the parameters, and adding columns and/or formatting them    
       net = check_parameter(net, time_steps, parameters, number, 'gen') 
       # read the values from results, and sort them for the output
       gen_table = sort_results(net, number, time_steps, results['gen'], column, parameters)                  
       # write the values in excel table
       cell['gen'] = write_in_the_excel(gen_table, wb["Generation"], 'gen', column, initial_line, cell)
       # delete the results and table to free memory
       del results['gen']
       del gen_table
   else:
       print('\n No generators in the net - 2/5')
     
    # %% Lines' sheet   
   if number['line'] > 0:     
       # checking values of the parameters, and adding columns and/or formatting them  
       net = check_parameter(net, time_steps, parameters, number, 'line')                                  
       # read the values from results, and sort them for the output
       line_table = sort_results(net, number, time_steps,results['line'], column, parameters,'line')
       # write the values in excel table
       cell['line'] = write_in_the_excel(line_table, wb['Lines'], 'line', column, initial_line, cell)
       # delete the results and table to free memory
       del results['line']
       del line_table
       
    # %% Trafos' sheet     
   if number['trafo'] > 0:
       # checking values of the parameters, and adding columns and/or formatting them  
       net = check_parameter(net, time_steps, parameters, number, 'trafo') 
       # read the values from results, and sort them for the output
       trafo_table = sort_results(net, number, time_steps,results['trafo'], column, parameters,'trafo')
       # write the values in excel table   
       cell['trafo'] = write_in_the_excel(trafo_table, wb['Trafos'], 'trafo', column, initial_line,cell)
       # delete the results and table to free memory        
       del results['trafo']
       del trafo_table
    # %% Buses Sheet    
   if number['bus'] > 0:
       # checking values of the parameters, and adding columns and/or formatting them  
       net = check_parameter(net, time_steps, parameters, number, 'bus')  
       # read the values from results, and sort them for the output
       bus_table = sort_results(net, number, time_steps, results['bus'], column, parameters, 'bus')
       # write the values in excel table
       cell['bus'] = write_in_the_excel(bus_table, wb['Buses'], 'bus', column, initial_line, cell)
       # delete the results and table to free memory
       del results['bus']
       del bus_table

def sort_gen_results(net, number, time_steps, gen, gen_fuel_tech, column, parameters): 
    gen_table = preallocate_table('gen', column, number)
    element = 'gen'
    i = 0
    for i in range(number['gen']): # extract the values from power flow results
        gen_table.loc[i,'step'] = 0 
        
        gen_table.loc[i,'index'] = net.gen.index[i]
        gen_table.loc[i,parameters['net_' + element]] = net.gen.loc[i,parameters['net_' + element]]
         
        if time_steps == 1:
            gen_table.loc[i,'p_mw'] = net.res_gen.loc[i,'p_mw']
            gen_table.loc[i,'q_mvar'] = net.res_gen.loc[i,'q_mvar']
        else: 
            gen_table.loc[i,'p_mw'] = gen['p_mw'].loc[0,i] 
            gen_table.loc[i,'q_mvar'] = gen['q_mvar'].loc[0,i]
                    
             
    print('\n Sorting the results: generators - 2/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['gen']
                
        for i in range(number['gen']): 
            gen_table.loc[offset + i,'step'] = step
            gen_table.loc[offset + i,'index'] = net.gen.index[i]
             
            gen_table.loc[offset + i, parameters['net_' + element]] = net.gen.loc[i,parameters['net_' + element]]
             
            gen_table.loc[offset + i,'p_mw'] = gen['p_mw'].loc[step,i] 
            gen_table.loc[offset + i,'q_mvar'] = gen['q_mvar'].loc[step,i]
                
        step = step + 1

    return gen_table


def sort_line_results(net, number, time_steps, line, column, parameters): 
    line_table = preallocate_table('line', column, number)
    element = 'line'
    i = 0                                
    for line_index in net.line.index: # extract the values from power flow results
        line_table.loc[i,'zone'] = 0

        line_table.loc[i,'index'] = line_index
        line_table.loc[i,parameters['net_' + element]] = net.line.loc[line_index, parameters['net_' + element]]
         
        if time_steps == 1:
            line_table.loc[i, parameters['res_' + element]] = net.res_line.loc[line_index ,parameters['res_' + element]]

        else: 
            line_table.loc[i,'p_from_mw'] = line['p_from_mw'].loc[0,line_index ]
            line_table.loc[i,'q_from_mvar'] =  line['q_from_mvar'].loc[0,line_index ]
            line_table.loc[i,'p_to_mw'] =  line['p_to_mw'].loc[0,line_index ]
            line_table.loc[i,'q_to_mvar'] =  line['q_to_mvar'].loc[0,line_index ]
            line_table.loc[i,'pl_mw'] =  line['pl_mw'].loc[0,line_index ]
            line_table.loc[i,'ql_mvar'] =  line['ql_mvar'].loc[0,line_index ]
            line_table.loc[i,'loading_percent'] =  line['loading_percent'].loc[0,line_index ]
        i = i + 1    

    print('\n Sorting the results: lines - 3/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['line']
        i = 0        
        for line_index in net.line.index: 
            line_table.loc[offset + i,'step'] = step
            
            line_table.loc[offset + i,'index'] = line_index
            line_table.loc[offset + i, parameters['net_' + element]] = net.line.loc[line_index ,parameters['net_' + element]]
            
            line_table.loc[offset + i,'p_from_mw'] = line['p_from_mw'].loc[step,line_index ]
            line_table.loc[offset + i,'q_from_mvar'] =  line['q_from_mvar'].loc[step,line_index ]
            line_table.loc[offset + i,'p_to_mw'] =  line['p_to_mw'].loc[step,line_index ]
            line_table.loc[offset + i,'q_to_mvar'] =  line['q_to_mvar'].loc[step,line_index ]
            line_table.loc[offset + i,'pl_mw'] =  line['pl_mw'].loc[step,line_index ]
            line_table.loc[offset + i,'ql_mvar'] =  line['ql_mvar'].loc[step,line_index ]
            line_table.loc[offset + i,'loading_percent'] =  line['loading_percent'].loc[step,line_index ]
            i = i + 1     
        step = step + 1

    return line_table 


def sort_trafo_results(net, number, time_steps, trafo, column, parameters): 
    trafo_table = preallocate_table('trafo', column, number)        
    element = 'trafo'    
    i = 0 
    for trafo_index in net.trafo.index: # extract the values from power flow results
        trafo_table.loc[i,'step'] = 0 
        
        trafo_table.loc[i,'index'] = trafo_index
        trafo_table.loc[i,parameters['net_' + element]] = net.trafo.loc[trafo_index, parameters['net_' + element]]
        
        if time_steps == 1:                
            trafo_table.loc[i,parameters['res_' + element]] = net.res_trafo.loc[trafo_index, parameters['res_' + element]]
        else: 
            trafo_table.loc[i,'p_hv_mw'] = trafo['p_hv_mw'].loc[0,trafo_index] 
            trafo_table.loc[i,'q_hv_mvar'] = trafo['q_hv_mvar'].loc[0,trafo_index] 
            trafo_table.loc[i,'p_lv_mw'] = trafo['p_lv_mw'].loc[0,trafo_index] 
            trafo_table.loc[i,'q_lv_mvar'] = trafo['q_lv_mvar'].loc[0,trafo_index] 
            trafo_table.loc[i,'pl_mw'] = trafo['pl_mw'].loc[0,trafo_index] 
            trafo_table.loc[i,'ql_mvar'] = trafo['ql_mvar'].loc[0,trafo_index] 
            trafo_table.loc[i,'loading_percent'] = trafo['loading_percent'].loc[0,trafo_index] 
        i = i + 1
            
    print('\n Sorting the results: trafos - 4/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['trafo']
        i = 0        
        for trafo_index in net.trafo.index: 
            trafo_table.loc[offset + i,'step'] = step
            
            trafo_table.loc[offset + i,'index'] = trafo_index
            trafo_table.loc[offset + i,parameters['net_' + element]] = net.trafo.loc[trafo_index, parameters['net_' + element]]
            
            trafo_table.loc[offset + i,'p_hv_mw'] = trafo['p_hv_mw'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'q_hv_mvar'] = trafo['q_hv_mvar'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'p_lv_mw'] = trafo['p_lv_mw'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'q_lv_mvar'] = trafo['q_lv_mvar'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'pl_mw'] = trafo['pl_mw'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'ql_mvar'] = trafo['ql_mvar'].loc[step,trafo_index] 
            trafo_table.loc[offset + i,'loading_percent'] = trafo['loading_percent'].loc[step,trafo_index] 
            i = i + 1
        step = step + 1

    return trafo_table


def sort_bus_results(net, number, time_steps, bus, column, parameters):    
    bus_table = preallocate_table('bus', column, number)
    element = 'bus'
    i = 0     
    for bus_index in net.bus.index:
        bus_table.loc[i,'step'] = 0
        
        bus_table.loc[i,'index'] = bus_index
        bus_table.loc[i,parameters['net_' + element]] = net.bus.loc[bus_index, parameters['net_' + element]]
        
        if time_steps == 1:
            bus_table.loc[i,parameters['res_' + element]] = net.res_bus.loc[bus_index, parameters['res_' + element]]
        else: 
            bus_table.loc[i,'vm_pu'] = bus['vm_pu'].loc[0,bus_index]
            bus_table.loc[i,'va_degree'] = bus['va_degree'].loc[0,bus_index]
            bus_table.loc[i,'p_mw'] = bus['p_mw'].loc[0,bus_index]
            bus_table.loc[i,'q_mvar'] = bus['q_mvar'].loc[0,bus_index]
        i = i + 1
    
    print('\n Sorting the results: buses - 5/5')
    for step in tqdm(range(1,time_steps,1)):
        offset = step*number['bus']
        i = 0
        for bus_index in net.bus.index: 
            bus_table.loc[offset + i,'step'] = step
            
            bus_table.loc[offset + i,'index'] = bus_index
            bus_table.loc[offset + i,parameters['net_' + element]] = net.bus.loc[bus_index, parameters['net_' + element]]
            
            bus_table.loc[offset + i,'vm_pu'] = bus['vm_pu'].loc[step,bus_index]
            bus_table.loc[offset + i,'va_degree'] = bus['va_degree'].loc[step,bus_index]
            bus_table.loc[offset + i,'p_mw'] = bus['p_mw'].loc[step,bus_index]
            bus_table.loc[offset + i,'q_mvar'] = bus['q_mvar'].loc[step,bus_index]      
            i = i + 1
        step = step + 1
        
    return bus_table



def check_bus(net, time_steps, parameters, number):
    for i in range(len(parameters['net_bus'])):
         if not parameters['net_bus'][i] in net.bus.keys():
             net.bus[parameters['net_bus'][i]] = [None]*number['bus']
    if time_steps ==1:
        for i in range(len(parameters['res_bus'])):
            if not parameters['res_bus'][i] in net.res_bus.keys():
                net.res_bus[parameters['res_bus'][i]] = [None]*number['bus']      
    
    return net

def check_gen(net, time_steps, parameters, number, gen_fuel_tech):
    for i in range(len(parameters['net_gen'])):
         if not parameters['net_gen'][i] in net.gen.keys():
             net.gen[parameters['net_gen'][i]] = [None]*number['gen']
    if time_steps == 1:
        for i in range(len(parameters['res_gen'])):
             if not parameters['res_gen'][i] in net.res_gen.keys():
                 net.res_gen[parameters['res_gen'][i]] = [None]*number['gen']         
    if not len(gen_fuel_tech) > 0:
         d = {'fuel': ['---'], 'tech': ['---']}
         gen_fuel_tech = pd.DataFrame(data = d)
         df_aux = pd.DataFrame(data = d)
         for i in range(number['gen']):
             gen_fuel_tech = pd.concat([gen_fuel_tech, df_aux],ignore_index=True)  
    
    return net

def check_line(net, time_steps, parameters, number):
    for i in range(len(parameters['net_line'])):
         if not parameters['net_line'][i] in net.line.keys():
             net.line[parameters['net_line'][i]] = [None]*number['line']
    if time_steps == 1:
         for i in range(len(parameters['res_line'])):
             if not parameters['res_line'][i] in net.res_line.keys():
                 net.res_line[parameters['res_line'][i]] = [None]*number['line']     

    return net

def check_trafo(net, time_steps, parameters, number):
    for i in range(len(parameters['net_trafo'])):
         if not parameters['net_trafo'][i] in net.trafo.keys():
             net.trafo[parameters['net_trafo'][i]] = [None]*number['trafo']
    if time_steps == 1:
        for i in range(len(parameters['res_trafo'])):
             if not parameters['res_trafo'][i] in net.res_trafo.keys():
                 net.res_trafo[parameters['res_trafo'][i]] = [None]*number['trafo']      
    
    return net

