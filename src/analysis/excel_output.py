# -*- coding: utf-8 -*-
"""
Created on Tue Nov 15 21:40:43 2022

@author: marti
"""
from openpyxl import load_workbook
from tqdm import tqdm
from src.analysis.sort_results import sort_results 
from src.analysis.parameters import check_parameter
from src.analysis.parameters import output_parameters
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
   
def create_excel(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps, results) :    
    # read the template, write the results and add important parameters from the network topology
    # finally save into a new excel workbook, according to the network topology and scenario name 

    # read paramters from net, and the ones that are going to be written in the excel
    number, column, parameters = output_parameters(net, gen_fuel_tech, scenario_name)
    # cell in the template were all the tables start
    initial_cell = 'C4' # row 3: all the parameters names
    initial_line = 5 # row 4: where the fisrt line of parameter values are written
    # read the template and retrieve the sheets
    output_template = 'src/analysis/output_templates/output_template.xlsm'
    print('\n Opening excel template')
    wb = load_workbook(filename = output_template, read_only = False, keep_vba = True) 
    print(' > Done')
    cell = {} # preallocate dictionary with last cell of tables, for references
    tables = {} # preallocate tables
    sheet = {'load': 'Demand',
             'bus': 'Buses',
             'gen':'Generation',
             'sgen':'Generation',
             'line': 'Lines',
             'trafo': 'Trafos'}
    table_name = sheet # in the excel template the tables have the same as the sheet 
    #%%
    for element in number:
        if number[element] > 0:  # if there is no load, jump to the next element            
            # checking values of the parameters, and adding columns and/or formatting them    
            net = check_parameter(net, time_steps, parameters, number, element) 
            # read the values from results, and sort them for the output
            tables[element] = sort_results(net, number, time_steps, results[element], column, parameters, element) 
            # write the values in excel table
            cell[element] = write_in_the_excel(tables[element], wb[sheet[element]], element, column, initial_line, number)
            # delete the results to free memory
            del results[element]
        else:
            print('\n No '+ element +'s in the net')
      
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

    return tables