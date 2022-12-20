# -*- coding: utf-8 -*-
"""
Created on Tue Nov 29 00:02:46 2022

@author: thoug
"""

import numpy as np
import pandas as pd
import pandapower as pp

import tempfile


import Adv_network_only as addnet


import openpyxl as oxl 
from io import StringIO

net = addnet.net

text='-'
h_line = text.rjust(100,'-')

#%%
"""
Due to new change, I think I have to first make Analysis function different.
So it will be time-series compatiable.

Probably better to make one that read temp file to judge if something is over or not.


"""

output_dir = os.path.join(tempfile.gettempdir(), "propens_pandapower_time_series")


#######################################################
"""
Due to folder location issue, I couldn't run this file yet.

Bus basically

1. Read the temp file
2. Make each sheet and value I want to look at, a Variable
3. Based on the variable, it will analyze where's wrong...blahblah...
"""

#this function can prolly cover all the time series issues?
def Anal_read(output_dir):
    Bus_info= pd.read_excel(io=output_dir, sheet_name ='bus', index_col = 0 ) #reading from temp file, in this case, looking at bus sheet
    Trafo_info = pd.read_excel(io=output_dir, sheet_name ='trafo', index_col = 0 )
    Trafo3w_info = pd.read_excel(io=output_dir, sheet_name ='trafo3w', index_col = 0 )
    Line_info = pd.read_excel(io=output_dir, sheet_name ='line', index_col = 0 )
    
    Bus_Voltage_info = Bus_info['vm_pu']  # now each variable only have value of each column I would like to see
    Trafo_Loading_info = Trafo_info['loading_percent']
    Trafo3w_Loading_info = Trafo3w_info['loading_percent']
    Line_Loading_info = Line_info['loading_percent']
    
    if Bus_Voltage_info is True:
        under_bus = []
        under_value = []
        under_voltage=[]
        over_bus =[]
        over_value=[]
        over_voltage =[]

        
        for i in range(0, len(Bus_info)):
            if Bus_info[i] <0.95:
              print(f"bus %d is undervoltage, it is {format(Bus_info[i], '.4f')} p.u. now" % i)            #Based on Data of Map, maybe I can make function to color this node RED sth like that in case
              under_bus.append(i)
              under_value.append(Bus_info[i])
              under_voltage = {'index':under_bus,'value':under_value}
              under_voltage_df = pd.DataFrame.from_dict(data=under_voltage)
              
            if Bus_info[i]>1.03:
                print(f"bus %d is overvoltage, it is {format(Bus_info[i], '.4f')} p.u. now" % i)           #In case of over voltage, value is still arbitrary
                over_bus.append(i)
                over_value.append(Bus_info[i])
                over_voltage = {'index':over_bus,'value':over_value}           
                over_voltage_df = pd.DataFrame.from_dict(data=over_voltage)
              
            elif all( 0.95< i < 1.03 for i in Bus_info) == True:   #Incase when all are in standard, give one line of notice all are good
                print("There is no voltage violation")
                under_voltage = {'index':['---'],'value':['---']}           
                under_voltage_df = pd.DataFrame.from_dict(data=under_voltage)
                over_voltage = {'index':['---'],'value':['---']}           
                over_voltage_df = pd.DataFrame.from_dict(data=over_voltage)
                
    if Trafo_Loading_info is True:
        over_trafo_index = []
        over_trafo_value = []
        over_trafo = []
                   
        for i in range(0, len(Trafo_Loading_info)):
            if Trafo_info[i]>100.0:
                print(f"Transformer %d is overloaded, it is {format(Trafo_info[i], '.4f')} percent used" % i,  )
                over_trafo_index.append(i)
                over_trafo_value.append(Trafo_Loading_info[i])
                over_trafo = {'index':over_trafo_index,'value':over_trafo_value}
                over_trafo_df = pd.DataFrame(data=over_trafo)
            elif all(i < 100 for i in Trafo_info) == True:
                print("All Transformer is in standard")
                over_trafo = {'index':['---'],'value':['---']}
                over_trafo_df = pd.DataFrame(data=over_trafo)
                
    if Trafo3w_Loading_info is True:
        over_trafo3w_index = []
        over_trafo3w_value = []
        over_trafo3w = []
        
        for i in range(0, len(Trafo3w_Loading_info)):
            if Trafo3w_Loading_info[i]>100.0:
                print(f"3-winding Transformer %d is overloaded, it is {format(Trafo3w_Loading_info[i], '.4f')} percent used" % i,  )
                over_trafo3w_index.append(i)
                over_trafo3w_value.append(Trafo3w_Loading_info[i])
                over_trafo3w = {'index':over_trafo3w_index,'valuel':over_trafo3w_value}      
                over_trafo3w_df = pd.DataFrame(data=over_trafo3w)
            elif all(i < 100 for i in Trafo3w_Loading_info) == True:
                print("All 3-winding Transformer is in standard")
                over_trafo3w = {'index':['---'],'value':['---']}
                over_trafo3w_df = pd.DataFrame(data=over_trafo3w)
    
    if Line_Loading_info is True:
        over_line_index = []
        over_line_value = []
        over_line = []
        
              
        for i in range(0, len(Line_Loading_info)):
            if Line_Loading_info[i]>100.0:
                print(f"|line %d  overloadedd" %i, end='') 
                print(f": {format(Line_Loading_info[i], '.4f')}%|" )
                over_line_index.append(i)
                over_line_value.append(Line_Loading_info[i])
                over_line = {'index':over_line_index,'value':over_line_value}
                over_line_df = pd.DataFrame(data=over_line)
                
            elif all(i < 100 for i in Line_Loading_info) == True:
                print("|All line is in standard|")
                over_line = {'index':['---'],'value':['---']}           
                over_line_df = pd.DataFrame.from_dict(data=over_line)
        
    return under_voltage_df, over_voltage_df, over_trafo_df, over_trafo3w_df, over_line_df
        
    
        

    
    
    
    
    
    
    
    
    



#######################################################

def Anal(net):
    Anal_Bus_Under(net.res_bus)
    Anal_Bus_Over(net.res_bus)
    Anal_Trafo_Loading(net.res_trafo)
    Anal_Trafo3w_Loading(net.res_trafo3w)
    Anal_Line_Loading_Better(net.res_line)
    
#####################################################3
#####################################################3

def Anal_xl():
    Bus_Voltage_Under = Anal_Bus_Under(net.res_bus)
    Bus_Voltage_Over = Anal_Bus_Over(net.res_bus)
    Bus_Voltage_Sum = Bus_Voltage_Under.append(Bus_Voltage_Over, ignore_index = True)
    Trafo_Loading_Sum = Anal_Trafo_Loading(net.res_trafo)
    Trafo3w_Loading_Sum = Anal_Trafo3w_Loading(net.res_trafo3w)
    Line_Loading_Sum = Anal_Line_Loading_Better(net.res_line)
    
    
    
    return Bus_Voltage_Sum, Trafo_Loading_Sum, Trafo3w_Loading_Sum, Line_Loading_Sum

#####################################################3
#####################################################3
def Anal_pd():
    Bus_Voltage_Under = Anal_Bus_Under(net.res_bus)
    Bus_Voltage_Over = Anal_Bus_Over(net.res_bus)
    pd_Bus_Voltage = Bus_Voltage_Under.append(Bus_Voltage_Over, ignore_index = True)
    
    pd_Line_Loading = Anal_Line_Loading_Better(net.res_line)
    pd_trafo_Loading = Anal_Trafo_Loading(net.res_trafo)
    pd_trafo3w_Loading = Anal_Trafo3w_Loading(net.res_trafo3w)
    

    
    return pd_Bus_Voltage, pd_Line_Loading, pd_trafo_Loading, pd_trafo3w_Loading

#####################################################3
#####################################################3

def Anal_Bus_Voltage(x):
    s = x['vm_pu']
    under_bus = []
    under_value = []
    under_voltage=[]
    over_bus =[]
    over_value=[]
    over_voltage =[]
    
    """
    Store the value of Bus and p.u. of voltage and save it in a list, and return it 
    """
    
    for i in range(0, len(x)):
        if s[i] <0.95:
          print(f"bus %d is undervoltage, it is {format(s[i], '.4f')} p.u. now" % i)            #Based on Data of Map, maybe I can make function to color this node RED sth like that in case
          under_bus.append(i)
          under_value.append(s[i])
          under_voltage = {'index':under_bus,'value':under_value}
          under_voltage_df = pd.DataFrame.from_dict(data=under_voltage)

          #return under_voltage
          #print(under_voltage)
        elif s[i]>1.03:
            print(f"bus %d is overvoltage, it is {format(s[i], '.4f')} p.u. now" % i)           #In case of over voltage, value is still arbitrary
            over_bus.append(i)
            over_value.append(s[i])
            over_voltage = {'index':over_bus,'value':over_value}           
            over_voltage_df = pd.DataFrame.from_dict(data=over_voltage)
           
            #return over_voltage
            
        elif all(i < 1.0 for i in s) == True:   #Incase when all are in standard, give one line of notice all are good
            print("All Bus Voltage is in standard")
    
    
    return under_voltage_df, over_voltage_df
    print('----------Bus Analysis Over----------------')
    
#####################################################3
#####################################################3

def Anal_Bus_Under(x):    
    s = x['vm_pu']
    under_bus = []
    under_value = []
    under_voltage=[]
    
    for i in range(0, len(x)):
        if s[i] <0.95:
          print(f"bus %d is undervoltage, it is {format(s[i], '.4f')} p.u. now" % i)            #Based on Data of Map, maybe I can make function to color this node RED sth like that in case
          under_bus.append(i)
          under_value.append(s[i])
          under_voltage = {'index':under_bus,'value':under_value}
          under_voltage_df = pd.DataFrame.from_dict(data=under_voltage)
          
        elif all(i < 1.0 for i in s) == True:   #Incase when all are in standard, give one line of notice all are good
            print("There is no undervoltage violation")
            under_voltage = {'index':['---'],'value':['---']}           
            under_voltage_df = pd.DataFrame.from_dict(data=under_voltage)
            
    return under_voltage_df

def Anal_Bus_Over(x):
    s = x['vm_pu']
    over_bus =[]
    over_value=[]
    over_voltage =[]
    
    for i in range(0, len(x)):
        if s[i]>1.03:
            print(f"bus %d is overvoltage, it is {format(s[i], '.4f')} p.u. now" % i)           #In case of over voltage, value is still arbitrary
            over_bus.append(i)
            over_value.append(s[i])
            over_voltage = {'index':over_bus,'value':over_value}           
            over_voltage_df = pd.DataFrame.from_dict(data=over_voltage)
           
            #return over_voltage
            
        elif all(i < 1.0 for i in s) == True:   #Incase when all are in standard, give one line of notice all are good
            print("There is no overvoltage violation")
            over_voltage = {'index':['---'],'value':['---']}           
            over_voltage_df = pd.DataFrame.from_dict(data=over_voltage)
            
    return over_voltage_df
        

    
#####################################################3


def Anal_Trafo_Loading(x):
    s = x['loading_percent']
    
    over_trafo_index = []
    over_trafo_value = []
    over_trafo = []
               
    for i in range(0, len(x)):
        if s[i]>100.0:
            print(f"Transformer %d is overloaded, it is {format(s[i], '.4f')} percent used" % i,  )
            over_trafo_index.append(i)
            over_trafo_value.append(s[i])
            over_trafo = {'index':over_trafo_index,'value':over_trafo_value}
            over_trafo_df = pd.DataFrame(data=over_trafo)
        elif all(i < 100 for i in s) == True:
            print("All Transformer is in standard")
            over_trafo = {'index':['---'],'value':['---']}
            over_trafo_df = pd.DataFrame(data=over_trafo)
    
    
    return over_trafo_df
    print('----------Transformer Analysis Over----------------')    
         
#####################################################3

def Anal_Trafo3w_Loading(x):
    s = x['loading_percent']
    a = all(s)
   
    over_trafo3w_index = []
    over_trafo3w_value = []
    over_trafo3w = []
    
    for i in range(0, len(x)):
        if s[i]>100.0:
            print(f"3-winding Transformer %d is overloaded, it is {format(s[i], '.4f')} percent used" % i,  )
            over_trafo3w_index.append(i)
            over_trafo3w_value.append(s[i])
            over_trafo3w = {'index':over_trafo3w_index,'valuel':over_trafo3w_value}      
            over_trafo3w_df = pd.DataFrame(data=over_trafo3w)
        elif all(i < 100 for i in s) == True:
            print("All 3-winding Transformer is in standard")
            over_trafo3w = {'index':['---'],'value':['---']}
            over_trafo3w_df = pd.DataFrame(data=over_trafo3w)
            
    return over_trafo3w_df
    print('----------3 Winding Transformer Analysis Over----------------')
   
#####################################################3
#####################################################3

def Anal_Line_Loading_Better(x):
    s = x['loading_percent']
    
    over_line_index = []
    over_line_value = []
    over_line = []
    
    for i in range(0, len(x)):
        if all(i<100 for i in s) == False:
            print(h_line)
            print("There are issues in loading of the line")
            print()
            print("| Line     State      Percentage|")
            print("_________________________________")
            
            break
          
    
    for i in range(0, len(x)):
        if s[i]>100.0:
            print(f"|line %d  overloadedd" %i, end='') 
            print(f": {format(s[i], '.4f')}%|" )
            over_line_index.append(i)
            over_line_value.append(s[i])
            over_line = {'index':over_line_index,'value':over_line_value}
            over_line_df = pd.DataFrame(data=over_line)
            
        elif all(i < 100 for i in s) == True:
            print("|All line is in standard|")
            over_line = {'index':['---'],'value':['---']}           
            over_line_df = pd.DataFrame.from_dict(data=over_line)
            
    return over_line_df
    print('----------Line Analysis Over----------------')
   
      
#####################################################3