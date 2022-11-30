# -*- coding: utf-8 -*-
"""
Created on Tue Nov 29 00:02:46 2022

@author: thoug
"""

import numpy as np
import pandas as pd
import pandapower as pp


import Adv_network_only as addnet


import openpyxl as oxl 
from io import StringIO

net = addnet.net

text='-'
h_line = text.rjust(100,'-')

"""
By Calling Anal(net) function, now I can call all the sub function all at once
"""

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