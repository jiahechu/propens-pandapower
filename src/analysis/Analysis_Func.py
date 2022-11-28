# -*- coding: utf-8 -*-
"""
Created on Tue Nov 29 00:02:46 2022

@author: thoug
"""

import numpy as np
import pandas as pd
import pandapower as pp


import Adv_network_only as addnet


net = addnet.net

text='-'
h_line = text.rjust(100,'-')

"""
By Calling Anal(net) function, now I can call all the sub function all at once
"""

########################################################

def Anal(net):
    Anal_Bus_Voltage(net.res_bus)
    Anal_Trafo_Loading(net.res_trafo)
    Anal_Trafo3w_Loading(net.res_trafo3w)
    Anal_Line_Loading_Better(net.res_line)

#####################################################3
#####################################################3

def Anal_Bus_Voltage(x):
    s = x['vm_pu']
    
    for i in range(0, len(x)):
        if s[i] <0.95:
          print(f"bus %d is undervoltage, it is {format(s[i], '.4f')} p.u. now" % i)            #Based on Data of Map, maybe I can make function to color this node RED sth like that in case
        elif s[i]>1.03:
            print(f"bus %d is overvoltage, it is {format(s[i], '.4f')} p.u. now" % i)           #In case of over voltage, value is still arbitrary
            
        elif all(i < 1.0 for i in s) == True:   #Incase when all are in standard, give one line of notice all are good
            print("All Bus Voltage is in standard")
    
    print('----------Bus Analysis Over----------------')
    
#####################################################3


def Anal_Trafo_Loading(x):
    s = x['loading_percent']
               
    for i in range(0, len(x)):
        if s[i]>100.0:
            print(f"Transformer %d is overloaded, it is {format(s[i], '.4f')} percent used" % i,  )
        elif all(i < 100 for i in s) == True:
            print("All Transformer is in standard")
  
    print('----------Transformer Analysis Over----------------')    
         
#####################################################3

def Anal_Trafo3w_Loading(x):
    s = x['loading_percent']
    a = all(s)
   
    
    for i in range(0, len(x)):
        if s[i]>100.0:
            print(f"3-winding Transformer %d is overloaded, it is {format(s[i], '.4f')} percent used" % i,  )
        elif all(i < 100 for i in s) == True:
            print("All 3-winding Transformer is in standard")
            
    print('----------3 Winding Transformer Analysis Over----------------')
   
#####################################################3
#####################################################3

def Anal_Line_Loading_Better(x):
    s = x['loading_percent']
    
    
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
        elif all(i < 100 for i in s) == True:
            print("|All line is in standard|")
    print('----------Line Analysis Over----------------')
   
      
#####################################################3