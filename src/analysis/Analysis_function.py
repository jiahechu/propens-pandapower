# -*- coding: utf-8 -*-
"""
Created on Wed Nov  9 17:32:44 2022

@author: thoug
"""

#so far we use arbitrary number to check function work or not

import pandapower as pp
net = pp.create_empty_network()

a = 'net.res_bus'
b = eval(a)

c_line = ''



text='-'
h_line = text.rjust(100,'-')

m=[]
k=[]

def Anal_Bus_Voltage(x):
    s = x['vm_pu']
    
    for i in range(0, len(x)):
        if s[i] <0.95:
          print(f"bus %d is undervoltage, it is {format(s[i], '.4f')} p.u. now" % i)            #Based on Data of Map, maybe I can make function to color this node RED sth like that in case
        elif s[i]>1.03:
            print(f"bus %d is overvoltage, it is {format(s[i], '.4f')} p.u. now" % i)           #In case of over voltage, value is still arbitrary
            
        elif all(i < 1.0 for i in s) == True:   #Incase when all are in standard, give one line of notice all are good
            print("All Bus Voltage is in standard")
         
#####################################################3

def Anal_Line_Loading(x):
    s = x['loading_percent']
    
  

    for i in range(0, len(x)):
        if s[i]>100.0:
            print(f"line %d is overloaded, it is {format(s[i], '.4f')} percent used" % i)
            c_line = "#e66363" #incase of color changed
            return c_line
        elif all(i < 100 for i in s) == True:
            print("All line is in standard")
            c_line = "#000000"
            return c_line
     
            
            
#####################################################3


def Anal_Trafo_Loading(x):
    s = x['loading_percent']
               
    for i in range(0, len(x)):
        if s[i]>100.0:
            print(f"Transformer %d is overloaded, it is {format(s[i], '.4f')} percent used" % i,  )
        elif all(i < 100 for i in s) == True:
            print("All Transformer is in standard")
  
            
         
#####################################################3

def Anal_Trafo3w_Loading(x):
    s = x['loading_percent']
    a = all(s)
   
    
    for i in range(0, len(x)):
        if s[i]>100.0:
            print(f"3-winding Transformer %d is overloaded, it is {format(s[i], '.4f')} percent used" % i,  )
        elif all(i < 100 for i in s) == True:
            print("All 3-winding Transformer is in standard")

   
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
   
      
#####################################################3
"""   
#Calling Anal() will make all other sub function to be called <-- still don't know how to do it
def Anal():
   Anal_Bus_Voltage(b)
"""


