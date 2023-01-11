# -*- coding: utf-8 -*-
"""
Created on Wed Jan 11 07:10:14 2023

@author: thoug
"""

import numpy as np
import pandas as pd
import pandapower as pp

import tempfile


import Adv_network_only as addnet

import os
import openpyxl as oxl 
from io import StringIO

net = addnet.net

text='-'
h_line = text.rjust(100,'-')
#%%

class pd_Analysis:
    def __init__(self):
        self.output_dir = os.path.join("C:/Users/thoug/OneDrive/WS2022/ENS_Panda/2023/result/results_Network01_Scenario01.xlsm")
        
        self.excel_to_write = os.path.join("C:/Users/thoug/OneDrive/WS2022/ENS_Panda/2023/result/result_anal_try.xlsx")
        
        self.vol_under = 0.99
        self.vol_over  = 1.03
        self.loading_over = 10
        
        self.sheet_name = ''
        self.col_name = ''
        
        self.kwargs =''
        
        
    def anal_sheet(self, sheet_name):
        info_sheet= pd.read_excel(io=self.output_dir, sheet_name =sheet_name, index_col = 3, skiprows=2 )
        
        return info_sheet
        
    def anal_col(self, info_sheet, col_name):
        info_col = info_sheet[col_name]
        
        return info_col
        
        
    def anal_func(self, info_col):
       
        index_a = []
        index_a_value = []
        index_a_extra = []
        index_a_df = []
        
        index_b = []
        index_b_value = []
        index_b_extra = []
        index_b_df = []
        
        index_ab_df = []
        
        index_c = []
        index_c_value = []
        index_c_extra = []
        index_c_df = []
        

        
        for i in range(0, len(info_col)):
            if info_col[i] < self.vol_under:
                print(f"bus %d is undervoltage, it is {format(info_col[i], '.4f')} p.u. now" % i)
                
                index_a.append(i)
                index_a_value.append(info_col[i])
                index_a_extra ={'under voltage':index_a,'value':index_a_value}
                index_a_df = pd.DataFrame.from_dict(data=index_a_extra)
                
                # return index_a_df
                
            elif self.loading_over > info_col[i] > self.vol_over:
                print(f"bus %d is overvoltage, it is {format(info_col[i], '.4f')} p.u. now" % i)
                
                index_b.append(i)
                index_b_value.append(info_col[i])
                index_b_extra ={'over voltage':index_b,'value':index_b_value}
                index_b_df = pd.DataFrame.from_dict(data=index_b_extra)
                
                # return index_b_df
                
            elif all( self.vol_under < i < self.vol_over for i in info_col) == True:
                print("There is no voltage violation")
                
                index_a_extra = {'under voltage':['---'],'value':['---']}           
                index_a_df = pd.DataFrame.from_dict(data=index_a_extra)
                index_b_extra = {'over voltage':['---'],'value':['---']}           
                index_b_df = pd.DataFrame.from_dict(data=index_b_extra)
                index_ab_df = index_a_df + index_b_df
                
                # return index_ab_df
                
            elif info_col[i] > self.loading_over:
                print(f"|component %d  overloadedd" %i, end='') 
                print(f": {format(info_col[i], '.4f')}%|" )
                index_c.append(i)
                index_c_value.append(info_col[i])
                index_c_extra = {'over loading':index_c,'value':index_c_value}
                index_c_df = pd.DataFrame(data=index_c_extra)
                
                # return index_c_df
                
            elif all(10< i < self.loading_over for i in info_col) == True:
                print("|All line is in standard|")
                index_c_extra = {'over loading':['---'],'value':['---']}           
                index_c_df = pd.DataFrame.from_dict(data=index_c_extra)
                
                # return index_c_df
        
        return index_a_df, index_b_df, index_ab_df, index_c_df
        
    def anal_vol(self, info_col, kwargs='vol'):
        if kwargs == 'vol':
            
            index_a = []
            index_a_value = []
            index_a_extra = []
            index_a_df = []
            
            index_b = []
            index_b_value = []
            index_b_extra = []
            index_b_df = []
            
            index_ab_df = []
            
            for i in range(0, len(info_col)):
                if info_col[i] < self.vol_under:
                    print(f"bus %d is undervoltage, it is {format(info_col[i], '.4f')} p.u. now" % i)
                    
                    index_a.append(i)
                    index_a_value.append(info_col[i])
                    index_a_extra ={'under voltage':index_a,'value':index_a_value}
                    index_a_df = pd.DataFrame.from_dict(data=index_a_extra)
                    
                    index_ab_df = index_a_df
                    
                    # return index_a_df
                    
                elif self.loading_over > info_col[i] > self.vol_over:
                    print(f"bus %d is overvoltage, it is {format(info_col[i], '.4f')} p.u. now" % i)
                    
                    index_b.append(i)
                    index_b_value.append(info_col[i])
                    index_b_extra ={'over voltage':index_b,'value':index_b_value}
                    index_b_df = pd.DataFrame.from_dict(data=index_b_extra)
                    
                    index_ab_df = index_b_df
                    
                    # return index_b_df
                    
                elif all( self.vol_under < i < self.vol_over for i in info_col) == True:
                    print("There is no voltage violation")
                    
                    index_a_extra = {'index':['---'],'value':['---']}           
                    index_a_df = pd.DataFrame.from_dict(data=index_a_extra)
                    index_b_extra = {'index':['---'],'value':['---']}           
                    index_b_df = pd.DataFrame.from_dict(data=index_b_extra)
                    index_ab_df = index_a_df + index_b_df
                        
            return index_ab_df
        
    def anal_line_load(self, info_col, kwargs='line'):
        if kwargs == 'line':
            
            index_c = []
            index_c_value = []
            index_c_extra = []
            index_c_df = []
            
            for i in range(0, len(info_col)):
                
                if info_col[i] > self.loading_over:
                    print(f"|component %d  overloadedd" %i, end='') 
                    print(f": {format(info_col[i], '.4f')}%|" )
                    index_c.append(i)
                    index_c_value.append(info_col[i])
                    index_c_extra = {'Line over loading':index_c,'value':index_c_value}
                    index_c_df = pd.DataFrame(data=index_c_extra)
                    
                    # return index_c_df
                    
                elif all(i < self.loading_over for i in info_col) == True:
                    print("|All line is in standard|")
                    index_c_extra = {'Line over loading':['---'],'value':['---']}           
                    index_c_df = pd.DataFrame.from_dict(data=index_c_extra)
                
            return index_c_df
        
        def anal_trafo_load(self, info_col, kwargs='trafo'):
            if kwargs == 'trafo':
                
                index_c = []
                index_c_value = []
                index_c_extra = []
                index_c_df = []
                
                for i in range(0, len(info_col)):
                    
                    if info_col[i] > self.loading_over:
                        print(f"|component %d  overloadedd" %i, end='') 
                        print(f": {format(info_col[i], '.4f')}%|" )
                        index_c.append(i)
                        index_c_value.append(info_col[i])
                        index_c_extra = {'Transformer over loading':index_c,'value':index_c_value}
                        index_c_df = pd.DataFrame(data=index_c_extra)
                        
                        # return index_c_df
                        
                    elif all(i < self.loading_over for i in info_col) == True:
                        print("|All transformer is in standard|")
                        index_c_extra = {'over loading':['---'],'value':['---']}           
                        index_c_df = pd.DataFrame.from_dict(data=index_c_extra)
                    
                return index_c_df
        
            
        
    
    def anal_excel_out(self, df1=None, df2=None, df3=None):
        
        
        
        excel_writer = self.excel_to_write
        sheet_name = 'Summary'
        float_format = '%.4f'
        
        #Writing starts from B-20
        
        startrow = 20
        startcol = 2
        
        df = pd.concat([df1, df2, df3], axis=1)
        
        with pd.ExcelWriter(
                path= self.output_dir,
                mode = 'a',
                engine = 'openpyxl',
                if_sheet_exists='overlay',
                engine_kwargs={"keep_vba": True})as writer:
            df.to_excel(excel_writer=writer, sheet_name=sheet_name, float_format=float_format, startrow=startrow, startcol=startcol)
        
        
        
#%%
#test space
a = pd_Analysis()
b1 = a.anal_sheet('Buses')
b2 = a.anal_sheet('Lines')

c1 = a.anal_col(b1,'Voltage [p.u]')
c2 = a.anal_col(b2, 'Loading Percent [%]')

d1 = a.anal_vol(c1)
d2 = a.anal_line_load(c2)

#%%
e=a.anal_excel_out(df1=d1, df2=d2)
        
        
        
        
        

        
        
        
       
                    
                
