# -*- coding: utf-8 -*-
"""
Created on Wed Jan 11 07:10:14 2023

@author: Taeyoung
"""

import numpy as np
import pandas as pd
import pandapower as pp

import tempfile


# import Adv_network_only as addnet

import os
import openpyxl as oxl 
from io import StringIO

from tqdm import tqdm
import time

# net = addnet.net

text='-'
h_line = text.rjust(100,'-')

# path_result = "./../../result/results_Network.xlsm"
path_result = './result/results_Network.xlsm'
path_test = "C:/Users/thoug/OneDrive/WS2022/ENS_Panda/Feb/result/results_Network.xlsm"


#%%

class pd_Analysis:
    
    vol_under = 0.99
    vol_over = 1.03
    loading_over = 10
    
    
    def __init__(self):
        self.output_dir = os.path.join(path_result)
        
        self.excel_to_write = os.path.join(path_result)
        
        self.vol_under = 0.99
        self.vol_over  = 1.03
        self.loading_over = 10
        
        self.sheet_name = ''
        self.col_name = ''
        
        self.kwargs =''
        
        
    def anal_sheet(self, sheet_name):
        
        info_sheet= pd.read_excel(io=self.output_dir, sheet_name =sheet_name, index_col =3 , skiprows=2 )
        
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
            
            for i in tqdm(range(0, len(info_col))):
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
            
            for i in tqdm(range(0, len(info_col))):
                
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
                
                for i in tqdm(range(0, len(info_col))):
                    
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
        
    
    def call_anal():
        
        print("Analyze the output grid")
        aa = pd_Analysis()
        bb1 = aa.anal_sheet('Buses')
        bb2 = aa.anal_sheet('Lines')
        bb3 = aa.anal_sheet('Trafos')
        cc1 = aa.anal_col(bb1,'index','step','vm_pu') #info_col
        cc2 = aa.anal_col(bb2,'index','step','loading_percent') #info_col
        cc3 = aa.anal_col(bb3,'index','step','loading_percent') #info_col
        print('Analyzing Voltage of the Bus')
        dd1 = aa.anal_vol(cc1) 
        print('Analyzing Loading of the Lines')
        dd2 = aa.anal_line(cc2) 
        print('Analyzing Loading of the Transformer')
        dd3 = aa.anal_trafo(cc3)
        print('You can find Analysis information on Summary Tab')
        ee1 = aa.anal_excel_out(df1=dd1, df2=dd2, df3=dd3)
        
        return
        
        
#%%
#test space
# a = pd_Analysis()
# b1 = a.anal_sheet('Buses')
# b2 = a.anal_sheet('Lines')

# c1 = a.anal_col(b1,'Voltage [p.u]')
# c2 = a.anal_col(b2, 'Loading Percent [%]')

# d1 = a.anal_vol(c1)
# d2 = a.anal_line_load(c2)

# #%%
# e=a.anal_excel_out(df1=d1, df2=d2)
        
        
        
        
#%%

class pd_ts_Analysis(pd_Analysis):
    
    nr_ts = 95
    
    def __init__(self):
        
        super().__init__()
        # self.output_ts_dir = os.path.join("C:/Users/thoug/OneDrive/WS2022/ENS_Panda/Feb/result/results_Network.xlsm")
        self.output_ts_dir = os.path.join(path_result)
        
        self.nr_ts = 95
        

        
        
    def anal_sheet(self, sheet_name):
        
        info_sheet= pd.read_excel(io=self.output_ts_dir, sheet_name =sheet_name, index_col =None , skiprows=3 ) #index_col = 1 : Time Step / index_col=3 : Bus Index / None for just regular
        
        return info_sheet
        
        
    def anal_col(self, info_sheet, bus_name, ts_name, scenario_name, col_name):
        
        bus_index = info_sheet[bus_name]
        ts_index = info_sheet[ts_name]  #ts_name = 'step'
        information_col = info_sheet[col_name]
        
        scenario_name = info_sheet[scenario_name]
        
        number_col_index = pd.DataFrame(range(len(information_col)))
        
        info_col = pd.concat([ts_index, bus_index, scenario_name, information_col], axis=1)
        
        return info_col
    
    def anal_vol(self, info_col, kwargs='vm_pu'):
        if kwargs == 'vm_pu':
            
            vol_under = 0.98 #originally 0.98
            vol_over = 1.02
            
            # info_per_bus = []
            
            # info_col_id = list(info_col.index)
            
            info_filter = info_col.replace({'vm_pu':{'---':np.nan}})
            
            info_here = info_filter.values
            
            index_a = []
            index_a_value = []
            index_a_extra = []
            index_a_df = []
            index_aa = pd.DataFrame()
            
            index_b = []
            index_b_value = []
            index_b_extra = []
            index_b_df = []
            index_bb = pd.DataFrame()
            
            index_c=[]
            index_c_value=[]
            index_c_extra=[]
            index_cc = pd.DataFrame()
            
            
            index_ab_df = []
            
            index_ab=[]
            index_abc=[]
            index_abc_df=[]
            
            info_series = info_col[kwargs].squeeze()

            
            
            # for i in range(len(info_col)):                # i=0 --> bus1~178 = one time step  // i=179 --> bus1~178 = two time step
            for j in tqdm(range(int(len(info_col)))):
                # info_here[j][1]= info_col[kwargs][j]
                
                if info_here[j][3] < vol_under:
                    
                    index_a.append(j)
                    index_a_value.append(info_here[j])
                    index_a_extra =pd.DataFrame.from_records(data=index_a_value, columns=['Time Step', 'Bus Index', 'Scenario', 'Under Voltage [p.u]'])
                    
                    index_aa = pd.DataFrame(index_a_extra)
                    
                    # index_a_df = pd.DataFrame.from_dict(data=index_a_extra)
                    
                    index_ab_df = index_a_extra
                    
                    # i =+ len(info_col)/self.nr_ts
                    
                elif info_here[j][3] > vol_over:
                    
                    index_b.append(j)
                    index_b_value.append(info_here[j])
                    index_b_extra =pd.DataFrame.from_records(data=index_b_value, columns=['Time Step', 'Bus Index','Scenario', 'Over Voltage [p.u]'])
                    
                    index_ab_df = index_b_extra
                    
                    index_bb = pd.DataFrame(index_b_extra)
                
                elif all(vol_under < j < vol_over for j in info_series) == True:
                    
                    index_c_value = {'Time Step':['---'],'Bus Index':['---'], 'Scenario':['---'],'Voltage [p.u]':['---']}  
                    index_c_extra = pd.DataFrame(index_c_value)
                    
                    index_ab_df = index_c_extra
                    
                    return index_ab_df  #if there's no error, it return this
                    
            time.sleep(1)
            # index_ab = index_a_extra.append(index_b_extra)
            
            index_ab = pd.concat([index_aa, index_bb], axis=1)

            
            return index_ab
        
    def anal_line(self, info_col, kwargs='line'):
        if kwargs == 'line':
            
            line_over = 100 #100
            
            info_here_line = info_col.values
            info_here_series = info_col['loading_percent'].squeeze()
            
            index_d=[]
            index_d_value=[]
            index_d_extra=[]
            
            for j in tqdm(range(int(len(info_col)))):
                if info_here_line[j][3] > line_over:
                    
                    index_d.append(j)
                    index_d_value.append(info_here_line[j])
                    index_d_extra = pd.DataFrame.from_records(data=index_d_value, columns=['Time Step', 'Line Index','Scenario', 'Loading Percent[%]'])
                    
                if all(j < line_over for j in info_here_series) == True:
                    
                    index_d_value = {'Time Step':['---'],'Line Index':['---'], 'Scenario':['---'], 'Loading Percentage[%]':['---']}  
                    index_d_extra = pd.DataFrame(index_d_value)
                    
            time.sleep(1)
            return index_d_extra
        
        
    def anal_trafo(self, info_col, kwargs='trafos'):
        if kwargs == 'trafos':
            
            trafo_over = 100 #100
            
            info_here = info_col.values
            
            info_here_series = info_col['loading_percent'].squeeze()
            
            index_c=[]
            index_c_value=[]
            index_c_extra=[]
            
            for j in tqdm(range(int(len(info_col)))):
                if len(info_col) > 1:
                    if info_here[j][3] > trafo_over:
                        
                        index_c.append(j)
                        index_c_value.append(info_here[j])
                        index_c_extra = pd.DataFrame.from_records(data=index_c_value, columns=['Time Step', 'Trafo Index','Scenario', 'Loading Percent[%]'])
                        
                    if all(j < trafo_over for j in info_here_series) == True:
                        
                        index_c_value = {'Time Step':['---'],'Trafo Index':['---'], 'Scenario':['---'],  'Loading Percentage[%]':['---']}  
                        index_c_extra = pd.DataFrame(index_c_value)
                        
                else:   # incase only one transformer exist
                    if info_here[0][3] >= trafo_over:
                        
                        index_c_extra = pd.DataFrame(data=info_here, columns=['Time Step', 'Trafo Index', 'Scenario', 'Loading Percent[%]'] )
                        
                    elif info_here[0][3] < trafo_over:
                        
                        index_c_value = {'Time Step':['---'],'Trafo Index':['---'], 'Scenario':['---'], 'Loading Percentage[%]':['---']}  
                        index_c_extra = pd.DataFrame(index_c_value)
                    
            time.sleep(1)            
            return index_c_extra
            
            
            
    
    def anal_excel_out(self, df1=None, df2=None, df3=None):
        
        
        
        excel_writer = self.output_ts_dir
        sheet_name = 'Summary'
        float_format = '%.4f'
        
        #Writing starts from B-20
        
        startrow = 15
        startcol = 2
        
        df = pd.concat([df1, df2, df3], axis=1)
        
        with pd.ExcelWriter(
                path= self.output_ts_dir,
                mode = 'a',
                engine = 'openpyxl',
                if_sheet_exists='overlay',
                engine_kwargs={"keep_vba": True})as writer:
            df.to_excel(excel_writer=writer, sheet_name=sheet_name, float_format=float_format, startrow=startrow, startcol=startcol)
            
            
    def call_anal():
        
        print("Analyze the output grid")
        aa = pd_ts_Analysis()
        bb1 = aa.anal_sheet('Buses')
        bb2 = aa.anal_sheet('Lines')
        bb3 = aa.anal_sheet('Trafos')
        cc1 = aa.anal_col(bb1,'index','step','scenario','vm_pu') #info_col
        cc2 = aa.anal_col(bb2,'index','step','scenario','loading_percent') #info_col   #need to update scenario tab
        cc3 = aa.anal_col(bb3,'index','step', 'scenario','loading_percent') #info_col
        print('Analyzing Voltage of the Bus')
        dd1 = aa.anal_vol(cc1) 
        print('Analyzing Loading of the Lines')
        dd2 = aa.anal_line(cc2) 
        print('Analyzing Loading of the Transformer')
        dd3 = aa.anal_trafo(cc3)
        print('You can find Analysis information on Summary Tab')
        ee1 = aa.anal_excel_out(df1=dd1, df2=dd2, df3=dd3)
        
        return
    
    
#%%

if __name__ == "__main__":
    


        aa = pd_ts_Analysis()
        bb1 = aa.anal_sheet('Buses')
        bb2 = aa.anal_sheet('Lines')
        bb3 = aa.anal_sheet('Trafos')
        cc1 = aa.anal_col(bb1,'index','step','scenario','vm_pu') #info_col
        cc2 = aa.anal_col(bb2,'index','step','scenario','loading_percent') #info_col   #need to update scenario tab
        cc3 = aa.anal_col(bb3,'index','step', 'scenario','loading_percent') #info_col
        print('Analyzing Voltage of the Bus')
        dd1 = aa.anal_vol(cc1) 
        print('Analyzing Loading of the Lines')
        dd2 = aa.anal_line(cc2) 
        print('Analyzing Loading of the Transformer')
        dd3 = aa.anal_trafo(cc3)
        print('You can find Analysis information on Summary Tab')
        ee1 = aa.anal_excel_out(df1=dd1, df2=dd2, df3=dd3)



#%%

"""
for i in range(0, q):
    for j in range(0, int(q/w)):
        m.append(j)
        n.append(i)
        
        i += int(q/w)
"""
    

        
        
        
        
        
        
        

        
        
        
       
                    
                
