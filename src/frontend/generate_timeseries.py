"""
Generate networks for time series analysis.
"""

import pandas as pd
from pandapower.timeseries import DFData
from pandapower.control import ConstControl


def generate_timeseries(net, ts_path):
    """
    Generate controllers for time series analysis.

    Args:
        net: pandapower network.
        ts_path: path of time series data.

    Returns:
        net: pandapower network with controllers for ts.
        num_ts: number of time steps.
    """
    # initialize
    ts_file = pd.ExcelFile(ts_path)
    profiles = pd.DataFrame()

    # summarize all data to one profile as data source
    for sheet_name in ts_file.sheet_names:
        element, index = sheet_name.split('-')
        ts_data = ts_file.parse(sheet_name, index_col=0)
        for column in ts_data.columns[:-1]:
            if ts_data['type'][0] == 'scale':
                if element == 'sgen':
                    ts_data.loc[:, column] *= net.sgen.loc[int(index), column]
                elif element == 'gen':
                    ts_data.loc[:, column] *= net.gen.loc[int(index), column]
                elif element == 'load':
                    ts_data.loc[:, column] *= net.load.loc[int(index), column]

            profiles[sheet_name + '-' + column] = ts_data.loc[:, column]

    # replace all undefined value with the last defined value, create data source
    profiles.fillna(method='backfill', axis=0, inplace=True)
    profiles = profiles[:100]   # delete after development
    data_source = DFData(profiles)
    num_ts = len(profiles.index)

    # # generate controllers
    # for column in profiles.columns:
    #     element, index, variable = column.split('-')
    #     ConstControl(net, element=element, variable=variable, element_index=[index],
    #                  data_source=data_source, profile_name=[column])

    # generate controllers
    sgen_index_p = []
    sgen_index_q = []
    gen_index_p = []
    gen_index_q = []
    load_index_p = []
    load_index_q = []
    sgen_pn_p = []
    sgen_pn_q = []
    gen_pn_p = []
    gen_pn_q = []
    load_pn_p = []
    load_pn_q = []
    for column in profiles.columns:
        element, index, variable = column.split('-')
        index = int(index)
        if element == 'sgen' and variable == 'p_mw':
            sgen_index_p.append(index)
            sgen_pn_p.append(column)
        elif element == 'sgen' and variable == 'q_mvar':
            sgen_index_q.append(index)
            sgen_pn_q.append(column)
        elif element == 'gen' and variable == 'p_mw':
            gen_index_p.append(index)
            gen_pn_p.append(column)
        elif element == 'gen' and variable == 'q_mvar':
            gen_index_q.append(index)
            gen_pn_q.append(column)
        elif element == 'load' and variable == 'p_mw':
            load_index_p.append(index)
            load_pn_p.append(column)
        elif element == 'load' and variable == 'q_mvar':
            load_index_q.append(index)
            load_pn_q.append(column)

    ConstControl(net, element='sgen', variable='p_mw', element_index=sgen_index_p,
                 data_source=data_source, profile_name=sgen_pn_p)
    ConstControl(net, element='sgen', variable='q_mvar', element_index=sgen_index_q,
                 data_source=data_source, profile_name=sgen_pn_q)
    ConstControl(net, element='gen', variable='p_mw', element_index=gen_index_p,
                 data_source=data_source, profile_name=gen_pn_p)
    ConstControl(net, element='gen', variable='q_mvar', element_index=gen_index_q,
                 data_source=data_source, profile_name=gen_pn_q)
    ConstControl(net, element='load', variable='p_mw', element_index=load_index_p,
                 data_source=data_source, profile_name=load_pn_p)
    ConstControl(net, element='load', variable='q_mvar', element_index=load_index_q,
                 data_source=data_source, profile_name=load_pn_q)

    return net, num_ts
