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
    data_source = DFData(profiles)
    num_ts = len(profiles.index)

    # create controllers from data source
    controller_data = dict()
    for column in profiles.columns:
        element, index, variable = column.split('-')
        index = int(index)
        if element + '-' + variable not in controller_data.keys():
            controller_data[element + '-' + variable] = []
        controller_data[element + '-' + variable].append(index)

    for key in controller_data.keys():
        element, variable = key.split('-')
        profile_names = [element + '-' + str(index) + '-' + variable for index in controller_data[key]]
        ConstControl(net, element=element, variable=variable, element_index=controller_data[key],
                     data_source=data_source, profile_name=profile_names)

    return net, num_ts
