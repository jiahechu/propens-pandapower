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
        None
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

    # generate controllers
    for column in profiles.columns:
        element, index, variable = column.split('-')
        ConstControl(net, element=element, variable=variable, element_index=index,
                     data_source=data_source, profile_name=column)
