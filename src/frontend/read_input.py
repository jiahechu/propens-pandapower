"""
Read network data from Excel file.
"""

import pandapower as pp
import pandas as pd
import tempfile


def read_input(scenario_path, topology_path):
    """
    Read network data from Excel file.

    Args:
        scenario_path: path of the Excel file with scenario data.
        topology_path: path of the Excel file with topology data.

    Returns:
        net: pandapower network.
        ts_setup: time series setup.
        fuel_data: fuel data for technologies.
    """
    # init
    tempxlsx = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)

    # get time series information
    ts_setup = pd.read_excel(scenario_path, sheet_name='general')

    # get fuel information
    fuel_data = pd.read_excel(scenario_path, sheet_name='fuel')

    # read data from topology and scenario
    topology_xlsx = pd.ExcelFile(topology_path)
    scenario_xlsx = pd.ExcelFile(scenario_path)

    # merge topology and scenario into one excel file
    writer = pd.ExcelWriter(tempxlsx.name)
    for name in topology_xlsx.sheet_names:
        topology_xlsx.parse(sheet_name=name, index_col=0).to_excel(writer, sheet_name=name)
    for name in scenario_xlsx.sheet_names:
        scenario_xlsx.parse(sheet_name=name, index_col=0).to_excel(writer, sheet_name=name)
    writer.close()

    # read network from final excel file
    net = pp.from_excel(tempxlsx.name)
    tempxlsx.close()

    return net, ts_setup, fuel_data