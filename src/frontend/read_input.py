"""
Read network data from Excel file.
"""

import pandapower as pp
import pandas as pd


def read_input(loadgen_path, topology_path, result_path):
    """
    Read network data from Excel file.

    Args:
        loadgen_path: path of the Excel file with load/gen data.
        topology_path: path of the Excel file with topology data.
        result_path: path to store results.

    Returns:
        net: pandapower network.
    """
    # init
    net = pp.create_empty_network()
    result_path += './network.xlsx'

    # read data from topology and load/gen
    topology_xlsx = pd.ExcelFile(topology_path)
    loadgen_xlsx = pd.ExcelFile(loadgen_path)

    # merge topology and load/gen into one excel file
    writer = pd.ExcelWriter(result_path)    # pylint: disable=abstract-class-instantiated
    for name in topology_xlsx.sheet_names:
        topology_xlsx.parse(sheet_name=name, index_col=0).to_excel(writer, sheet_name=name)
    for name in loadgen_xlsx.sheet_names:
        loadgen_xlsx.parse(sheet_name=name, index_col=0).to_excel(writer, sheet_name=name)
    writer.close()

    # read network from final excel file
    net = pp.from_excel(result_path)

    return net


# net = read_input('./template_loadgen.xlsx', 'template_topology.xlsx','./re.xlsx')
# print(net)
