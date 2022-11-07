"""
Read network data from excel file.
"""

import pandas as pd
import pandapower as pp


def read_input(file_path):
    """
    Read network data from excel file.

    Args:
        file_path: path of the excel file.

    Returns:
        net: pandapower network.
    """
    net = pp.create_empty_network()
    return net
