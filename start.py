"""
script to start the toolbox.
"""

from src import executor


# summarize all input setups to a dictionary
input_setup = {
    'loadgen_path': '',
    'topology_path': '',
    'output_path': '',
    'used_scenario': '',
    'used_pre_defined_net': ''}

# summarize all output setups to a dictionary
output_setup = {}

# execute the toolbox
executor.executor(input_setup, output_setup)

