"""
script to start the toolbox.
"""

from src import executor


# summarize all input setups to a dictionary
input_setup = {
    'loadgen_path': '',
    'topology_path': '',
    'used_scenario': '',
    'scenario_para': {},
    'used_pre_defined_net': ''}

# summarize all output setups to a dictionary
output_setup = {
        'output_path': ''}

# execute the toolbox
executor.executor(input_setup, output_setup)
