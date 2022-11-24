"""
script to start the toolbox.
"""

from src import executor


# summarize all simulation setups to a dictionary
simulation_setup = {
    'topology_path': '',
    'used_pre_defined_net': '',
    'use_ts': False,
    'ts_path': ''
    }

# summarize all scenarios setup in a dictionary
scenarios_setup = {
    'loadgen_path': ['', ''],
    'used_scenario': ['', ''],
    'scenario_para': {}
    }

# summarize all output setups to a dictionary
output_setup = {
    'output_path': ''
    }

# execute the toolbox
executor.executor(simulation_setup, output_setup, output_setup)
