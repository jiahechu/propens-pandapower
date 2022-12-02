"""
script to start the toolbox.
"""

from src import executor


# summarize all basic simulation setups to a dictionary
simulation_setup = {
    'topology_path': './template_topology.xlsx',
    'used_pre_defined_net': '',
    'use_ts': False,
    'ts_path': ''
    }

# summarize all scenarios setup in a dictionary
scenarios_setup = {
    'loadgen_paths': ['./template_loadgen.xlsx'],
    'used_pre_defined_scenario': ['pv_gen'],
    'pre_defined_scenario_para': [0.9]
    }

# summarize all output setups to a dictionary
output_setup = {
    'output_path': './result',
    # the excel output will have topology and scenario name e.g. Results_Network01_Scenario01.xlsm
    'topology_name' : 'Network01', 
    'scenario_name' : 'Scenario01'
    }

# execute the toolbox
executor.executor(simulation_setup, scenarios_setup, output_setup)
