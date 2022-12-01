"""
script to start the toolbox.
"""

from src import executor

# summarize all input setup in a dictionary
# available pre-defined scenarios: pv_gen, wind_gen, conventional_pp_gen, load, trafo_cap, line_cap, storage, switch
input_setup = {     # (scenario path, name of used pre-defined scenario, ggf. pre-defined scenario parameter)
    'topology_path': './example/kerber_landnetz_fl2/topology.xlsx',
    'scenario_setup': [('./example/kerber_landnetz_fl2/scenarios/basic.xlsx', 'load', 0.9),
                       ('./example/kerber_landnetz_fl2/scenarios/pv.xlsx', 'pv_gen', 1.5),
                       ('./example/kerber_landnetz_fl2/scenarios/pv_storage.xlsx', '', 0)]
    }

# summarize all output setups to a dictionary
output_setup = {
    'output_path': './result'
    }

# execute the toolbox
executor.executor(input_setup, output_setup)
