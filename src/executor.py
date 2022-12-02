"""
Execute the toolbox.
"""
import pandapower as pp
import read_input
import apply_scenario
from output import create_excel


def executor(simulation_setup, scenarios_setup, output_setup):
    for scenario_path, pd_scenario, pd_para in zip(scenarios_setup['loadgen_paths'], scenarios_setup['used_pre_defined_scenario'], scenarios_setup['pre_defined_scenario_para']):
        # create pandapower network from user-defined Excel file or pre-defined topology
        if simulation_setup['used_pre_defined_net'] is '':
            net = read_input.read_input(scenario_path, simulation_setup['topology_path'], output_setup['output_path'])
        else:
            pass

        # apply scenario from data
        if pd_scenario != '':
            net = apply_scenario.apply_scenario(net, pd_scenario, pd_para)


    print(net)
    
    ''' As we discussed before, so far the "net" that is used below this line should include the results (as list, biig df, etc) '''
    # list_of_net = [net_step1, net_step2, ... ] 
    
    network_name = output_setup['topology_name']
    scenario_name = output_setup['scenario_name']
    gen_fuel_tech = []
    output_path = output_setup['output_path']
    
    create_excel(network_name, scenario_name, list_of_net, gen_fuel_tech, output_path)
    

    return 0
