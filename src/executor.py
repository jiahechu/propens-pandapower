"""
Execute the toolbox.
"""
from src.frontend.read_input import read_input
from src.frontend.generate_timeseries import generate_timeseries
from src.scenarios.apply_scenario import apply_scenario
from src.analysis.solver import solve
from src.analysis.excel_output import create_excel
from src.analysis.save import save_results
from src.analysis.parameters import preallocate_tables
from datetime import datetime


def executor(input_setup, output_setup):
    """
    Execute the toolbox.

    Args:
        input_setup: dictionary contains input setups.
        output_setup: dictionary contains output setups.

    Returns:
        temporary files: dictionary that contains the temporary files for time series 
        tables: dictionary of dataframes, which are the tables to be written in the excel output
    """
    time_start = datetime.now()
    print('\n Running: read_conventional_generation')    # preallocate tables and temporary files dictionary for the different scenarios results
    tables = preallocate_tables(input_setup)
    temporary_files = {} 
    
    # simulate for each scenario
    for scenario_name, scenario_path, pd_scenario, pd_para in input_setup['scenario_setup']:
        print('\nStart simulation with scenario ' + scenario_name)

        """Frontend"""
        # create pandapower network from Excel
        try:
            print('\nReading input excel files')
            net, general, gen_fuel_tech = read_input(scenario_path, input_setup['topology_path'])
            time_steps = 1
            print('> Done')
        except:
            print('\nError while reading excel files, please check input file again!')
            print('Program stops.')
            print('Detail error arguments: ')
            raise

        # apply scenario from data
        if pd_scenario != '':
            try:
                print('\nApplying pre-defined scenarios')
                net = apply_scenario(net, pd_scenario, pd_para)
                print('> Done')
            except:
                print('\nError while applying pre-defined scenarios, please check input parameters again!')
                print('Program stops.')
                print('Detail error arguments: ')
                raise

        # apply time series
        if general['use_ts'][0]:
            try:
                print('\nGenerating controllers for time series analysis')
                net, time_steps = generate_timeseries(net, general['ts_path'][0])
                print('> Done')
            except:
                print('\nError while generating controllers for time series analysis, please check time series data again!')
                print('Program stops.')
                print('Detail error arguments: ')
                raise

        print('\nSuccessfully read pandapower network from Excel:')
        print(net)

        """Analysis"""
        # temporary files are the results from the time series analysis,
        # in case of one iteration, the results are in net.res_####
        try:
            temporary_files[scenario_name], net = solve(input_setup['topology_name'], scenario_name, gen_fuel_tech, 
                                                        output_setup['output_path'], net, time_steps, general)
        except:
            print('\nError while solving network, e.g. not converging')
            print('Program stops.')
            print('Detail error arguments: ')
            raise

        # tables contains the tables that are going to be written into the excel output,
        # they are saved in a dict according to the calculated scenarios
        try:
            tables[scenario_name] = save_results(net, gen_fuel_tech, scenario_name, time_steps, temporary_files[scenario_name])    
        except:
            print('\nError while saving the results')
            print('Program stops.')
            print('Detail error arguments: ')
            raise

    # Call the excel template, fill up with the results from all scenarios
    try:
        create_excel(input_setup['topology_name'], output_setup['output_path'], tables)
    except:
        print('\nError while creating output excel file')
        print('Program stops.')
        print('Detail error arguments: ')
        raise
    time_end = datetime.now()
    td = (time_end - time_start).total_seconds()
    print('\n ------  The total time to calculate all scenarios was ' + str(td) + ' seconds ------')
    
    return 