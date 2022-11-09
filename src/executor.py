"""
Execute the toolbox.
"""
from frontend import read_input


def executor(input_setup, output_setup):
    if input_setup['used_pre_defined_net'] is '':
        net = read_input.read_input(input_setup['input_path'])
    else:
        pass

    return 0
