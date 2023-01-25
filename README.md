# propens-pandapower

**Basic info:** This is a toolbox using pandapower to do network simulations for different scenarios. The toolbox contains a frontend, a database (pre-defined topology and scenarios), and an analysis part.

**Developers:** Jiahe Chu, Karim Kaaniche, Martin Oviedo, Rohith Sureshbabu, Taeyoung Kim (TU Munich)


## Getting started

### Create software environment

1. Download / clone repository to local
2. Install conda environment using `conda env create -f propens-env.yaml` 
3. Activate conda environment. For PyCharm:

       - Open PyCharm
       - Go to 'File->Open'
       - Navigate to PyCharmProjects and open propens
       - When the project has opened, go to 
           `File->Settings->Project->Python Interpreter->Show all->Add->Conda Environment
            ->Existing environment->Select folder->OK`

4. Don't forget to update `propens-env.yaml` file while developing


## Create a simulation

### Create input files

1. Create topology file using `template_topology.xlsx`, which contains:
   - bus
   - line
   - switch
   - dc line
   - impedance
   - shunt
2. Create (multiple) scenarios files using `template_scenario.xlsx`, which contains:
   - motor
   - asymmetric load
   - static generator
   - asymmetric static generator
   - external grid
   - transformer
   - transformer3w
   - generator
   - storage
   - ward
   - extended ward
3. If 


### Start the program

1. Open or copy and create a new `start.py`
2. Give the path contains topology file and the name of topology as strings in input setup
3. Define scenarios in input setup as tuples. Format: `(scenario name, scenario path, name of used pre-defined scenario, pre-defined scenario parameter)`
   - Available pre-defined scenarios: _pv_gen, wind_gen, conventional_pp_gen, load, trafo_cap, line_cap, storage_
4. Give the path in which the result files should be saved as string in output setup
5. Run `start.py`, the program progresses will be showed in the terminal

## Output file
