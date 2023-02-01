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
3. If need to do time series analysis or optimal power flow calculation, adjust the setup in `template_scenario.xlsx -general`

### Start the program

1. Open or copy and create a new `start.py`
2. Give the path contains topology file and the name of topology as strings in input setup
3. Define scenarios in input setup as tuples. Format: `(scenario name, scenario path, name of used pre-defined scenario, pre-defined scenario parameter)`
   - Available pre-defined scenarios: _pv_gen, wind_gen, conventional_pp_gen, load, trafo_cap, line_cap, storage_
4. Write the folder directory where the results and plots should be written - `output_path` , by default it is saved in propens-pandappower/results
5. Plots configuration, there are tree options for plotting the network and the results, the three of them are from pandapower plotting, in order to obtained them just write `True` for the corresponding dictionary variable. 
   - `topology` creates a simple network image, that then can be saved; 
   - the other two plots are created for each scenario in case they are `True` and if the scenario is only one iteration (no time series); 
   - these two are: `interactive map`, which provide an html file where the voltages, names, loads, generation of the buses can be seen by hovering the mouse over them, as well some branches parameters; `interactive heat map` also creates an html file, and as the name implies, it shows a heat map of the branches loading and the buses voltages.
6. Run `start.py`, the program progresses will be showed in the terminal

## Output file

The ouput file is an excel's spreadsheet with the following sheets: Summary, dashboard(if it is created), demand, buses, generation, lines, transformers, and data
In summary the buses voltages and loading percentages of the lines and transformers. In addition, by pressing the button to call the macro (should be activated), the dashboard will be created.
then, all the other sheets contains the data from all the scenarios and time steps in a tabular form, so the results can be quickly analyzed by filtering in the corresponding sheets. Or if more post-processing would be done, the last sheet, data, provide all the results (scenarios, elements, parameters) in one table/dataframe, that then can be directly be used with filter or pivots.

## Further parameters or elements
The elements that are going to be written in the output are:
 - Demand: loads
 - Buses: bus
 - Generation: gen, sgen, ext_grid
 - Lines: line
 - Transformer: trafo

If more elements are needed, they can be added in the `parameters` module in src/analysis, in the variable `element_by_type`, then all the elements will be added to the corresping sheets, in case a new type of element should be added, it is necessary to create a sheet with the same name in the `output_template` that is saved in src/analysis/output_template.
If more parameters are needed, they can be added in the `parameters` module in src/analysis, in thee variable parameters (only parameters available in the network elements can be added directly, further details are in the module description)
 
