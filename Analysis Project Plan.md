# Analysis – What to do?

## Analysis features

1. **Scenario set-up**
  1. Type of calculation: PF or OPF
  2. Load scaling: full loading, partial loading, loading according to areas
  3. Generation on/off or scaling based on:
    1. Technologies: Conventional, RE (+DG)
    2. Areas (countries, zones, etc)
2. **Filter**
3. **The user can select a standard to verify the operation**
  1. Set a VDE/DIE standard library with their values
4. **Allow the user to set its own values Export output into excel**
  1. Lines Loading (power and percentage)
  2. Buses Voltage (p.u.)
  3. Lines and buses that that don't comply with the standard\
5. **Export topology graphs (if there it doesn't geodata, use the build in function)**
  1. Simple graph to see the topology
  2. Power Flow graph
  3. Simple graph with lines/buses that don't comply with the standard highlighted
6. **Create charts with the output for excel** (further steps: a dashboard, with interactive bottoms)
  1. Loading vs Lines (bar chart)
  2. Voltage vs buses (bar chart)
  3. Generation vs technology (bar chart)
  4. Generation vs time (line chart)\*
7. **Make Scenarios and prepare topology – work with front/back-end team**
  1. 2 buses model (super simple)
  2. IEEE 5 bus
  3. IEEE 14 bus
  4. IEEE 30 bus
8. **Time series analysis**
  1. Configurable time steps (to be defined by the user, where the minimum time step could be based in the available data, our program should check it, and then multiple of it)
  2. Time steps results organized by elements (as can be seen in the picture you sent)

## Comments:

1. So far, the scenario set-up (I think you call it filter), will be fixed for the whole period of simulation
2. Now we have more or less what we want to accomplish, we should make the plan, organize it in sequence

## Idea for Analysis Project Plan\*

1. Create/define a base model to work with
2. Write a script that modifies the parameters in the model (e.g. set standard - OPF)
3. Run the simulation (one or several time for time series)
4. Script to extract the data to export
5. Analyze the results (check standards - PF)
6. Export the results to excel or graph format

\* This can be done in a iterative way, by starting with a simple model, then we add functionalities we follow the same process
