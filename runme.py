import os
import pandas as pd
import pyomo.environ
import shutil
import urbs
from datetime import datetime
from pyomo.opt.base import SolverFactory


# SCENARIOS
def scenario_base(data):
    # do nothing
    return data
   
def scenario_pv(data):
    # allow renewable expansion
    pro = data['process']
    nuclear_plants = (pro.index.get_level_values('Process') == 'Nuclear plant')
    pro.loc[nuclear_plants, 'inst-cap'] = 0
    pro.loc[nuclear_plants, 'cap-lo'] = 0
    pro.loc[nuclear_plants, 'cap-up'] = 0
    
    pv = (pro.index.get_level_values('Process') == 'Photovoltaics')
                
    pro.loc[pv, 'cap-up'] = float('inf')
    
    # allow transmission expansion
    tra = data['transmission']
    lines = (tra.index.get_level_values('Transmission') == 'hvac')
    tra.loc[lines, 'cap-up'] = float('inf')
    
    # allow storage expansion
    sto = data['storage']
    battery = (sto.index.get_level_values('Storage') == ('Battery'))
    sto.loc[battery, 'cap-up-c'] = float('inf')
    sto.loc[battery, 'cap-up-p'] = float('inf')
    return data
    
def scenario_wind(data):
    # allow renewable expansion
    pro = data['process']
    nuclear_plants = (pro.index.get_level_values('Process') == 'Nuclear plant')
    pro.loc[nuclear_plants, 'inst-cap'] = 0
    pro.loc[nuclear_plants, 'cap-lo'] = 0
    pro.loc[nuclear_plants, 'cap-up'] = 0
    
    renewables = ((pro.index.get_level_values('Process') == 'Wind park')
                | (pro.index.get_level_values('Process') == 'Photovoltaics'))
                
    pro.loc[renewables, 'cap-up'] = float('inf')
    
    # allow transmission expansion
    tra = data['transmission']
    lines = (tra.index.get_level_values('Transmission') == 'hvac')
    tra.loc[lines, 'cap-up'] = float('inf')
    
    # allow storage expansion
    sto = data['storage']
    battery = (sto.index.get_level_values('Storage') == ('Battery'))
    sto.loc[battery, 'cap-up-c'] = float('inf')
    sto.loc[battery, 'cap-up-p'] = float('inf')
    return data

def scenario_biomass(data):
    # allow renewable expansion
    pro = data['process']
    nuclear_plants = (pro.index.get_level_values('Process') == 'Nuclear plant')
    pro.loc[nuclear_plants, 'inst-cap'] = 0
    pro.loc[nuclear_plants, 'cap-lo'] = 0
    pro.loc[nuclear_plants, 'cap-up'] = 0
    
    renewables = ((pro.index.get_level_values('Process') == 'Wind park')
                | (pro.index.get_level_values('Process') == 'Photovoltaics'))
    biomass = (pro.index.get_level_values('Process') == 'Biomass plant')
                
    pro.loc[renewables, 'cap-up'] = float('inf')
    pro.loc[biomass, 'cap-up'] = pro.loc[biomass, 'inst-cap']*1.5
    
    # allow transmission expansion
    tra = data['transmission']
    lines = (tra.index.get_level_values('Transmission') == 'hvac')
    tra.loc[lines, 'cap-up'] = float('inf')
    
    # allow storage expansion
    sto = data['storage']
    battery = (sto.index.get_level_values('Storage') == ('Battery'))
    sto.loc[battery, 'cap-up-c'] = float('inf')
    sto.loc[battery, 'cap-up-p'] = float('inf')
    return data
    

    
def prepare_result_directory(result_name):
    """ create a time stamped directory within the result folder """
    # timestamp for result directory
    now = datetime.now().strftime('%Y%m%dT%H%M')

    # create result directory if not existent
    result_dir = os.path.join('result', '{}-{}'.format(result_name, now))
    if not os.path.exists(result_dir):
        os.makedirs(result_dir)

    return result_dir


def setup_solver(optim, logfile='solver.log'):
    """ """
    if optim.name == 'gurobi':
        # reference with list of option names
        # http://www.gurobi.com/documentation/5.6/reference-manual/parameters
        optim.set_options("logfile={}".format(logfile))
        # optim.set_options("timelimit=7200")  # seconds
        # optim.set_options("mipgap=5e-4")  # default = 1e-4
    elif optim.name == 'glpk':
        # reference with list of options
        # execute 'glpsol --help'
        optim.set_options("log={}".format(logfile))
        # optim.set_options("tmlim=7200")  # seconds
        # optim.set_options("mipgap=.0005")
    else:
        print("Warning from setup_solver: no options set for solver "
              "'{}'!".format(optim.name))
    return optim


def run_scenario(input_file, timesteps, scenario, result_dir,
                 plot_tuples=None, plot_periods=None, report_tuples=None):
    """ run an urbs model for given input, time steps and scenario

    Args:
        input_file: filename to an Excel spreadsheet for urbs.read_excel
        timesteps: a list of timesteps, e.g. range(0,8761)
        scenario: a scenario function that modifies the input data dict
        result_dir: directory name for result spreadsheet and plots
        plot_tuples: (optional) list of plot tuples (c.f. urbs.result_figures)
        plot_periods: (optional) dict of plot periods (c.f. urbs.result_figures)
        report_tuples: (optional) list of (sit, com) tuples (c.f. urbs.report)

    Returns:
        the urbs model instance
    """

    # scenario name, read and modify data for scenario
    sce = scenario.__name__
    data = urbs.read_excel(input_file)
    # drop Source lines added in Excel
    for key in data:
        data[key].drop('Source', axis=0, inplace=True, errors='ignore')
    data = scenario(data)

    # create model
    prob = urbs.create_model(data, timesteps)
    #prob.write('model.lp', io_options={'symbolic_solver_labels': True})

    # refresh time stamp string and create filename for logfile
    now = prob.created
    log_filename = os.path.join(result_dir, '{}.log').format(sce)

    # solve model and read results
    optim = SolverFactory('gurobi')  # cplex, glpk, gurobi, ...
    optim = setup_solver(optim, logfile=log_filename)
    result = optim.solve(prob, tee=True)

    # copy input file to result directory
    shutil.copyfile(input_file, os.path.join(result_dir, input_file))
    
    # save problem solution (and input data) to HDF5 file
    urbs.save(prob, os.path.join(result_dir, '{}.h5'.format(sce)))

    # write report to spreadsheet
    urbs.report(
        prob,
        os.path.join(result_dir, '{}.xlsx').format(sce),
        report_tuples=report_tuples)

    # result plots
    urbs.result_figures(
        prob,
        os.path.join(result_dir, '{}'.format(sce)),
        plot_title_prefix=sce.replace('_', ' '),
        plot_tuples=plot_tuples,
        periods=plot_periods,
        figure_size=(24, 9),
        extensions=['png'])
    return prob

if __name__ == '__main__':
    input_file = 'bavaria.xlsx'
    result_name = os.path.splitext(input_file)[0]  # cut away file extension
    result_dir = prepare_result_directory(result_name)  # name + time stamp

    # simulation timesteps
    (offset, length) = (0, 24*365) # time step selection
    timesteps = range(offset, offset+length+1)

    # plotting commodities/sites
    plot_tuples = [
        ('Mittelfranken', 'Elec'),
        ('Niederbayern', 'Elec'),
        ('Oberbayern', 'Elec'),
        ('Oberfranken', 'Elec'),
        ('Oberpfalz', 'Elec'),
        ('Schwaben', 'Elec'),
        ('Unterfranken', 'Elec'),
        (['Mittelfranken', 'Niederbayern',
        'Oberbayern', 'Oberfranken',
        'Oberpfalz', 'Schwaben',
        'Unterfranken'], 'Elec')]

    # detailed reporting commodity/sites
    report_tuples = [
        ('Mittelfranken', 'Elec'),
        ('Niederbayern', 'Elec'),
        ('Oberbayern', 'Elec'),
        ('Oberfranken', 'Elec'),
        ('Oberpfalz', 'Elec'),
        ('Schwaben', 'Elec'),
        ('Unterfranken', 'Elec'),
        ('Mittelfranken', 'CO2'),
        ('Niederbayern', 'CO2'),
        ('Oberbayern', 'CO2'),
        ('Oberfranken', 'CO2'),
        ('Oberpfalz', 'CO2'),
        ('Schwaben', 'CO2'),
        ('Unterfranken', 'CO2')]

    # plotting timesteps
    plot_periods = {
        'all': timesteps[1:],
        'spring': timesteps[1417:1585],
        'summer': timesteps[3625:3793],
        'autumn': timesteps[5833:6001],
        'winter': timesteps[8017:8185]
    }

    # add or change plot colors
    my_colors = {
        'Unterfranken': (230, 200, 200),
        'Mittelfranken': (200, 230, 200),
        'Oberfranken': (200, 200, 230),
        'Oberpfalz': (130, 200, 200),
        'Niederbayern': (200, 130, 200),
        'Oberbayern': (200, 200, 130),
        'Schwaben': (230, 100, 100),
        'South': (230, 200, 200),
        'Mid': (200, 230, 200),
        'North': (200, 200, 230)}
    for country, color in my_colors.items():
        urbs.COLORS[country] = color

    # select scenarios to be run
    scenarios = [
        scenario_base,
        scenario_pv,
        scenario_wind,
        scenario_biomass
        ]

    for scenario in scenarios:
        prob = run_scenario(input_file, timesteps, scenario, result_dir,
                            plot_tuples=plot_tuples,
                            plot_periods=plot_periods,
                            report_tuples=report_tuples)
