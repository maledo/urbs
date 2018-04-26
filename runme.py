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


def scenario_nuclear_free(data):
    # shut down nuclear plants
    pro = data['process']
    nuclear_plants = (pro.index.get_level_values('Process') == 'Nuclear plant')
    pro.loc[nuclear_plants, 'inst-cap'] = 0
    pro.loc[nuclear_plants, 'cap-lo'] = 0
    pro.loc[nuclear_plants, 'cap-up'] = 0

    # allow renewable expansion    
    renewables = ((pro.index.get_level_values('Process') == 'Wind plant')
                  | (pro.index.get_level_values('Process') == 'Solar plant'))

    pro.loc[renewables, 'cap-up'] = float('inf')

    # allow cc expansion
    cc_plants = (pro.index.get_level_values('Process') == 'CC plant')
    pro.loc[cc_plants, 'cap-up'] *= 1.5

    slack = (pro.index.get_level_values('Process') == 'Slack powerplant')
    pro.loc[slack, 'inst-cap'] = 0
    pro.loc[slack, 'cap-lo'] = 0
    pro.loc[slack, 'cap-up'] = 0

    # allow bio expansion
    bio_plants = (pro.index.get_level_values('Process') == 'Biomass plant')
    pro.loc[bio_plants, 'cap-up'] *= 1.2

    # allow transmission expansion
    tra = data['transmission']
    lines = (tra.index.get_level_values('Transmission') == 'hvac')
    tra.loc[lines, 'cap-up'] = float('inf')

    return data

def scenario_co2_2030(data):
    # change global CO2 limit
    global_prop = data['global_prop']
    global_prop.loc['CO2 limit', 'value'] = 183000000

    # shut down nuclear plants
    pro = data['process']
    nuclear_plants = (pro.index.get_level_values('Process') == 'Nuclear plant')
    pro.loc[nuclear_plants, 'inst-cap'] = 0
    pro.loc[nuclear_plants, 'cap-lo'] = 0
    pro.loc[nuclear_plants, 'cap-up'] = 0

    # allow cc expansion
    cc_plants = (pro.index.get_level_values('Process') == 'CC plant')
    pro.loc[cc_plants, 'cap-up'] *= 1.5

    slack = (pro.index.get_level_values('Process') == 'Slack powerplant')
    pro.loc[slack, 'inst-cap'] = 0
    pro.loc[slack, 'cap-lo'] = 0
    pro.loc[slack, 'cap-up'] = 0

    # allow bio expansion
    bio_plants = (pro.index.get_level_values('Process') == 'Biomass plant')
    pro.loc[bio_plants, 'cap-up'] *= 1.2

    # allow renewable expansion    
    renewables = ((pro.index.get_level_values('Process') == 'Wind plant')
                  | (pro.index.get_level_values('Process') == 'Solar plant'))

    pro.loc[renewables, 'cap-up'] = float('inf')

    # allow transmission expansion
    tra = data['transmission']
    lines = (tra.index.get_level_values('Transmission') == 'hvac')
    tra.loc[lines, 'cap-up'] = float('inf')


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
                 plot_tuples=None,  plot_sites_name=None, plot_periods=None,
                 report_tuples=None, report_sites_name=None):
    """ run an urbs model for given input, time steps and scenario

    Args:
        input_file: filename to an Excel spreadsheet for urbs.read_excel
        timesteps: a list of timesteps, e.g. range(0,8761)
        scenario: a scenario function that modifies the input data dict
        result_dir: directory name for result spreadsheet and plots
        plot_tuples: (optional) list of plot tuples (c.f. urbs.result_figures)
        plot_sites_name: (optional) dict of names for sites in plot_tuples
        plot_periods: (optional) dict of plot periods(c.f. urbs.result_figures)
        report_tuples: (optional) list of (sit, com) tuples (c.f. urbs.report)
        report_sites_name: (optional) dict of names for sites in report_tuples

    Returns:
        the urbs model instance
    """

    # scenario name, read and modify data for scenario
    sce = scenario.__name__
    data = urbs.read_excel(input_file)

    # drop source lines added in Excel
    for key in data:
        data[key].drop('Source', axis=0, inplace=True, errors='ignore')
        data[key].drop('Source', axis=1, inplace=True, errors='ignore')
    data = scenario(data)

    urbs.validate_input(data)

    # create model
    prob = urbs.create_model(data, timesteps)

    # refresh time stamp string and create filename for logfile
    now = prob.created
    log_filename = os.path.join(result_dir, '{}.log').format(sce)

    # solve model and read results
    optim = SolverFactory('glpk')  # cplex, glpk, gurobi, ...
    optim = setup_solver(optim, logfile=log_filename)
    result = optim.solve(prob, tee=True)

    # save problem solution (and input data) to HDF5 file
    urbs.save(prob, os.path.join(result_dir, '{}.h5'.format(sce)))

    # write report to spreadsheet
    urbs.report(
        prob,
        os.path.join(result_dir, '{}.xlsx').format(sce),
        report_tuples=report_tuples,
        report_sites_name=report_sites_name)

    # result plots
    urbs.result_figures(
        prob,
        os.path.join(result_dir, '{}'.format(sce)),
        plot_title_prefix=sce.replace('_', ' '),
        plot_tuples=plot_tuples,
        plot_sites_name=plot_sites_name,
        periods=plot_periods,
        figure_size=(24, 9))
    return prob

if __name__ == '__main__':
    input_file = 'germany.xlsx'
    result_name = os.path.splitext(input_file)[0]  # cut away file extension
    result_dir = prepare_result_directory(result_name)  # name + time stamp

    # copy input file to result directory
    shutil.copyfile(input_file, os.path.join(result_dir, input_file))
    # copy runme.py to result directory
    shutil.copy(__file__, result_dir)

    # simulation timesteps
    (offset, length) = (0, 24)  # time step selection
    timesteps = range(offset, offset+length+1)

    # plotting commodities/sites
    plot_tuples = [
        # ('Baden-Württemberg', 'Elec'),
        # ('Bavaria', 'Elec'),
        # ('Berlin', 'Elec'),
        # ('Brandenburg', 'Elec'),
        # ('Bremen', 'Elec'),
        # ('Hamburg', 'Elec'),
        # ('Hesse', 'Elec'),
        # ('Lower Saxony', 'Elec'),
        # ('Mecklenburg-Vorpommern', 'Elec'),
        # ('North Rhine-Westphalia', 'Elec'),
        # ('Rhineland-Palatinate', 'Elec'),
        # ('Saarland', 'Elec'),
        # ('Saxony', 'Elec'),
        # ('Saxony-Anhalt', 'Elec'),
        # ('Schleswig-Holstein', 'Elec'),
        # ('Thuringia', 'Elec'),
        (['Baden-Württemberg', 'Bavaria', 'Berlin', 'Brandenburg',
          'Bremen', 'Hamburg', 'Hesse', 'Lower Saxony', 'Mecklenburg-Vorpommern',
          'North Rhine-Westphalia', 'Rhineland-Palatinate', 'Saarland', 'Saxony',
          'Saxony-Anhalt', 'Schleswig-Holstein', 'Thuringia', 'Offshore'], 'Elec')]

    # optional: define names for plot_tuples
    plot_sites_name = {
        ('Baden-Württemberg', 'Bavaria', 'Berlin', 'Brandenburg', 'Bremen',
         'Hamburg', 'Hesse', 'Lower Saxony', 'Mecklenburg-Vorpommern', 'North Rhine-Westphalia',
         'Rhineland-Palatinate', 'Saarland', 'Saxony', 'Saxony-Anhalt', 'Schleswig-Holstein',
         'Thuringia', 'Offshore'): 'Germany'}

    # detailed reporting commodity/sites
    report_tuples = [
        ('Baden-Württemberg', 'Elec'),
        ('Bavaria', 'Elec'),
        ('Berlin', 'Elec'),
        ('Brandenburg', 'Elec'),
        ('Bremen', 'Elec'),
        ('Hamburg', 'Elec'),
        ('Hesse', 'Elec'),
        ('Lower Saxony', 'Elec'),
        ('Mecklenburg-Vorpommern', 'Elec'),
        ('North Rhine-Westphalia', 'Elec'),
        ('Rhineland-Palatinate', 'Elec'),
        ('Saarland', 'Elec'),
        ('Saxony', 'Elec'),
        ('Saxony-Anhalt', 'Elec'),
        ('Schleswig-Holstein', 'Elec'),
        ('Thuringia', 'Elec'),
        ('Offshore', 'Elec'),
        (['Baden-Württemberg', 'Bavaria', 'Berlin', 'Brandenburg',
          'Bremen', 'Hamburg', 'Hesse', 'Lower Saxony', 'Mecklenburg-Vorpommern',
          'North Rhine-Westphalia', 'Rhineland-Palatinate', 'Saarland', 'Saxony',
          'Saxony-Anhalt', 'Schleswig-Holstein', 'Thuringia', 'Offshore'], 'Elec'),
        ('Baden-Württemberg', 'CO2'),
        ('Bavaria', 'CO2'),
        ('Berlin', 'CO2'),
        ('Brandenburg', 'CO2'),
        ('Bremen', 'CO2'),
        ('Hamburg', 'CO2'),
        ('Hesse', 'CO2'),
        ('Lower Saxony', 'CO2'),
        ('Mecklenburg-Vorpommern', 'CO2'),
        ('North Rhine-Westphalia', 'CO2'),
        ('Rhineland-Palatinate', 'CO2'),
        ('Saarland', 'CO2'),
        ('Saxony', 'CO2'),
        ('Saxony-Anhalt', 'CO2'),
        ('Schleswig-Holstein', 'CO2'),
        ('Thuringia', 'CO2'),
        ('Offshore', 'CO2'),
        (['Baden-Württemberg', 'Bavaria', 'Berlin', 'Brandenburg',
          'Bremen', 'Hamburg', 'Hesse', 'Lower Saxony', 'Mecklenburg-Vorpommern',
          'North Rhine-Westphalia', 'Rhineland-Palatinate', 'Saarland', 'Saxony',
          'Saxony-Anhalt', 'Schleswig-Holstein', 'Thuringia', 'Offshore'], 'CO2')]

    # optional: define names for report_tuples
    report_sites_name = {
        ('Baden-Württemberg', 'Bavaria', 'Berlin', 'Brandenburg', 'Bremen',
         'Hamburg', 'Hesse', 'Lower Saxony', 'Mecklenburg-Vorpommern', 'North Rhine-Westphalia',
         'Rhineland-Palatinate', 'Saarland', 'Saxony', 'Saxony-Anhalt', 'Schleswig-Holstein',
         'Thuringia', 'Offshore'): 'Germany'}

    # plotting timesteps
    plot_periods = {
        'all': timesteps[1:]
    }

    # add or change plot colors
    my_colors = {}
    for country, color in my_colors.items():
        urbs.COLORS[country] = color

    # select scenarios to be run
    scenarios = [
        scenario_base,
        scenario_nuclear_free,
        scenario_co2_2030
    ]

    for scenario in scenarios:
        prob = run_scenario(input_file, timesteps, scenario, result_dir,
                            plot_tuples=plot_tuples,
                            plot_sites_name=plot_sites_name,
                            plot_periods=plot_periods,
                            report_tuples=report_tuples,
                            report_sites_name=report_sites_name)
