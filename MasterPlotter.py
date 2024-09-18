#  Import List
import matplotlib.pyplot as plt
from itertools import cycle
import pandas as pd
import numpy as np
import os
import re
import warnings

#
warnings.simplefilter("ignore", category=UserWarning)
pd.options.mode.chained_assignment = None

# Declaring Variables
filename, excel_data, cell_no = None, None, None  # Electrode_Weight, Active_Material_Weight, Excel File Name, cell number variables, actual cell number
FMN, RPF, CYC = None, None, None  # Formation, Rate Profile, Long Cycle
AC, Fresh, AFFM, AFCY, cell_state, title = None, None, None, None, None, None  # AC File, Fresh cell, After Formation, After Long Cycle


# functions #

# Getting the file input from user
def file():
    global filename
    # raw_path = input("Enter the path of the file: ")
    raw_path = r'"C:\Users\Mano-BRCGE\Desktop\NeWare Data\20240905-Oven-Com-LiDFP\240905-C01-CNCMA90-LiDFOB+LiDFP-0.01144-0.1C.xlsx"'
    path = raw_path.strip(' " " ')
    filename = str(os.path.basename(path).split('/'))

    file.path = path

    csv, CSV, txt, xlsx = ".csv" in file.path, ".CSV" in file.path, ".txt" in file.path, ".xlsx" in file.path
    if txt or csv or CSV or xlsx:
        pass
    else:
        file.path += ".csv"
    return

# A function to get the cell number from the file name to be used in output image file name.
def get_CellNo():
    global cell_no
    pattern = re.compile(r'C0[1-8]')
    match = pattern.search(filename)
    cell_no = match.group() if match else None
    return

# Finds the state of cell when the impedance was measured (Whether Fresh, After Formation or After Cycling)
def cell_state():
    global cell_state, title
    if 'AFFM' in filename:
        cell_state = 'After Formation'
    elif 'Fresh' in filename:
        cell_state = 'Fresh'
    elif 'AFCY' in filename:
        cell_state = 'After cycling'
    else:
        cell_state = None

# Creating dataframes for each Excel sheet
def parse_excel(path, sheets):
    global excel_data
    excel_data = pd.ExcelFile(path)
    dataframes = {}
    for i in sheets:
        try:
            dataframes[i] = pd.read_excel(path, engine="openpyxl", sheet_name=i)
        except ValueError:
            print(f"Warning: Sheet {i} not found in the file.")
    return dataframes

# Removing the unnecessary zeros from ACTION and ACCMAH
def rest_remover(src_df):
    search_rest = ['Rest']
    only_rest = (src_df['Step Type'].isin(search_rest))
    src_df = src_df.loc[~only_rest]
    return src_df

# Converting capacity from Ah to mAh and add it to a new column named 'Chg. Cap.(mAH)' and 'DChg. Cap.(mAh)'
def conv_Ah_to_mAh(df):
    df['Chg. Cap.(mAh)'] = (df['Chg. Cap.(Ah)'] / 1000).__round__(4)
    df['DChg. Cap.(mAh)'] = (df['DChg. Cap.(Ah)'] / 1000).__round__(4)
    return df

# Function to filter and select specific columns
def filter_trim(df, filters, XY):
    """
    Filters a dataframe based on specified conditions and selects certain columns.

    Parameters:
    - df (DataFrame): The original dataframe to filter.
    - filters (dict): A dictionary where keys are column names and values are the filter values.
    - XY (list): A list of columns to select from the filtered data.

    Returns:
    - DataFrame: The filtered and selected dataframe.
    """
    # Apply filters based on the provided conditions
    filtered_df = df.copy()  # Copy the original DataFrame to avoid modifying it
    for column, value in filters.items():
        filtered_df = filtered_df[filtered_df[column] == value]

    # Select the specified columns
    selected_df = filtered_df[XY]

    return selected_df

# Getting plot choice form user
def user_choice():
    global FMN, RPF, CYC, iV, CE, AC
    FMN, RPF, CYC, iV, CE, AC = "FMN" in filename, "RPF" in filename, "CYC" in filename, "iV" in filename, "CE" in filename, "AC" in filename
    if FMN or RPF or CYC or iV or CE or AC:
        pass
    else:
        plt_choice = int(input("1. Formation Cycles Plot \n"
                               "or\n"
                               "2. Rate Profile Plot \n"
                               "or\n"
                               "3. Long Cycles Plot \n"
                               "Please choose 1 or 2 or 3>>: "))
        if plt_choice == 2:
            rpf_choice = int(input("4. Voltage Profile \n"
                                   "or\n"
                                   "5. Step Profile\n"
                                   "Please choose a or b>>: "))
            plt_choice = rpf_choice
            return plt_choice
        return plt_choice
    return FMN, RPF, CYC, iV, CE

# Finds what kind of data is being given and creates bool values to respective variables.
def plot_find():
    global FMN, RPF, CYC, iV, CE, AC, Fresh, AFFM, AFCY
    FMN, RPF, CYC, iV, CE = "FMN" in filename, "RPF" in filename, "CYC" in filename, "iV" in filename, "CE" in filename
    AC, Fresh, AFFM, AFCY = "AC" in filename, "Fresh" in filename, "AFFM" in filename, "AFCY" in filename
    return FMN, RPF, CYC, iV, CE, AC, Fresh, AFFM, AFCY

# Plot function for formation plots
def fmn_plotter(src_df, x, y):
    cy_bool = (src_df[' CYCLE'].isin([x]))
    cy = src_df.loc[cy_bool]
    st_bool = (cy[' STEP'].isin([y]))
    step = cy.loc[st_bool]
    plt.plot(step[' ACCMAHG'], step[' VOLTS'],
             color=next(colors),
             label=next(labels))
    return

# Plot function for rate profile gcd plots
def rpf_plotter(src_df, x):
    trend_file_check()
    cycle_bool = (src_df[' CYCLE'].isin([1]))
    cycle = src_df.loc[cycle_bool]
    lp_bool = (cycle[' LOOP'].isin([2]))
    loop_2 = cycle.loc[lp_bool]
    st_bool = (loop_2[' STEP'].isin([x]))
    step = loop_2.loc[st_bool]

    plt.plot(step[' ACCMAHG'], step[' VOLTS'],
             color=next(colors),
             label=next(labels))
    return

# Plot function for rate profile step plots
def rpf_step_plotter(src_df):
    CE_file_check()

    global a_m_w

    src_df['Acc mAHg'] = (src_df['Acc mAH'] / a_m_w).round(2)
    dschg_bool = (src_df['Action'].isin(['Discharge']))
    dschg = src_df.loc[dschg_bool]
    cycle_bool = (dschg['Cycle'].isin([1]))
    cycle_df = dschg.loc[cycle_bool]
    plt.plot(cycle_df['StepID'], cycle_df['Acc mAHg'],
             marker='s',
             color='k',
             linewidth=3)
    return

# Plot function for cycle life plots
def cyc_plotter(src_df):
    CE_file_check()

    global a_m_w
    src_df['Acc mAHg'] = (src_df['Acc mAH'] / a_m_w).__round__(2)

    chg_bool = (src_df['Action'].isin(['Charge']))
    chg = src_df.loc[chg_bool]
    dischg = src_df.loc[~chg_bool]
    CE_df = pd.DataFrame({'Cycle': chg['Cycle']})
    CE_df['Chg mAH'] = chg['Acc mAH'].to_numpy()
    CE_df['Dischg mAH'] = dischg['Acc mAH'].to_numpy()
    CE_df['CE'] = (CE_df['Dischg mAH'] / CE_df['Chg mAH'] * 100).round(2)

    fig, ax1 = plt.subplots()
    ax1.set_xlabel("Cycles")
    ax1.set_ylabel("Specific discharge capacity (mA h $g^{-1}$)", color='blue')
    cap_plot, = ax1.plot(dischg['Cycle'], dischg['Acc mAHg'],
                         label='Discharge Capacity',
                         color='blue',
                         linewidth=3,
                         marker='o')
    ax1.tick_params(axis='y', labelcolor='blue')
    ax1.set_ylim(0, 200)
    # ax1.set_xlim(0, 200)

    ax2 = ax1.twinx()
    ax2.set_ylabel("Coulombic Efficiency (%)", color='red')
    CE_plot, = ax2.plot(CE_df['Cycle'], CE_df['CE'],
                        label='CE',
                        color='red',
                        linewidth=3,
                        marker='*')
    ax2.tick_params(axis='y', labelcolor='red')
    ax2.set_ylim(0, 110)
    plt.legend([cap_plot, CE_plot], ["Discharge Capacity", "CE"])
    return

# Plot function for long cycle GCD plots
def cyc_gcd_plotter(src_df, x, y):
    trend_file_check()

    cy_bool = (src_df[' CYCLE'].isin([x]))
    cy = src_df.loc[cy_bool]
    st_bool = (cy[' STEP'].isin([y]))
    step = cy.loc[st_bool]
    plt.plot(step[' ACCMAHG'], step[' VOLTS'])
    return

# Plot function for AC-Impedance plots
def acimp_plotter(raw_df):
    # Extract Z' (Ω) and -Z'' (Ω) columns
    z_prime = raw_df['Z\' (Ω)']
    minus_z_double_prime = raw_df['-Z\'\' (Ω)']

    # Set up the scatter plot
    plt.scatter(z_prime, minus_z_double_prime, marker='o', s=7, color='blue')
    # Set labels and title
    plt.xlabel('Z\' (Ω)')
    plt.ylabel('-Z\'\' (Ω)')





# Find number of cycles
def find_cycles(filename):
    match = re.search(r'@(\d+)', filename)
    if match:
        return int(match.group(1))
    else:
        return None

# Add the suffix to the plot title if a corresponding keyword is present
    if cell_state:
        title = f"{cell_no}_{e_w} g_{cell_state}"
    else:
        title = f"{cell_no}_{e_w} g"
    return title, cell_state


#################################### Run of Events Starts Here ##########################################
# get_file()
# cellno()
# date = filename[2:8]


# Finding out user choice
# auto_find = user_choice()
#
# Loading file as a Pandas DataFrame
# if CE or iV:
#     raw_df = pd.read_csv(get_file.csv_file)
#     if CYC:
#         cyc_no = find_cycles(filename)
#     else:
#         pass
# elif AC:
#     raw_df = pd.read_csv(get_file.csv_file, delimiter=';')


# if FMN or auto_find == 1:
#     cap_df = cap_calc(raw_df)
#     plot_ready = zero_remover(cap_df)
#     cycle_x = [1, 2, 3]
#     step_x = [1, 2]
#     colors = cycle(['k', 'k', 'r', 'r', 'b', 'b'])
#     labels = cycle(['Cycle 1', '', 'Cycle 2', '', 'Cycle 3', ''])
#     for a in cycle_x:
#         for b in step_x:
#             fmn_plotter(plot_ready, a, b)
#     # plt.autoscale(enable=True, axis='x')
#     #plt.xlim([0, 220])
#     plt.xlim(left=0)
#     plt.ylim([2.7, 4.4])
#     plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
#     plt.ylabel("Voltage vs. Li/Li$^+$ (V)")
#     plt.title(f"0.1 C Formation Cycles - {cellno} - {e_w} g")
#     plt.legend()
#     plt.savefig(f'{date}_{cellno}_FMN_GCD_Plot_{e_w}g.png',
#                 transparent=True,
#                 dpi=1000)

# elif RPF or auto_find == 4 or auto_find == 5:
#     if iV or auto_find == 4:
#         cap_df = cap_calc(raw_df)
#         plot_ready = zero_remover(cap_df)
#         step_y = [1, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 21]
#         colors = cycle(['k', 'k', 'r', 'r', 'b', 'b', 'g', 'g', 'm', 'm', 'c', 'c'])
#         labels = cycle(['', '0.2 C', '', '0.5 C', '', '1 C', '', '5 C', '', '10 C', '', '0.2 C'])
#         for i in step_y:
#             rpf_plotter(plot_ready, i)
#
#         plt.title(f"0.2 C - 10 C - 0.2 C Rate Profile  - {cellno} - {e_w} g")
#         plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
#         plt.ylabel("Voltage vs. Li/Li$^+$ / (V)")
#         plt.xlim([0, 200])
#         plt.ylim([2.7, 4.4])
#         plt.legend()
#         plt.savefig(f'{date}_{cellno}_RPF_GCD_Plot_{e_w}g.png',
#                     transparent=True,
#                     dpi=1000)
#     elif CE or auto_find == 5:
#         rpf_step_plotter(raw_df)
#         plt.title(f"0.2 C - 10 C - 0.2 C Step Profile  - {cellno} - {e_w} g")
#         plt.xlabel("C-Rates")
#         cus_tick_pos = [4, 12, 21, 30, 39, 46]
#         cus_tick_labels = ['0.2 C', '0.5 C', '1 C', '5 C', '10 C', '0.2 C']
#         plt.xticks(cus_tick_pos, cus_tick_labels)
#         plt.ylabel("Specific Discharge Capacity (mA h $g^{-1}$)")
#         # plt.autoscale(enable=True, axis='y')
#         plt.ylim([50, 220])
#         # plt.ylim(bottom=0)
#         plt.savefig(f'{date}_{cellno}_RPF_Step_Plot_{e_w}g.png',
#                 transparent=True,
#                 dpi=1000)

# elif CYC or auto_find == 3:
#     cyc_no = find_cycles(filename)
#     if CE:
#         cyc_plotter(raw_df)
#         plt.title(f"Long Cycling - {cellno} - {e_w}g - {cyc_no} cycles")
#         plt.savefig(f'{date}_{cellno}_Long_CYC_Plot_{e_w}g_{cyc_no}_cycles.png',
#                     transparent=True,
#                     dpi=1000)
#     elif iV:
# print("Long cycle iV file found")
#         cap_df = cap_calc(raw_df)
#         plot_ready = zero_remover(cap_df)
#         cycle_x = [1, 10, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200]
#         step_x = [1, 2]
#         num_of_plots = len(cycle_x)
#         colors = cycle(plt.cm.viridis(np.linspace(0, 1, num_of_plots)))
#
#         for a in cycle_x:
#             for b in step_x:
#                 cyc_gcd_plotter(plot_ready, a, b)
#             color = next(colors)
#             # label = next(cycle_x)
#         plt.autoscale(enable=True, axis='x')
#         # plt.xlim([0, 220])
#         plt.ylim([2.7, 4.4])
#         plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
#         plt.ylabel("Voltage vs. Li/Li$^+$ (V)")
#         plt.title(f'Long Cycling  - {cellno} - {e_w}g - {cyc_no} cycles')
#         plt.legend()
#         plt.savefig(f'{date}_{cellno}_Long_CYC_GCD_Plot_{e_w}g_{cyc_no}_cycles.png',
#                     transparent=True,
#                     dpi=1000)
#
# elif AC:
#     acimp_plotter(raw_df)
#     cell_state()
#     # Set equal scaling for both axes based on the last data point
#     max_value = max(raw_df['Z\' (Ω)'].max(), raw_df['-Z\'\' (Ω)'].max())
#     # plt.axis('equal')
#     plt.title(title)
#     plt.xlim(-10, max_value+50)
#     plt.ylim(-10, max_value+50)
#     plt.savefig(f'{date}_{cellno}_AC Impedance_Plot_{e_w}g_{cell_state}.png',
#                 transparent=True,
#                 dpi=1000)
#     plt.xlim(-10, 1000)
#     plt.ylim(-10, 1000)
#     plt.savefig(f'{date}_{cellno}_AC Impedance_Plot_{e_w}g_{cell_state}_1K.png',
#                 transparent=True,
#                 dpi=1000)
#################################### Run of Events Ends Here ##########################################
print("Started...")
#################################### Testing Bay ##########################################
file()  # Gets the file names and file paths
print(file.path)
print(filename)
sheets4plot = ['cycle', 'step', 'record']  # the actual sheets I'm interested in the Excel file
work_frames = parse_excel(file.path, sheets4plot)  # Trims the dataframe to only the selected sheets
#Converting each sheet into a separate dataframes
for sheet_name, df in work_frames.items():
    globals()[sheet_name] = df  # Creates a global variable with the name of the sheet

gcd_record = rest_remover(record)
print(f"Dataframe RECORD is {gcd_record}")

# Create an empty dataframe to store results
plot_data = pd.DataFrame()

# Get the unique values for 'Cycle Index' and 'Step Type'
cycle_values = gcd_record['Cycle Index'].unique()
step_values = ['CC Chg', 'CC DChg']  # Since you have these two step types

# Nested loop: first for 'Cycle Index', then for 'Step Type'
for cycle in cycle_values:
    for step in step_values:
        # Define the filters
        filters = {
            'Cycle Index': cycle,
            'Step Type': step
        }

        # Select columns to keep
        if filters.get('Step Type') == 'CC Chg':
            XY = ['Chg. Spec. Cap.(mAh/g)', 'Voltage(V)']
            print(f"Filters Cycle {cycle} & {step} applied")
        elif filters.get('Step Type') == 'CC DChg':
            XY = ['DChg. Spec. Cap.(mAh/g)', 'Voltage(V)']
            print(f"Filters Cycle {cycle} & {step} applied")
        else:
            print(f"There are no CC Chg or CC Dchg steps found")
            quit()

        # Filter and select data
        filtered_data = filter_trim(gcd_record, filters, XY)

        # Append the filtered data to plot_data
        # plot_data = pd.concat([plot_data, filtered_data], axis=0)

# Reset index after concatenation if needed
# plot_data.reset_index(drop=True, inplace=True)


# Now 'plot_data' contains the filtered and selected data based on Cycle Index and Step Type
# print(plot_data)
print("Finished")
#################################### Testing Bay ##########################################