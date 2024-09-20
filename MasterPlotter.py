#  Import List
from fileinput import filename
from os.path import pathsep

import matplotlib.pyplot as plt
from itertools import cycle
import pandas as pd
import numpy as np
import os
import re
import warnings

# To avoid errors
warnings.simplefilter("ignore", category=UserWarning)
pd.options.mode.chained_assignment = None

# Declaring Variables
excel_data = None  # Electrode_Weight, Excel File Name, excel data, actual cell number, date
cell_state, title = None, None  # AC File, Fresh cell, After Formation, After Long Cycle, cell state, Plot title

# To parse file information from file's path
def file():
    # raw_path = input("Enter the path of the file: ")
    raw_path = r'"C:\Users\Mano-BRCGE\Desktop\NeWare Data\20240912- Oven-Com-ECDEC\240912-C02-C-NCMA90-EC-DEC-0.01166-CYC@59.xlsx"'
    file_path = raw_path.strip(' " " ')
    filename = str(os.path.basename(file_path))
    file.file_path = file_path
    #finding date from file name
    date = filename[0:6]
    #finding cell number
    pattern = re.compile(r'C0[1-8]')
    cell_no_match = pattern.search(filename)
    cell_no = cell_no_match.group() if cell_no_match else None
    # finding cell weight
    ew_match = re.search(r'[-+]?\d*\.\d+', filename)
    e_w = float(ew_match.group()) * -1 if ew_match else None
    #find number of cycles if any
    cyc_match = re.search(r'@(\d+)', filename)
    cyc_no = int(cyc_match.group(1)) if cyc_match else None
    # checking if the path has extension, if not add.
    # csv, CSV, txt, xlsx = ".csv" in file.path, ".CSV" in file.path, ".txt" in file.path, ".xlsx" in file.path
    # if txt or csv or CSV or xlsx:
    #     pass
    # else:
    #     file.path += ".csv"
    return {
        'filename': filename,
        'date': date,
        'cell_no': cell_no,
        'e_w': e_w,
        'cyc_no': cyc_no
    }

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

# Finds the state of cell when the impedance was measured (Whether Fresh, After Formation or After Cycling)
def get_cell_state(filename):
    global cell_state
    if 'AFFM' in filename:
        cell_state = 'After Formation'
    elif 'Fresh' in filename:
        cell_state = 'Fresh'
    elif 'AFCY' in filename:
        cell_state = 'After cycling'
    else:
        cell_state = None

# Converting capacity from Ah to mAh and add it to a new column named 'Chg. Cap.(mAH)' and 'DChg. Cap.(mAh)'
def conv_Ah_to_mAh(df):
    df['Chg. Cap.(mAh)'] = (df['Chg. Cap.(Ah)'] / 1000).__round__(4)
    df['DChg. Cap.(mAh)'] = (df['DChg. Cap.(Ah)'] / 1000).__round__(4)
    return df

# Function to filter and select specific columns
def filter_xyz(df, filters, XY):
    filtered_df = df.copy()  # Copy the original DataFrame to avoid modifying it
    for column, value in filters.items():
        filtered_df = filtered_df[filtered_df[column] == value]
    # Select the specified columns
    selected_df = filtered_df[XY]
    return selected_df

# Getting plot choice form user
def user_choice(filename):
    choices = {
        'FMN': "FMN" in filename,
        'RPF': "RPF" in filename,
        'CYC': "CYC" in filename,
        'iV': "iV" in filename,
        'CE': "CE" in filename,
        'AC': "AC" in filename
    }
    if any(choices.values()):
        return choices
    while True:
        try:
            plot_choice = int(input("1. Formation Cycles Plot \n"
                                   "or\n"
                                   "2. Rate Profile Plot \n"
                                   "or\n"
                                   "3. Long Cycles Plot \n"
                                   "Please choose 1 or 2 or 3>>: "))
            if plot_choice == 2:
                rpf_choice = int(input("4. Voltage Profile \n"
                                       "or\n"
                                       "5. Step Profile\n"
                                       "Please choose 4 or 5 >>: "))
                if rpf_choice in [4, 5]:
                    return rpf_choice
                else:
                    print("Invalid choice. Please choose 4 or 5.")
            elif plot_choice in [1, 3]:
                return plot_choice
            else:
                print("Invalid choice. Please choose 1, 2, or 3.")
        except ValueError:
            print("Invalid input. Please enter a number.")

# Finds what kind of data is being given and creates bool values to respective variables.
def plot_find(filename):
    # Keywords to check in the filename
    keywords = ["FMN", "RPF", "CYC", "iV", "CE", "AC", "Fresh", "AFFM", "AFCY"]
    # Dictionary comprehension to check if each keyword is in the filename
    result = {key: key in filename for key in keywords}
    return result

# Plot function for long cycle GCD plots
def cyc_gcd_plotter(file_info, df):
    num_cols = df.shape[1]  # Total number of columns
    num_cycles = num_cols // 4  # Each cycle has 4 columns (2 XY pairs)

    # Define a colormap with distinct colors for each cycle
    colors = plt.get_cmap('tab10', num_cycles)
    labels = cycle(['', '1', '', '10', '', '1 C', '', '5 C', '', '10 C', '', '0.2 C'])

# Plot function for making plots of a dataframe
def gcd_plotter(file_info, df):
    num_cols = df.shape[1]  # Total number of columns
    num_cycles = num_cols // 4  # Each cycle has 4 columns (2 XY pairs)

    # Define a colormap with distinct colors for each cycle
    colors = plt.get_cmap('tab10', num_cycles)
    # colors = ['k', 'r', 'b']

    # plt.figure(figsize=(10, 6))
    for i in range(num_cycles):
        # Get the indices for each XY pair of the cycle
        X1, Y1 = df.iloc[:, 4 * i], df.iloc[:, 4 * i + 1]
        X2, Y2 = df.iloc[:, 4 * i + 2], df.iloc[:, 4 * i + 3]

        # Plot both XY pairs for the same cycle in the same color
        # cycle_color = colors[i % 3]
        cycle_color = colors(i)
        line_thickness = 2.5

        plt.plot(X1, Y1, label=f'Cycle {i + 1}', color=cycle_color, linestyle='-', linewidth = line_thickness)
        plt.plot(X2, Y2, color=cycle_color, linestyle='-', linewidth = line_thickness)

    # Customize the plot
    plt.xlim(left=0)
    plt.ylim([2.7, 4.4])
    plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
    plt.ylabel("Voltage vs. Li/Li$^+$ (V)")
    if plot_type['FMN']:
        plt.title(f"Formation GCD - {file_info['cell_no']} - {file_info['e_w']} g")
        plt.savefig(f'{file_info['date']}_{file_info['cell_no']}_FMN_GCD_Plot_{file_info['e_w']}g.png',
                transparent=True,
                dpi=1000)
        plt.legend(loc='lower left', frameon=False, fancybox=False)
    elif plot_type['CYC']:
        plt.title(f"Long Cycle GCD - {file_info['cell_no']} - {file_info['e_w']} g - {file_info['cyc_no']} cycles")
        plt.savefig(f'{file_info['date']}_{file_info['cell_no']}_FMN_GCD_Plot_{file_info['e_w']}g_{file_info['cyc_no']} cycles.png',
                    transparent=True,
                    dpi=1000)
        plt.legend(loc='lower left', frameon=False, fancybox=False)
    else: pass
    plt.show()

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
def cyc_plotter(file_info, df):
    fig, ax1 = plt.subplots()
    ax1.set_xlabel("Cycles")
    ax1.set_ylabel("Specific discharge capacity (mA h $g^{-1}$)", color='blue')
    cap_plot, = ax1.plot(df['Cycles'], df['DChg. Spec. Cap.(mAh/g)'],
                         label='Discharge Capacity',
                         color='blue',
                         linewidth=3,
                         marker='o')
    ax1.tick_params(axis='y', labelcolor='blue')
    ax1.set_ylim(0, 200)
    # ax1.set_xlim(0, 200)

    ax2 = ax1.twinx()
    ax2.set_ylabel("Coulomb. Efficiency (%)", color='red')
    CE_plot, = ax2.plot(df['Cycles'], df['Coulomb. Eff.(%)'],
                        label='CE',
                        color='red',
                        linewidth=3,
                        marker='*')
    ax2.tick_params(axis='y', labelcolor='red')
    ax2.set_ylim(0, 110)
    plt.legend([cap_plot, CE_plot], ["Specific Discharge Capacity", "CE"],
        loc = 'lower left',
        frameon =False,
        fancybox =False)
    plt.title(f"Long Cycling - {file_info['cell_no']} - {file_info['e_w']}g - {file_info['cyc_no']} cycles")
    plt.savefig(f'{file_info['date']}_{file_info['cell_no']}_Long_CYC_Plot_{file_info['e_w']}g_{file_info['cyc_no']}_cycles.png',
                        transparent=True,
                        dpi=1000)
    plt.show()
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
def cell_state():
    if cell_state:
        title = f"{cell_no}_{e_w} g_{cell_state}"
    else:
        title = f"{cell_no}_{e_w} g"
    return title, cell_state

def gcd_dataset(df):
    if plot_type.get('FMN'):
        cycle_values = df['Cycle Index'].unique()
    elif plot_type.get('CYC'):
        cycle_values = [1, 10, 25, 50, 75, 100, 150, 200]
    else:
        print("Provided file is neither FMN or CYC")
        quit()

    step_values = ['CC Chg', 'CC DChg']
    plot_ready = pd.DataFrame()  # Create an empty dataframe to store results

    # Nested loop: first for 'Cycle Index', then for 'Step Type'
    for cycle in cycle_values:
        if cycle in df['Cycle Index'].values:
            for step in step_values:
                # Define the filters
                filters = {
                    'Cycle Index': cycle,
                    'Step Type': step
                }
                if filters.get('Step Type') == 'CC Chg':
                    XY = ['Chg. Spec. Cap.(mAh/g)', 'Voltage(V)']
                    print(f"Cycle {cycle} & {step} found")
                elif filters.get('Step Type') == 'CC DChg':
                    XY = ['DChg. Spec. Cap.(mAh/g)', 'Voltage(V)']
                    print(f"Cycle {cycle} & {step} found")
                else:
                    print(f"There are no CC Chg or CC Dchg steps found")
                    quit()

                # Filter and select data
                filtered_data = filter_xyz(df, filters, XY)
                dynamic_col_names = {col: f"Cycle_{cycle}_{step}" for col in filtered_data.columns}
                filtered_data.rename(columns=dynamic_col_names, inplace=True)
                # Append the filtered data to plot_data
                plot_ready = pd.concat([plot_ready.reset_index(drop=True), filtered_data.reset_index(drop=True)], axis=1)
        else: pass
    return plot_ready

def cyc_dataset(df):
    return pd.DataFrame({
        'Cycles': df['Cycle Index'],
        'Chg. Cap.(mAh)': (df['Chg. Cap.(Ah)'] * 1000).round(4),
        'DChg. Cap.(mAh)': (df['DChg. Cap.(Ah)'] * 1000).round(4),
        'Coulomb. Eff.(%)': df['Chg.-DChg. Eff(%)'],
        'Chg. Spec. Cap.(mAh/g)': df['Chg. Spec. Cap.(mAh/g)'],
        'DChg. Spec. Cap.(mAh/g)': df['DChg. Spec. Cap.(mAh/g)'],
        'Cap. Retention(%)': df['Cap. Retention(%)'],
        'Chg. Spec. Energy(mWh/g)': df['Chg. Spec. Energy(mWh/g)'],
        'DChg. Spec. Energy(mWh/g)': df['DChg. Spec. Energy(mWh/g)'],
        'Energy Eff.(%)': df['Energy Eff.(%)']
    })


#################################### Run of Events Starts Here ##########################################
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
print("Started..")

file_info = file()
filename = file_info['filename']
plot_type = plot_find(filename)

sheets4plot = ['cycle', 'step', 'record']  # the actual sheets I'm interested in the Excel file

work_frames = parse_excel(file.file_path, sheets4plot)  # Trims the dataframe to only the selected sheets
for sheet_name, df in work_frames.items(): #Converting each sheet into a separate dataframes
    globals()[sheet_name] = df  # Creates a global variable with the name of the sheet

if plot_type.get('FMN'):
    print("Formation is being plotted")
    gcd_plotter(file_info, gcd_dataset(record))
elif plot_type.get('RPF'):
    print("Rate Profile is being plotted")
elif plot_type.get('CYC'):
    print("Cycle Life is being plotted")
    cyc_plotter(file_info, cyc_dataset(cycle))
    print("Cycle Life GCD is being plotted for all cycles")
    gcd_plotter(file_info, gcd_dataset(record))
print("..Finished")