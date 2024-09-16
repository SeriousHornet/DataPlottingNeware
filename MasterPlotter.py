#  Import List
import matplotlib.pyplot as plt
from itertools import cycle
import pandas as pd
import numpy as np
import os
import re

#
pd.options.mode.chained_assignment = None

# Declaring Variables
e_w, a_m_w, filename, cellno= None, None, None, None  # Electrode_Weight, Active_Material_Weight, Excel File Name, cell number variables, actual cell number
FMN, RPF, CYC, iV, CE = None, None, None, None, None  # Formation, Rate Profile, Long Cycle, GCD Plot, CYC vs. CAP plot, Ac Impedance
AC, Fresh, AFFM, AFCY, cell_state, title = None, None, None, None, None, None # AC File, Fresh cell, After Formation, After Long Cycle

# functions #

# Getting the file input from user
def get_file():
    global e_w, a_m_w, filename
    eal = 0.0043
    raw_csv_file = input("Enter the path of the file: ")
    path = raw_csv_file.strip(' " " ')
    filename = str(os.path.basename(path).split('/'))
    list_e_w = re.findall("\d+\.\d+", filename)
    e_w = float(np.array(list_e_w, dtype='float'))
    
    a_m_w = float(((e_w - eal) * 0.8).__round__(8))
    
    get_file.csv_file = path
    
    csv, CSV, txt = ".csv" in get_file.csv_file, ".CSV" in get_file.csv_file, ".txt" in get_file.csv_file
    if txt or csv or CSV:
        pass
    else:
        get_file.csv_file += ".csv"
    return 

# Getting weight as input
def get_weight():
    global e_w, a_m_w
    e_w = float(input("Enter the electrode weight in g: "))
    if 0.1 > e_w > 0 and len(str(e_w)) <= 7:
        pass
    else:
        print("Error: Invalid weight. Please give weight of electrode in 0.0xxxx g format!")
        exit()
    eal = 0.0043
    # eal = float(input("Enter empy Al Foil weight in g: "))  # Empty_AlFoil_Weight
    if 0.1 > eal > 0 and len(str(e_w)) <= 7:
        pass
    else:
        print("Error: Invalid weight. Please give weight of Al foil in 0.00xx g format!")
        exit()
    a_m_w = float(((e_w - eal) * 0.8).__round__(8))
    return


# A function to check if the given file is trend CSV
def trend_file_check():
    file_check = pd.read_csv(get_file.csv_file)
    if 'Cell' in file_check.columns:
        print("Error: Incorrect file! Please give the trend CSV file to proceed")
        exit()
    else:
        pass


# A function to check if the given file is CE CSV
def CE_file_check():
    file_check = pd.read_csv(get_file.csv_file)
    if ' CYCLE' in file_check.columns:
        print("Error: Incorrect file! Please give the CE CSV file to proceed")
        exit()
    else:
        pass


# Removing the unnecessary zeros from ACTION and ACCMAH
def zero_remover(src_df):
    search_0 = [0]
    only_zeros = (src_df[' ACTION'].isin(search_0)) | (src_df[' ACCMAH'].isin(search_0))
    src_df = src_df.loc[~only_zeros]
    return src_df


# Calculating Specific Capacity and adding it as a column named 'ACCMAHG'
def cap_calc(df):
    global a_m_w
    df[' ACCMAHG'] = (df[' ACCMAH'] / a_m_w).__round__(2)
    df[' VOLTS'] = (df[' VOLTAGE'] / 1000)
    return df


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


def plot_find():
    global FMN, RPF, CYC, iV, CE, AC, Fresh, AFFM, AFCY
    FMN, RPF, CYC, iV, CE = "FMN" in filename, "RPF" in filename, "CYC" in filename, "iV" in filename, "CE" in filename
    AC, Fresh, AFFM, AFCY = "AC" in filename, "Fresh" in filename, "AFFM" in filename, "AFCY" in filename
    return FMN, RPF, CYC, iV, CE, AC, Fresh, AFFM, AFCY


# Plot function for formation plots
def fmn_plotter(src_df, x, y):
    trend_file_check()

    cy_bool = (src_df[' CYCLE'].isin([x]))
    cy = src_df.loc[cy_bool]
    st_bool = (cy[' STEP'].isin([y]))
    step = cy.loc[st_bool]
    plt.plot(step[' ACCMAHG'], step[' VOLTS'],
             color=next(colors),
             label=next(labels))
    return


# Plot function for rate profile plots
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
    #ax1.set_xlim(0, 200)

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


def acimp_plotter(raw_df):
    # Extract Z' (Ω) and -Z'' (Ω) columns
    z_prime = raw_df['Z\' (Ω)']
    minus_z_double_prime = raw_df['-Z\'\' (Ω)']

    # Set up the scatter plot
    plt.scatter(z_prime, minus_z_double_prime, marker='o', s=7, color='blue')   
    # Set labels and title
    plt.xlabel('Z\' (Ω)')
    plt.ylabel('-Z\'\' (Ω)')
  
# A function to get the cell number from the file name to be used in output image file name.
def cellno():
    global cellno
    pattern = re.compile(r'C0[1-8]')
    match = pattern.search(filename)
    cellno = match.group() if match else None
    return


# Find number of cycles
def find_cycles(filename):
    match = re.search(r'@(\d+)', filename)
    if match:
        return int(match.group(1))
    else:
        return None
 
 
def cell_state():
    global cell_state, title
    if 'AFFM' in filename:
        cell_state = 'After Formation'
    elif 'Fresh' in filename:
        cell_state = 'Fresh'
    elif 'AFCY' in file_name:
        cell_state = 'After cycling'
    else:
        cell_state = None

    # Add the suffix to the title if a corresponding keyword is present
    if cell_state:
        title = f"{cellno}_{e_w} g_{cell_state}"
    else:
        title = f"{cellno}_{e_w} g"
    return title, cell_state
    
    
# Run of Events Starts Here #
get_file()
cellno()
date = filename[2:8]


# Finding out user choice
auto_find = user_choice()

# Loading file as a Pandas DataFrame
if CE or iV:
    raw_df = pd.read_csv(get_file.csv_file)
    if CYC:
        cyc_no = find_cycles(filename)
    else:
        pass 
elif AC: 
    raw_df = pd.read_csv(get_file.csv_file, delimiter=';')


if FMN or auto_find == 1:
    cap_df = cap_calc(raw_df)
    plot_ready = zero_remover(cap_df)
    cycle_x = [1, 2, 3]
    step_x = [1, 2]
    colors = cycle(['k', 'k', 'r', 'r', 'b', 'b'])
    labels = cycle(['Cycle 1', '', 'Cycle 2', '', 'Cycle 3', ''])
    for a in cycle_x:
        for b in step_x:
            fmn_plotter(plot_ready, a, b)
    # plt.autoscale(enable=True, axis='x')
    #plt.xlim([0, 220])
    plt.xlim(left=0)
    plt.ylim([2.7, 4.4])
    plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
    plt.ylabel("Voltage vs. Li/Li$^+$ (V)")
    plt.title(f"0.1 C Formation Cycles - {cellno} - {e_w} g")
    plt.legend()
    plt.savefig(f'{date}_{cellno}_FMN_GCD_Plot_{e_w}g.png',
                transparent=True,
                dpi=1000)

elif RPF or auto_find == 4 or auto_find == 5:
    if iV or auto_find == 4:
        cap_df = cap_calc(raw_df)
        plot_ready = zero_remover(cap_df)
        step_y = [1, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 21]
        colors = cycle(['k', 'k', 'r', 'r', 'b', 'b', 'g', 'g', 'm', 'm', 'c', 'c'])
        labels = cycle(['', '0.2 C', '', '0.5 C', '', '1 C', '', '5 C', '', '10 C', '', '0.2 C'])
        for i in step_y:
            rpf_plotter(plot_ready, i)

        plt.title(f"0.2 C - 10 C - 0.2 C Rate Profile  - {cellno} - {e_w} g")
        plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
        plt.ylabel("Voltage vs. Li/Li$^+$ / (V)")
        plt.xlim([0, 200])
        plt.ylim([2.7, 4.4])
        plt.legend()
        plt.savefig(f'{date}_{cellno}_RPF_GCD_Plot_{e_w}g.png',
                    transparent=True,
                    dpi=1000)
    elif CE or auto_find == 5:
        rpf_step_plotter(raw_df)
        plt.title(f"0.2 C - 10 C - 0.2 C Step Profile  - {cellno} - {e_w} g")
        plt.xlabel("C-Rates")
        cus_tick_pos = [4, 12, 21, 30, 39, 46]  
        cus_tick_labels = ['0.2 C', '0.5 C', '1 C', '5 C', '10 C', '0.2 C']
        plt.xticks(cus_tick_pos, cus_tick_labels)
        plt.ylabel("Specific Discharge Capacity (mA h $g^{-1}$)")
        # plt.autoscale(enable=True, axis='y')
        plt.ylim([50, 220])
        # plt.ylim(bottom=0)
        plt.savefig(f'{date}_{cellno}_RPF_Step_Plot_{e_w}g.png',
                transparent=True,
                dpi=1000)

elif CYC or auto_find == 3:
    cyc_no = find_cycles(filename)
    if CE:
        cyc_plotter(raw_df)
        plt.title(f"Long Cycling - {cellno} - {e_w}g - {cyc_no} cycles")
        plt.savefig(f'{date}_{cellno}_Long_CYC_Plot_{e_w}g_{cyc_no}_cycles.png',
                    transparent=True,
                    dpi=1000)
    elif iV:
        # print("Long cycle iV file found")
        cap_df = cap_calc(raw_df)
        plot_ready = zero_remover(cap_df)
        cycle_x = [1, 10, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200]
        step_x = [1, 2]
        num_of_plots = len(cycle_x)
        colors = cycle(plt.cm.viridis(np.linspace(0, 1, num_of_plots)))

        for a in cycle_x:
            for b in step_x:
                cyc_gcd_plotter(plot_ready, a, b)
            color = next(colors)
            # label = next(cycle_x)   
        plt.autoscale(enable=True, axis='x')
        # plt.xlim([0, 220])
        plt.ylim([2.7, 4.4])
        plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
        plt.ylabel("Voltage vs. Li/Li$^+$ (V)")
        plt.title(f'Long Cycling  - {cellno} - {e_w}g - {cyc_no} cycles')
        plt.legend()
        plt.savefig(f'{date}_{cellno}_Long_CYC_GCD_Plot_{e_w}g_{cyc_no}_cycles.png',
                    transparent=True,
                    dpi=1000)

elif AC:
    acimp_plotter(raw_df)
    cell_state()
    # Set equal scaling for both axes based on the last data point
    max_value = max(raw_df['Z\' (Ω)'].max(), raw_df['-Z\'\' (Ω)'].max())
    # plt.axis('equal')
    plt.title(title)
    plt.xlim(-10, max_value+50)
    plt.ylim(-10, max_value+50)
    plt.savefig(f'{date}_{cellno}_AC Impedance_Plot_{e_w}g_{cell_state}.png',
                transparent=True,
                dpi=1000)
    plt.xlim(-10, 1000)
    plt.ylim(-10, 1000)
    plt.savefig(f'{date}_{cellno}_AC Impedance_Plot_{e_w}g_{cell_state}_1K.png',
                transparent=True,
                dpi=1000)
    