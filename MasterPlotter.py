#  Import List
import matplotlib.pyplot as plt
from itertools import cycle
import pandas as pd
import os
import re
import warnings
import sys

from openpyxl.styles.builtins import output

# To avoid errors
warnings.simplefilter("ignore", category=UserWarning)
pd.options.mode.chained_assignment = None


# Parse file information from file's path
def file():
    # raw_path = input("Enter the path of the file:")
    # raw_path = r'"C:\Users\Mano-BRCGE\Desktop\NeWare Data\20240923\240912-C01-C-NCMA90-EC-DEC-0.01098-1C-CYC@191.xlsx"'
    raw_path = sys.argv[1]
    file_path = raw_path.strip(' " " ')
    filename = str(os.path.basename(file_path))
    file.file_path = file_path
    # find date from file name
    date = filename[0:6]
    # find cell number
    pattern = re.compile(r'C0[1-8]')
    cell_no_match = pattern.search(filename)
    cell_no = cell_no_match.group() if cell_no_match else None
    # find cell weight
    ew_match = re.search(r'[-+]?\d*\.\d+', filename)
    e_w = float(ew_match.group()) * -1 if ew_match else None
    # find number of cycles if any
    cyc_match = re.search(r'@(\d+)', filename)
    cyc_no = int(cyc_match.group(1)) if cyc_match else None
    return {
        'filename': filename,
        'date': date,
        'cell_no': cell_no,
        'e_w': e_w,
        'cyc_no': cyc_no
    }


# Creat dataframes for each Excel sheet
def parse_excel(path, sheets):
    if not os.path.exists(path):
        raise FileNotFoundError(f"File '{path}' not found.")
    dataframes = {}
    # excel_data = pd.ExcelFile(path)
    with pd.ExcelFile(path, engine="openpyxl") as excel_data:
        for i in sheets:
            if i in excel_data.sheet_names:
                dataframes[i] = pd.read_excel(excel_data, sheet_name=i)
            else:
                warnings.warn(f"Warning: Sheet {i} not found in the file.")
    return dataframes


# Get plot choice form user
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


# Find what kind of data is being given and creates bool values to respective variables.
def plot_find(filename):
    # Keywords to check in the filename
    keywords = ["FMN", "RPF", "CYC", "iV", "CE", "AC", "Fresh", "AFFM", "AFCY"]
    # Dictionary comprehension to check if each keyword is in the filename
    result = {key: key in filename for key in keywords}
    return result


# Create a dataset of GCD values for plotting
def gcd_dataset(df):
    if plot_type.get('FMN'):
        cycle_values = df['Cycle Index'].unique()
        print('Data for formation plots are being created')
    elif plot_type.get('CYC'):
        cycle_values = [1, 10, 50, 100, 200]
        last_cycle = df['Cycle Index'].iloc[-1]
        for i in cycle_values:
            if i not in df['Cycle Index'].values:
                print(f"{i} not found in the data. Adding last cycle (Cycle {last_cycle}) to the data.")
                cycle_values.append(last_cycle)
        print('Data for cycling plots are being created')
    else:
        print("Provided file is neither FMN nor CYC, checking RPF")
        return

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
                plot_ready = pd.concat([plot_ready.reset_index(drop=True), filtered_data.reset_index(drop=True)],
                                       axis=1)
        else:
            pass
    return plot_ready


# Create a dataset of rate profile's GCD values for plotting
def rpf_gcd_dataset(df):
    if plot_type.get('RPF'):
        cycle_index = [2, 4, 6, 8, 10, 12, 14]
        step_pairs = [[2, 4], [7, 10], [13, 16], [19, 22], [25, 28], [31, 34], [37, 39]]
        step_type_pairs = ['CC Chg', 'CC DChg']

        plot_ready = pd.DataFrame()  # Create an empty dataframe to store results
        # Iterate through cycle_index and corresponding step_pairs
        for i, cycle in enumerate(cycle_index):
            steps = step_pairs[i]  # Get the step pairs for this cycle

            for j, step in enumerate(steps):  # j is the index (0 or 1), step is the step index (e.g., 2, 4)
                st_type = step_type_pairs[j]  # Pick 'CC Chg' for the first step, 'CC DChg' for the second
                filters = {
                    'Cycle Index': cycle,
                    'Step Index': step,
                    'Step Type': st_type
                }
                if st_type == 'CC Chg':
                    XY = ['Chg. Spec. Cap.(mAh/g)', 'Voltage(V)']
                    print(f"Cycle {cycle}, Step {step} & {st_type} found")
                elif st_type == 'CC DChg':
                    XY = ['DChg. Spec. Cap.(mAh/g)', 'Voltage(V)']
                    print(f"Cycle {cycle}, Step {step} & {st_type} found")
                else:
                    continue

                # Filter and select data
                filtered_data = filter_xyz(df, filters, XY)
                dynamic_col_names = {col: f"Cycle_{cycle}_Step{step}_{st_type}_{col}" for col in filtered_data.columns}
                filtered_data.rename(columns=dynamic_col_names, inplace=True)
                # Append the filtered data to plot_data
                plot_ready = pd.concat([plot_ready.reset_index(drop=True), filtered_data.reset_index(drop=True)],
                                       axis=1)
    else:
        return
    return plot_ready


# Create a dataset of Cycling values for plotting
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


# Plot function for making plots of a dataframe
def fmn_gcd_plotter(file_info, df, output_dir):
    num_cols = df.shape[1]  # Total number of columns
    num_cycles = num_cols // 4  # Each cycle has 4 columns (2 XY pairs)
    # Define a colormap with distinct colors for each cycle
    colors = plt.get_cmap('tab10', num_cycles)

    plt.clf()
    for i in range(num_cycles):
        # Get the indices for each XY pair of the cycle
        X1, Y1 = df.iloc[:, 4 * i], df.iloc[:, 4 * i + 1]
        X2, Y2 = df.iloc[:, 4 * i + 2], df.iloc[:, 4 * i + 3]
        # Plot both XY pairs for the same cycle in the same color
        cycle_color = colors(i)
        line_thickness = 2.5
        plt.plot(X1, Y1, label=f'Cycle {i + 1}', color=cycle_color, linestyle='-', linewidth=line_thickness)
        plt.plot(X2, Y2, color=cycle_color, linestyle='-', linewidth=line_thickness)

    # Customize the plot
    plt.xlim(left=0)
    plt.ylim([2.7, 4.4])
    plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
    plt.ylabel("Voltage vs. Li/Li$^+$ (V)")
    plt.title(f"Formation GCD - {file_info['cell_no']} - {file_info['e_w']} g")
    output_file = os.path.join(output_dir,
                               f"{file_info['date']}_{file_info['cell_no']}_FMN_GCD_Plot_{file_info['e_w']}g.png")
    plt.savefig(output_file, transparent=True, dpi=1000)
    plt.legend(loc='lower left', frameon=False, fancybox=False)
    # plt.show()
    plt.close()


def rpf_gcd_plotter(file_info, df, output_dir):
    num_cols = df.shape[1]  # Total number of columns
    num_cycles = num_cols // 4  # Each cycle has 4 columns (2 XY pairs)
    # Define a colormap with distinct colors for each cycle
    colors = plt.get_cmap('tab10', num_cycles)
    labels = ['0.2 C', '0.5 C', '1 C', '3 C', '5 C', '10 C', '0.2 C']

    plt.clf()
    for i in range(num_cycles):
        # Get the indices for each XY pair of the cycle
        X1, Y1 = df.iloc[:, 4 * i], df.iloc[:, 4 * i + 1]
        X2, Y2 = df.iloc[:, 4 * i + 2], df.iloc[:, 4 * i + 3]

        # Plot both XY pairs for the same cycle in the same color
        cycle_color = colors(i)
        cycle_label = labels[i % len(labels)]
        line_thickness = 2.5

        plt.plot(X1, Y1, label=cycle_label, color=cycle_color, linestyle='-', linewidth=line_thickness)
        plt.plot(X2, Y2, color=cycle_color, linestyle='-', linewidth=line_thickness)

    # Customize the plot
    plt.xlim(left=0)
    plt.ylim([2.7, 4.4])
    plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
    plt.ylabel("Voltage vs. Li/Li$^+$ / (V)")
    plt.title(f"0.2 C - 10 C - 0.2 C Rate Profile - {file_info['cell_no']} - {file_info['e_w']} g")
    output_file = os.path.join(output_dir,
                               f"{file_info['date']}_{file_info['cell_no']}_FMN_GCD_Plot_{file_info['e_w']}g.png")
    plt.savefig(output_file, transparent=True, dpi=1000)
    plt.legend(loc='lower left', frameon=False, fancybox=False)
    # plt.show()
    plt.close()


# Plot function for long cycle GCD plots
def cyc_gcd_plotter(file_info, df, output_dir):
    num_cols = df.shape[1]  # Total number of columns
    num_cycles = num_cols // 4  # Each cycle has 4 columns (2 XY pairs)

    # Define a colormap with distinct colors for each cycle
    colors = plt.get_cmap('tab10', num_cycles)
    # labels = cycle(['Cycle 1', 'Cycle 10', 'Cycle 50', 'Cycle 100'])
    labels = ['1', '10', '50', '100', '200']

    plt.clf()
    for i in range(num_cycles):
        # Get the indices for each XY pair of the cycle
        X1, Y1 = df.iloc[:, 4 * i], df.iloc[:, 4 * i + 1]
        X2, Y2 = df.iloc[:, 4 * i + 2], df.iloc[:, 4 * i + 3]

        # Plot both XY pairs for the same cycle in the same color
        cycle_color = colors(i)
        line_thickness = 2.5

        plt.plot(X1, Y1, label=f'Cycle {labels[i]}', color=cycle_color, linestyle='-', linewidth=line_thickness)
        plt.plot(X2, Y2, color=cycle_color, linestyle='-', linewidth=line_thickness)

    # Customize the plot
    plt.xlim(left=0)
    plt.ylim([2.7, 4.4])
    plt.xlabel("Specific Capacity (mA h $g^{-1}$)")
    plt.ylabel("Voltage vs. Li/Li$^+$ (V)")
    plt.title(f"Long Cycle GCD - {file_info['cell_no']} - {file_info['e_w']} g - {file_info['cyc_no']} cycles")
    output_file = os.path.join(output_dir,
                               f"{file_info['date']}_{file_info['cell_no']}_Long Cycle GCD Plot_{file_info['e_w']}g_@{file_info['cyc_no']}cycles.png")
    plt.savefig(output_file, transparent=True, dpi=1000)
    plt.legend(loc='lower left', frameon=False, fancybox=False)
    # plt.show()
    plt.close()


# Plot function for cycle life plots
def cyc_plotter(file_info, df, output_dir):
    plt.clf()
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
               loc='lower left',
               frameon=False,
               fancybox=False)
    plt.title(f"Long Cycling - {file_info['cell_no']} - {file_info['e_w']}g - {file_info['cyc_no']} cycles")
    output_file = os.path.join(output_dir,
                               f"{file_info['date']}_{file_info['cell_no']}_Long Cycle Plot_{file_info['e_w']}g_@{file_info['cyc_no']}cycles.png")
    plt.savefig(output_file, transparent=True, dpi=1000)
    # plt.show()
    plt.close()
    return


# Function to filter and select specific columns
def filter_xyz(df, filters, XY):
    filtered_df = df.copy()  # Copy the original DataFrame to avoid modifying it
    for column, value in filters.items():
        filtered_df = filtered_df[filtered_df[column] == value]
    # Select the specified columns
    selected_df = filtered_df[XY]
    return selected_df


def export_excel(file_info, df1, df2, df3, output_dir):
    if plot_type.get('FMN'):
        output_file = os.path.join(output_dir,
                                   f"{file_info['date']}_{file_info['cell_no']}_Formation Data_{file_info['e_w']}g.xlsx")
        df1.to_excel(output_file, sheet_name='GCD Data', index=False)
    elif plot_type.get('CYC'):
        output_file = os.path.join(output_dir,
                                   f"{file_info['date']}_{file_info['cell_no']}_Cycle Life Data_{file_info['e_w']}g_@{file_info['cyc_no']}cycles.xlsx")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='GCD Data', index=False)
            df2.to_excel(writer, sheet_name='CYC Data', index=False)
    elif plot_type.get('RPF'):
        output_file = os.path.join(output_dir,
                                   f"{file_info['date']}_{file_info['cell_no']}_Rate Profile Data_{file_info['e_w']}g.xlsx")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df3.to_excel(writer, sheet_name='Rate GCD Data', index=False)
            df2.to_excel(writer, sheet_name='Rate Step Data', index=False)
    else:
        print("Error: Invalid plot type.")
        return


print("Started..")
print('Step 1')

file_info = file()
# print('Parsed file path and stored file details to file_info')
# output_dir = None
output_dir = sys.argv[2]
filename = file_info['filename']
# print('Retrieved file name from file_info to filename')

plot_type = plot_find(filename)
# print('Identified file type and assigned respective plot type')

sheets4plot = ['cycle', 'step', 'record']  # the actual sheets I'm interested in the Excel file

work_frames = parse_excel(file.file_path, sheets4plot)
# print('Parsed Excel file and stored necessary data in a dict')  # Trims the dataframe to only the selected sheets

for sheet_name, df in work_frames.items():  # Converting each sheet into a separate dataframes
    globals()[sheet_name] = df  # Creates a global variable with the name of the sheets.
# print('Converted each sheet into a separate dataframe')

cyc_df = cyc_dataset(cycle)
gcd_df = gcd_dataset(record)
rpf_df = rpf_gcd_dataset(record)

export_excel(file_info, gcd_df, cyc_df, rpf_df, output_dir)

if plot_type.get('FMN'):
    print("Formation GCD data is being plotted")
    fmn_gcd_plotter(file_info, gcd_df, output_dir)
    print("Formation cycle data is being plotted")
    cyc_plotter(file_info, cyc_df, output_dir)

elif plot_type.get('RPF'):
    print("Rate Profile GCD data is being plotted")
    rpf_gcd_plotter(file_info, rpf_df, output_dir)

elif plot_type.get('CYC'):
    print("Cycle Life data is being plotted")
    cyc_plotter(file_info, cyc_df, output_dir)
    print("Cycle Life GCD data is being plotted for selected cycles")
    cyc_gcd_plotter(file_info, gcd_df, output_dir)

print("..Finished")
