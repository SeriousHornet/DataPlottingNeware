# # Finds the state of cell when the impedance was measured (Whether Fresh, After Formation or After Cycling)
# def get_cell_state(filename):
#     global cell_state
#     if 'AFFM' in filename:
#         cell_state = 'After Formation'
#     elif 'Fresh' in filename:
#         cell_state = 'Fresh'
#     elif 'AFCY' in filename:
#         cell_state = 'After cycling'
#     else:
#         cell_state = None
#
# # Plot function for rate profile gcd plots
# def rpf_plotter(src_df, x):
#     trend_file_check()
#     cycle_bool = (src_df[' CYCLE'].isin([1]))
#     cycle = src_df.loc[cycle_bool]
#     lp_bool = (cycle[' LOOP'].isin([2]))
#     loop_2 = cycle.loc[lp_bool]
#     st_bool = (loop_2[' STEP'].isin([x]))
#     step = loop_2.loc[st_bool]
#
#     plt.plot(step[' ACCMAHG'], step[' VOLTS'],
#              color=next(colors),
#              label=next(labels))
#     return
#
#
# # Plot function for rate profile step plots
# def rpf_step_plotter(src_df):
#     CE_file_check()
#
#     global a_m_w
#
#     src_df['Acc mAHg'] = (src_df['Acc mAH'] / a_m_w).round(2)
#     dschg_bool = (src_df['Action'].isin(['Discharge']))
#     dschg = src_df.loc[dschg_bool]
#     cycle_bool = (dschg['Cycle'].isin([1]))
#     cycle_df = dschg.loc[cycle_bool]
#     plt.plot(cycle_df['StepID'], cycle_df['Acc mAHg'],
#              marker='s',
#              color='k',
#              linewidth=3)
#     return
#
#
# # Plot function for AC-Impedance plots
# def acimp_plotter(raw_df):
#     # Extract Z' (Ω) and -Z'' (Ω) columns
#     z_prime = raw_df['Z\' (Ω)']
#     minus_z_double_prime = raw_df['-Z\'\' (Ω)']
#
#     # Set up the scatter plot
#     plt.scatter(z_prime, minus_z_double_prime, marker='o', s=7, color='blue')
#     # Set labels and title
#     plt.xlabel('Z\' (Ω)')
#     plt.ylabel('-Z\'\' (Ω)')
#
#
# # Find number of cycles
# def find_cycles(filename):
#     match = re.search(r'@(\d+)', filename)
#     if match:
#         return int(match.group(1))
#     else:
#         return None
#
#
# # Add the suffix to the plot title if a corresponding keyword is present
# def cell_state():
#     if cell_state:
#         title = f"{cell_no}_{e_w} g_{cell_state}"
#     else:
#         title = f"{cell_no}_{e_w} g"
#     return title, cell_state

    # checking if the path has extension, if not add.
    # csv, CSV, txt, xlsx = ".csv" in file.path, ".CSV" in file.path, ".txt" in file.path, ".xlsx" in file.path
    # if txt or csv or CSV or xlsx:
    #     pass
    # else:
    #     file.path += ".csv"
# Converting capacity from Ah to mAh and add it to a new column named 'Chg. Cap.(mAH)' and 'DChg. Cap.(mAh)'
# def conv_ah_to_mah(df):
#     df['Chg. Cap.(mAh)'] = (df['Chg. Cap.(Ah)'] / 1000).__round__(4)
#     df['DChg. Cap.(mAh)'] = (df['DChg. Cap.(Ah)'] / 1000).__round__(4)
#     return df