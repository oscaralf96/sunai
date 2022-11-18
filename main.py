import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
sns.set_theme(style='darkgrid')

from datetime import datetime
import os

""" Import .xlsx files and converting them to DataFrames """
xlsx_files = []
for file in os.listdir('data/'):    
    if not os.path.isdir("results/"):
        os.mkdir("results/")
    if not os.path.isdir(f"results/{file.split('.')[0]}/"):
        os.mkdir(f"results/{file.split('.')[0]}/")
    xlsx_files.append({'data': pd.read_excel(f'data/{file}'), 'file_name': file.split('.')[0]})        

for file in xlsx_files:
    print('-' * 50)
    print(f"{file['file_name']} analysis")
    print('-' * 50)
    energy = file['data'][['fecha_im', 'id_i', 'active_energy_im']]
    print(f"Active energy missing values: {energy['active_energy_im'].isna().sum()}")
    energy = energy[energy['active_energy_im'].notnull()]
    power = file['data'][['fecha_im', 'id_i', 'active_power_im']]
    print(f"Active power missing values: {power['active_power_im'].isna().sum()}")
    power = power[power['active_power_im'].notnull()]

    print(f"There are {energy[energy['active_energy_im'].apply(pd.to_numeric, errors='coerce').notna() == False].shape[0]} active energy not number instances")
    print(f"There are {power[power['active_power_im'].apply(pd.to_numeric, errors='coerce').notna() == False].shape[0]} active power not number instances")

    power = power[power['active_power_im'].apply(pd.to_numeric, errors='coerce').notna()]
    energy = energy[energy['active_energy_im'].apply(pd.to_numeric, errors='coerce').notna()]

    sns.lineplot(x="fecha_im", y="active_power_im", hue='id_i', data=power)
    graph_path = f"results/{file['file_name']}/power.png"
    plt.savefig(graph_path)
    plt.close()

    with open(f"results/{file['file_name']}/{file['file_name']}.txt", 'w') as output_txt:
        output_txt.write(f"Today's active power sum: {power['active_power_im'].sum()}\nToday's active energy min: {energy['active_energy_im'].min()}\nToday's active energy max: {energy['active_energy_im'].max()}\ngraph path: {os.getcwd()}\\{graph_path}")