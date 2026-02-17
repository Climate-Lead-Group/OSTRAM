# -*- coding: utf-8 -*-
"""
Created on 2025

@author: ClimateLeadGroup, Andrey Salazar-Vargas
"""

import os
import pandas as pd
import yaml
from copy import deepcopy
import sys
import shutil
import numpy as np

if __name__ == '__main__': 
    
    main_path = sys.argv
    outputs_folder = main_path[1]
    output_file = main_path[2]
    
    # outputs_folder = 'C:\\Users\\ClimateLeadGroup\\Desktop\\CLG_repositories\\relac_tx\\t1_confection\\Executables\\BAU_0\\Outputs'
    # output_file = 'C:\\Users\\ClimateLeadGroup\\Desktop\\CLG_repositories\\relac_tx\\t1_confection\\Executables\\BAU_0\\Pre_processed_BAU_0_output'


    sets_csv = [
        "YEAR",
        "TECHNOLOGY",
        "TIMESLICE",
        "FUEL",
        "EMISSION",
        "MODE_OF_OPERATION",
        "REGION",
        "SEASON",
        "DAYTYPE",
        "DAILYTIMEBRACKET",
        "STORAGE"
    ]

    sets_corrects = deepcopy(sets_csv)
    sets_corrects.insert(0,'Parameter')
    sets_corrects.append('VALUE')
    

    sets_csv_temp = deepcopy(sets_csv)
    sets_csv_temp.insert(0,'Parameter')
    sets_csv_temp.append('VALUE')
    

    count = 0

            
    # Select folder path
    tier_dir = outputs_folder

    if os.path.exists(tier_dir):
        csv_file_list = sorted(os.listdir(tier_dir))  # Sort for deterministic order
        

        df_list = []
        
        parameter_list = []
        parameter_dict = {}
        
        for f in csv_file_list:
            
            local_df = pd.read_csv(tier_dir + '/' + f)
            
            # Delete columns of sets do not use in otoole config yaml
            columns_check = [column for column in local_df.columns if column in sets_corrects]
            local_df = local_df[columns_check]

            
            local_df['Parameter'] = f.split('.')[0]
            parameter_list.append(f.split('.')[0])
            parameter_dict.update({parameter_list[-1]: local_df})
            
            df_list.append(local_df)
        columns_check.insert(0,'Parameter')
        df_all = pd.concat(df_list, ignore_index=True, sort=True)  # Sort for deterministic column order
        common_values = sorted(list(set(df_all.columns) & set(sets_csv_temp)))  # Sort for deterministic order
        df_all = df_all[ common_values ]
        
        # df_all.to_csv(f'Data_plots_{case}.csv')
        
        # 3rd try
        # Assuming parameter_list and parameter_dict are defined
        # Initialize df_all_2 with the first DataFrame to ensure the dimension columns are set
        first_param = parameter_list[0]
        df_all_3 = pd.DataFrame()
        df_all_3 = df_all[df_all['Parameter'] == first_param]
        df_all_3 = df_all_3.rename(columns={'VALUE': first_param})
        df_all_3 = df_all_3.drop('Parameter', axis=1)
        # df_all_3 = df_all_3.drop(columns=[col for col in df_all_3.columns if col in set_no_needed], errors='ignore')
        df_all_3 = df_all_3.assign(**{col: 'nan' for col in sets_csv if col not in df_all_3.columns})

        
        
        # Iterate over the remaining parameters and merge their respective DataFrames on the dimension columns
        for p in parameter_list[1:]:  # Skip the first parameter since it's already added
            local_df_3 = df_all[df_all['Parameter'] == p]
            local_df_3 = local_df_3.rename(columns={'VALUE': p})
            local_df_3 = local_df_3.drop('Parameter', axis=1)
            # local_df_3 = local_df_3.drop(columns=[col for col in local_df_3.columns if col in set_no_needed], errors='ignore')
            local_df_3 = local_df_3.assign(**{col: 'nan' for col in sets_csv if col not in local_df_3.columns})
            count+=1

            df_all_3 = pd.merge(df_all_3, local_df_3, on=sets_csv, how='outer')
        

        # # Add NPV columns
        # parameters_reference = params['parameters_reference']
        # parameters_news = params['parameters_news']
        
        # for k in range(len(parameters_reference)):
        #     if parameters_reference[k] == 'AnnualTechnologyEmissionPenaltyByEmission':
        #         parameter_filter = {'EMISSION':params['this_combina']}
        #         calculate_npv_filtered(df_all_3, parameters_news[k], parameters_reference[k], params['round_#'], 'YEAR', parameter_filter, output_csv_r=params['disc_rate']*100, output_csv_year=params['year_apply_discount_rate'])
        #     else:
        #         calculate_npv(df_all_3, parameters_news[k], parameters_reference[k], params['round_#'], 'YEAR', output_csv_r=params['disc_rate']*100, output_csv_year=params['year_apply_discount_rate'])
                            
        
        # The 'outer' join ensures that all combinations of dimension values are included, filling missing values with NaN
        # df_all_3.to_csv(f'{file_df_dir}/Data_Output_{case[-1]}.csv')

        df_all_3.to_csv(output_file + '.csv')