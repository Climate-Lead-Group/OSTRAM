# -*- coding: utf-8 -*-
"""
Created on 2021
@author: Climate Lead Group, Luis Victor-Gallardo, Andrey Salazar-Vargas

Updated on 2025
"""
import pandas as pd
import pickle
import sys
import math
import numpy as np
import time
from copy import deepcopy
import yaml
import warnings
import os
#
start1 = time.time()
#
# Read yaml file with parameterization
with open('Config_MOMF_T1_A.yaml', 'r') as file:
    # Load content file
    params = yaml.safe_load(file)

baseyear = params['base_year']
endyear = params['final_year']
global time_range_vector
time_range_vector = [ n for n in range( int(baseyear), int(endyear)+1 ) ]
#
Wide_Param_Header = params['sets']
#
dict_xtra_scen = params['xtra_scen']
other_setup_parameters = pd.DataFrame(list(dict_xtra_scen.items()), columns=['Name', 'Param'])
other_setup_params_name = other_setup_parameters['Name'].tolist()
other_setup_params_param = other_setup_parameters['Param'].tolist()
other_setup_params_timeslices = other_setup_params_param[other_setup_params_name.index('Timeslices')]
other_setup_params_timeslices.sort()
other_setup_params = {}
for n in range( len( other_setup_params_name ) ):
    other_setup_params.update( { other_setup_params_name[n]:other_setup_params_param[n] } )
#
print_aid_parameter = False
#------------------------------------------------------------------------------
print('1 - Connect the model activity ratios.')
#
AR_Model_Base_Year = pd.ExcelFile(os.path.join(params['A1_outputs'], params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Print_Base_Year']))
AR_Projections = pd.ExcelFile(os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Print_Proj']))
groups_list = AR_Model_Base_Year.sheet_names # see all sheet names
# NOTE THIS IS THE SAME AS :: AR_Projections.sheet_names # see all sheet names

# This section affects the *InputActivityRatio* and the *OutputActivityRatio*
AR_Osemosys_Parameters = params['AR_Osemosys_Parameters']
df_IAR = pd.DataFrame( columns = Wide_Param_Header )
df_IAR_dict = {}
df_OAR = pd.DataFrame( columns = Wide_Param_Header )
df_OAR_dict = {}

All_Tech_list_system_group = []
All_Tech_list = []
All_Tech_list_names = []
All_Fuel_list = []
All_Fuel_list_names = []
#
All_Tech_len_long = 0
All_Fuel_len_long = 0
#
AR_Base_df = {}
AR_Base_proj_df = {}
AR_Base_proj_df_new = {}
AR_Base = {}
#
# Initialization of lists for accumulating data for bulk append operations
accumulated_rows_OAR = []
accumulated_rows_IAR = []

for g in range( len( groups_list ) ):
    AR_Base_df.update( { groups_list[g]:AR_Model_Base_Year.parse( groups_list[g] ) } )
    this_df = AR_Base_df[ groups_list[g] ]
    #
    AR_Base_proj_df.update( { groups_list[g]:AR_Projections.parse( groups_list[g] ) } )
    this_proj_df = AR_Base_proj_df[ groups_list[g] ]
    #
    this_proj_df_new = deepcopy( this_proj_df )
    this_proj_df_new.replace( { 'Direction': 'Output'}, 'OutputActivityRatio', inplace=True )
    this_proj_df_new.replace( { 'Direction': 'Input'}, 'InputActivityRatio', inplace=True )
    this_proj_df_new.rename(columns={ "Direction": "Parameter" }, inplace=True)
    this_proj_df_new = this_proj_df_new.drop(columns=['Projection.Mode', 'Projection.Parameter'])
    
    #
    if groups_list[g] != params['primary'] and groups_list[g] != params['transport']:
        this_df_fuel_i = this_df['Fuel.I'].tolist()
        this_df_fuel_i_name = this_df['Fuel.I.Name'].tolist()
    if groups_list[g] == params['primary']:
        this_df_fuel_i = []
        this_df_fuel_i_name = []
    if groups_list[g] == params['transport']:
        this_df_fuel_i = this_df['Fuel.I.1'].tolist()
        this_df_fuel_i_name = this_df['Fuel.I.1.Name'].tolist()
        #
        # this_df_filtered = this_df.loc[ ( this_df['Fuel.I.2'] != 'none' ) ]
        this_df_fuel_i2 = this_df['Fuel.I.2'].tolist()
        this_df_fuel_i2_name = this_df['Fuel.I.2.Name'].tolist()
        #

    # Create column "Fuel.I.Exist"
    if "Fuel.I" in this_df.columns:
        this_df["Fuel.I.Exist"] = this_df["Fuel.I"].apply(
            lambda x: "InputAcivityRatio" if pd.notna(x) else np.nan
        )
        this_df_fuel_i_exist = [
            value if pd.notna(value) else "not exist"
            for value in this_df["Fuel.I.Exist"].tolist()
        ]
    
    # Create column "Fuel.O.Exist"
    if "Fuel.O" in this_df.columns:
        this_df["Fuel.O.Exist"] = this_df["Fuel.O"].apply(
            lambda x: "OutputAcivityRatio" if pd.notna(x) else np.nan
        )
        this_df_fuel_o_exist = [
            value if pd.notna(value) else "not exist"
            for value in this_df["Fuel.O.Exist"].tolist()
        ]

    #
    this_df_techs = this_df['Tech'].tolist()
    this_df_techs_names = this_df['Tech.Name'].tolist()
    All_Tech_len_long += len( this_df_techs )
    #
    this_df_fuel_o = this_df['Fuel.O'].tolist()
    this_df_fuel_o_name = this_df['Fuel.O.Name'].tolist()
    #------------------------------------------------------#
    this_df_mode_of_operation= this_df['Mode.Operation'].tolist()
    #------------------------------------------------------#
    these_fuels = this_df_fuel_i + this_df_fuel_o
    these_fuels_names = this_df_fuel_i_name + this_df_fuel_o_name
    #------------------------------------------------------#
    for t in range( len( this_df_techs ) ):
        if this_df_techs[t] not in All_Tech_list:
            All_Tech_list.append( this_df_techs[t] )
            All_Tech_list_names.append( this_df_techs_names[t] )
    #
    for f in range( len( these_fuels ) ):
        if these_fuels[f] not in All_Fuel_list:
            All_Fuel_list.append( these_fuels[f] )
            All_Fuel_list_names.append( these_fuels_names[f] )
    #
    #------------------------------------------------------#
    tech_plus_fuel_unique_oar, tech_plus_fuel_unique_iar = [], []

    # Let us continue with a useful dictionary gathering all the data:
    for t in range( len( this_df_techs ) ):
        this_tech = this_df_techs[t]
        output_fuel = this_df_fuel_o[t]
        mode_of_operation_order = this_df_mode_of_operation[t]
        if "Fuel.I" in this_df.columns:
            this_input_fuel_exist = this_df_fuel_i_exist[t]
        if "Fuel.O" in this_df.columns:
            this_output_fuel_exist = this_df_fuel_o_exist[t]
        
        #
        if this_output_fuel_exist != 'not exist':
            # Query the output // Applies to all *groups_list* 
            this_df_select = this_df.loc[ ( this_df[ 'Tech' ] == this_tech ) & ( this_df[ 'Fuel.O' ] == output_fuel ) & ( this_df[ 'Mode.Operation' ] == mode_of_operation_order ) & ( this_df[ 'Fuel.O.Exist' ] == this_output_fuel_exist ) ]
            this_df_select_by_fo = this_df_select[ 'Value.Fuel.O' ].tolist()[0]
            
            this_proj_df_local = this_proj_df.loc[ ( this_proj_df[ 'Tech' ] == this_tech ) & ( this_proj_df[ 'Fuel' ] == output_fuel ) ] 
            this_proj_df_mode_o, this_proj_df_param_o = this_proj_df_local['Projection.Mode'].tolist()[0] , this_proj_df_local['Projection.Parameter'].tolist()[0]
        #
        if groups_list[g] != params['primary'] and groups_list[g] != params['transport']:
            # query 1 input and 1 output
            input_fuel = this_df_fuel_i[t]
            input_fuel_2 = 'none'
        
        #
        if groups_list[g] == params['transport']:
            # query 2 inputs
            input_fuel = this_df_fuel_i[t]
            input_fuel_2 = this_df_fuel_i2[t]
        #
        if groups_list[g] != params['primary']:
            if 'PWRBCK' not in this_tech:
                if groups_list[g] != params['transport']:
                    if this_input_fuel_exist != 'not exist':
                        this_df_select = this_df.loc[ ( this_df[ 'Tech' ] == this_tech ) & ( this_df[ 'Fuel.I' ] == input_fuel ) & ( this_df[ 'Mode.Operation' ] == mode_of_operation_order ) & ( this_df[ 'Fuel.I.Exist' ] == this_input_fuel_exist ) ]
                        this_df_select_by_fi = [ this_df_select[ 'Value.Fuel.I' ].tolist()[0] ]
                        #
                        this_proj_df_local = this_proj_df.loc[ ( this_proj_df[ 'Tech' ] == this_tech ) & ( this_proj_df[ 'Fuel' ] == input_fuel ) & ( this_proj_df[ 'Mode.Operation' ] == mode_of_operation_order ) ] 
                        this_proj_df_mode_i, this_proj_df_param_i = [ this_proj_df_local['Projection.Mode'].tolist()[0] ], [ this_proj_df_local['Projection.Parameter'].tolist()[0] ]  
                    #
                else:
                    this_df_select = this_df.loc[ ( this_df[ 'Tech' ] == this_tech ) & ( this_df[ 'Fuel.I.1' ] == input_fuel ) ]
                    this_df_select_by_fi = [ this_df_select[ 'Value.Fuel.I.1' ].tolist()[0] ]
                    #
                    this_proj_df_local = this_proj_df.loc[ ( this_proj_df[ 'Tech' ] == this_tech ) & ( this_proj_df[ 'Fuel' ] == input_fuel ) & ( this_proj_df[ 'Mode.Operation' ] == mode_of_operation_order ) ] 
                    this_proj_df_mode_i, this_proj_df_param_i = [ this_proj_df_local['Projection.Mode'].tolist()[0] ], [ this_proj_df_local['Projection.Parameter'].tolist()[0] ]  
                    #
                    if input_fuel_2 != 'none':
                        this_df_select = this_df.loc[ ( this_df[ 'Tech' ] == this_tech ) & ( this_df[ 'Fuel.I.2' ] == input_fuel_2 ) & ( this_df[ 'Mode.Operation' ] == mode_of_operation_order ) ]
                        this_df_select_by_fi += [ this_df_select[ 'Value.Fuel.I.2' ].tolist()[0] ]
                        #
                        this_proj_df_local = this_proj_df.loc[ ( this_proj_df[ 'Tech' ] == this_tech ) & ( this_proj_df[ 'Fuel' ] == input_fuel_2 ) & ( this_proj_df[ 'Mode.Operation' ] == mode_of_operation_order ) ] 
                        this_proj_df_mode_i += [ this_proj_df_local['Projection.Mode'].tolist()[0] ]
                        this_proj_df_param_i += [ this_proj_df_local['Projection.Parameter'].tolist()[0] ]
        #
        else:
            this_proj_df_mode_i = ''
        #
        if print_aid_parameter == True:
            print( groups_list[g], this_tech, this_proj_df_mode_o, this_proj_df_mode_i )    
        #
        
        for y in range( len( time_range_vector ) ):
            this_param = 'OutputActivityRatio'
            if this_output_fuel_exist != 'not exist':
                mask = ( this_proj_df_new[ 'Tech' ] == this_tech ) & ( this_proj_df_new[ 'Fuel' ] == output_fuel ) & ( this_proj_df_new[ 'Parameter' ] == this_param ) & ( this_proj_df_new[ 'Mode.Operation' ] == mode_of_operation_order )
    
                if this_proj_df_mode_o == params['flat']:
                    this_proj_df_new.loc[ mask , time_range_vector[y] ] = round( this_df_select_by_fo, 4 )
                
                #
                this_mask_index = this_proj_df_new.loc[ mask , str(time_range_vector[y]) ].index.tolist()[0]
                this_mask_index_moo = this_proj_df_new.loc[ mask , 'Mode.Operation' ].index.tolist()[0]
                # Handling OutputActivityRatio
                this_value_o = deepcopy(
                    this_proj_df_new.loc[mask, str(time_range_vector[y])][this_mask_index])
                this_mode_of_operation = deepcopy(this_proj_df_new.loc[ mask , 'Mode.Operation' ][this_mask_index_moo])
                #
                # Filling the data :
                if this_tech + '+' + output_fuel + '+' + str(this_mode_of_operation) not in tech_plus_fuel_unique_oar:                
                    
                    oar_row = {
                        'PARAMETER': 'OutputActivityRatio',
                        'Scenario': other_setup_params['Main_Scenario'],
                        'REGION': other_setup_params['Region'],
                        'TECHNOLOGY': this_tech,
                        'FUEL': output_fuel,
                        'MODE_OF_OPERATION': this_mode_of_operation,
                        'YEAR': time_range_vector[y],
                        'Value': deepcopy(this_value_o)  # Placeholder for the computed value for OutputActivityRatio
                    }
                    accumulated_rows_OAR.append(oar_row)
                    # if this_tech == 'PWRWONCRIXX01':
                    #     sys.exit(9)
                
            #
            if groups_list[g] != params['primary']:
                input_fuel_list = [ input_fuel ]
                if input_fuel_2 != 'none':
                    input_fuel_list += [ input_fuel_2 ]
                #
                for inp in range( len( input_fuel_list ) ):
                    this_input_fuel = input_fuel_list[inp]
                    this_proj_df_mode_i0 = this_proj_df_mode_i[inp]
                    this_proj_df_param_i0 = this_proj_df_param_i[inp]
                    #
                    this_param = 'InputActivityRatio'
                    if 'PWRBCK' in this_tech:
                        continue
                    if this_input_fuel_exist != 'not exist':
                        mask = ( this_proj_df_new[ 'Tech' ] == this_tech ) & ( this_proj_df_new[ 'Fuel' ] == this_input_fuel ) & ( this_proj_df_new[ 'Parameter' ] == this_param ) & ( this_proj_df_new[ 'Mode.Operation' ] == mode_of_operation_order )
                        ###################################################################################################
                        if this_proj_df_mode_i0 == params['flat']:
                            this_proj_df_new.loc[ mask , str(time_range_vector[y]) ] = round( this_df_select_by_fi[ inp ], 4 )
                        if this_proj_df_mode_i0 == params['yearly_per_change']:
                            if y == 0:
                                this_proj_df_new.loc[ mask , str(time_range_vector[y]) ] = round( this_df_select_by_fi[ inp ], 4 ) # round( this_df_select_by_fi[ inp ]*( 1 + this_proj_df_param_i0/100 ), 4 )
                            else:
                                this_proj_df_new.loc[ mask , str(time_range_vector[y]) ] = round( this_proj_df_new.loc[ mask , str(time_range_vector[y-1]) ]*( 1 + this_proj_df_param_i0/100 ), 4 )
                        #
                        if this_proj_df_mode_i0 == params['user_defined']:
                            
                            this_proj_df_new.loc[ mask , str(time_range_vector[y]) ] = round( this_proj_df.loc[ mask , str(time_range_vector[y]) ], 4 )
    
                        #
                        ###################################################################################################
                        # Filling the data :
                                                                                              
                        this_mask_index = this_proj_df_new.loc[ mask , str(time_range_vector[y]) ].index.tolist()[0]
                        this_mask_index_moo = this_proj_df_new.loc[ mask , 'Mode.Operation' ].index.tolist()[0]
    
                        # Here, instead of updating df_IAR_dict and appending directly to df_IAR:
                        this_value = deepcopy(this_proj_df_new.loc[mask, str(time_range_vector[y])][this_mask_index])
                        this_mode_of_operation = deepcopy(this_proj_df_new.loc[ mask , 'Mode.Operation' ][this_mask_index_moo])
    
                        # Create a new dictionary for each row to append
                        iar_row = {
                            'PARAMETER': this_param, 
                            'Scenario': other_setup_params['Main_Scenario'],
                            'REGION': other_setup_params['Region'], 
                            'TECHNOLOGY': this_tech, 
                            'FUEL': this_input_fuel,
                            'MODE_OF_OPERATION': this_mode_of_operation, 
                            'YEAR': time_range_vector[y], 
                            'Value': this_value
                        }
                        # Append the new dictionary to the list for bulk append later
                        accumulated_rows_IAR.append(iar_row)
                    #
                #
            #
        # Create a tech + fuel string to show uniqueness in oar values:
        if this_output_fuel_exist != 'not exist':
            if this_tech + '+' + output_fuel + '+' + str(this_mode_of_operation) not in tech_plus_fuel_unique_oar:
                tech_plus_fuel_unique_oar.append(this_tech + '+' + output_fuel + '+' + str(this_mode_of_operation) )
        #
    #
    AR_Base_proj_df_new.update({ groups_list[g]:this_proj_df_new })
#
# After the loops, convert accumulated data into DataFrames and concatenate them with the original ones
with warnings.catch_warnings():
    warnings.simplefilter("ignore", category=FutureWarning)
    df_OAR = pd.concat([df_OAR, pd.DataFrame(accumulated_rows_OAR)], ignore_index=True) if accumulated_rows_OAR else df_OAR
    df_IAR = pd.concat([df_IAR, pd.DataFrame(accumulated_rows_IAR)], ignore_index=True) if accumulated_rows_IAR else df_IAR


# HERE WE HAVE A FUNCTIONING IAR AND OAR FOR BOTH NEEDS: WIDE AND LONG FORMATS
#
print('3 (end) - The model has ben connected.')
#------------------------------------------------------------------------------
# DEMAND
print('4 - Process the model demand.')

Demand = pd.ExcelFile(os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Print_Demand']))
param_sheets = Demand.sheet_names # see all sheet names



# Define variable to check how many Timeslices model has
timeslices_dict = {}
if params['xtra_scen']['Timeslice'] == 'Some':
    timeslices_dict.update( { 'Profiles':Demand.parse( 'Profiles' ) } )
    timeslices_dict = timeslices_dict['Profiles']['Timeslices'].unique().tolist()
else:
    timeslices_dict = []
timeslices_list_check = timeslices_dict


#
df_SpecAnnualDemand = pd.DataFrame( columns = Wide_Param_Header )
df_SpecDemandProfile = pd.DataFrame( columns = Wide_Param_Header )
#


Projections = pd.ExcelFile(params['A2_extra_inputs'] + params['Xtra_Proj'])
Projections_sheet = Projections.parse( Projections.sheet_names[0] ) # see all sheet names
Projections_control = Projections.parse( Projections.sheet_names[1] )
Projection_Driver_Vars = Projections_control[ 'Variable' ].tolist()
Projection_Mode_per_Driver = Projections_control[ 'Projection Mode' ].tolist()
for n in range( len( Projection_Driver_Vars ) ):
    this_col_header = Projection_Driver_Vars[ n ]
    if Projection_Mode_per_Driver[n] == params['inter_final_value']:
        # Projections_sheet[ this_col_header ].interpolate(method ='linear', limit_direction ='forward', inplace = True)
        Projections_sheet[this_col_header] = Projections_sheet[this_col_header].interpolate(
            method='linear', limit_direction='forward'
        )

    if Projection_Mode_per_Driver[n] == params['flat']:
        # Projections_sheet[ this_col_header ].interpolate(method ='linear', limit_direction ='forward', inplace = True)
        Projections_sheet[this_col_header] = Projections_sheet[this_col_header].interpolate(
            method='linear', limit_direction='forward'
        )
    if Projection_Mode_per_Driver[n] == params['flat_aft_last_year']:
        # Projections_sheet[ this_col_header ].interpolate(method ='linear', limit_direction ='forward', inplace = True)
        Projections_sheet[this_col_header] = Projections_sheet[this_col_header].interpolate(
            method='linear', limit_direction='forward'
        )


#
demand_headers = params['demand_headers']
#
# Initialize lists to accumulate data
accumulated_rows_SpecAnnualDemand = []
accumulated_rows_SpecDemandProfile = []
#
for s in range( len( param_sheets ) ):
    Demand_df = Demand.parse( Demand.sheet_names[s] )
    list_demand_or_share = Demand_df[ 'Demand/Share' ].tolist()
    #
    list_fuel_or_tech = Demand_df[ 'Fuel/Tech' ].tolist()
    Fuels_dems = [ i for i in range( len( list_fuel_or_tech ) ) if params['E6'] in list_fuel_or_tech[i] ] 
    Fuels_techs = [ i for i in range( len( list_fuel_or_tech ) ) if params['techs'] in list_fuel_or_tech[i] ]
    Fuels_techs_2_dems = {}
    for i in range( len( list_fuel_or_tech ) ):
        if i in Fuels_techs:
            for k in range( len( Fuels_dems ) ):
                if Fuels_dems[k] > i:
                    Fuel_dem_index = Fuels_dems[k-1]
                    break
            Fuels_techs_2_dems.update( { list_fuel_or_tech[i]:list_fuel_or_tech[Fuel_dem_index] } )
    #
    list_projection_mode = Demand_df[ 'Projection.Mode' ].tolist()
    list_projection_param = Demand_df[ 'Projection.Parameter' ].tolist()

    for m in range( len( list_demand_or_share ) ):
        # This is the case for *Passenger* transport:
        if params['gdp'] in list_projection_mode[m]:
            this_tech = list_fuel_or_tech[m]
            other_value_BY = 0
            if params['joint'] in list_projection_mode[m]:
                other_tech = list_projection_mode[m].split(' ')[-1]
                other_tech_index = list_fuel_or_tech.index( other_tech )
                other_value_BY = Demand_df.loc[ other_tech_index, params['initial_year_gdp'] ]
            this_value_BY = Demand_df.loc[ m, params['initial_year_gdp'] ]
            #
            this_net_value_BY = this_value_BY + other_value_BY
            #
            demand_trajectory = [ this_net_value_BY ]
            for y in range( len( Projections_sheet['Year'].tolist() ) ):
    
                if params['passenger'] in list_projection_param[m]:
                    demand_trajectory.append( demand_trajectory[-1]*( 1 + Projections_sheet['e_Passenger'].tolist()[y]*Projections_sheet['Variation_GDP'].tolist()[y]/100 ) )
                if params['freight_column'] in list_projection_param[m]:
                    demand_trajectory.append( demand_trajectory[-1]*( 1 + Projections_sheet['e_Freight'].tolist()[y]*Projections_sheet['Variation_GDP'].tolist()[y]/100 ) )
                Demand_df.loc[ m, Projections_sheet['Year'].tolist()[y] ] = round( demand_trajectory[-1]*( this_value_BY/( this_value_BY + other_value_BY ) ), 4 )
            #
        #
        if 'Flat' == list_projection_mode[m]:                                               
            this_value_BY = Demand_df.loc[ m, params['initial_year'] ]
            for y in range(len(Projections_sheet['Year'].tolist())):
                Demand_df.loc[ m, Projections_sheet['Year'].tolist()[y] ] = round(this_value_BY, 4)
        #
        # Appending to df_SpecAnnualDemand and df_SpecDemandProfile
        if Demand_df['Demand/Share'].tolist()[m] == 'Demand' and param_sheets[s] == 'Demand_Projection':
            this_fuel = list_fuel_or_tech[m]
            for y in range(len(time_range_vector)):
                # SpecifiedAnnualDemand
                spec_annual_demand_row = {
                    'PARAMETER': 'SpecifiedAnnualDemand', 
                    'Scenario': other_setup_params['Main_Scenario'],
                    'REGION': other_setup_params['Region'], 
                    'FUEL': this_fuel,
                    'YEAR': time_range_vector[y],
                    'Value': Demand_df.loc[m, str(time_range_vector[y])]
                }
                accumulated_rows_SpecAnnualDemand.append(spec_annual_demand_row)
                
            
            if other_setup_params_timeslices != timeslices_list_check and other_setup_params['Timeslice'] == 'Some' and \
                timeslices_list_check != []:
                print('These variables are differents, so you need check A-O_Demand.xlsx sheet Timeslices and Config_MOMF_T1_A.yaml variable xtra_scen/Timeslices')
                sys.exit()
            elif other_setup_params['Timeslice'] == 'Some' and timeslices_list_check != []:
                timeslices_list = timeslices_list_check
            elif (
                    (other_setup_params['Timeslice'] == 'Some' and not timeslices_list_check)
                    or
                    (other_setup_params['Timeslice'] == 'All' and timeslices_list_check)
                ):
                print('Check the defintion of Timeslices, into A-O_Demand.xlsx sheet Timeslices and Config_MOMF_T1_A.yaml variable xtra_scen/Timeslice')
                sys.exit()
            elif other_setup_params['Timeslice'] == 'All':
                timeslices_list = [other_setup_params['Timeslice']]
            
    
            for ts in range(len(timeslices_list)):  
                for y in range(len(time_range_vector)):
                    # SpecifiedDemandProfile
                    
                    # Condition when the model has one timeslice
                    if param_sheets[s] != 'Profiles' and len(timeslices_list_check)<2:
                        spec_demand_profile_row = {
                            'PARAMETER': 'SpecifiedDemandProfile', 
                            'Scenario': other_setup_params['Main_Scenario'],
                            'REGION': other_setup_params['Region'], 
                            'FUEL': this_fuel,
                            'TIMESLICE': timeslices_list[ts], 
                            'YEAR': time_range_vector[y],
                            'Value': 1
                        }
                        accumulated_rows_SpecDemandProfile.append(spec_demand_profile_row)
    
        if param_sheets[s] == 'Profiles':
            this_fuel = list_fuel_or_tech[m]
            timeslices_sdp = Demand_df['Timeslices'].tolist()
            # for ts in range(len(timeslices_sdp)):  
            for y in range(len(time_range_vector)):
                # SpecifiedDemandProfile

                # Condition when the model has one or more than one timeslice
                # if Demand_df.loc[m, 'Timeslices'] == timeslices_sdp[m]:
                spec_demand_profile_row = {
                    'PARAMETER': 'SpecifiedDemandProfile', 
                    'Scenario': other_setup_params['Main_Scenario'],
                    'REGION': other_setup_params['Region'], 
                    'FUEL': this_fuel,
                    'TIMESLICE': timeslices_sdp[m], 
                    'YEAR': time_range_vector[y],
                    'Value': Demand_df.loc[m, str(time_range_vector[y])]
                }
                accumulated_rows_SpecDemandProfile.append(spec_demand_profile_row)
                    # sys.exit(8)

with warnings.catch_warnings():
    warnings.simplefilter("ignore", category=FutureWarning)
    # Convert the accumulated rows into DataFrames and append them to the original DataFrames
    if accumulated_rows_SpecAnnualDemand:
        new_rows_SpecAnnualDemand_df = pd.DataFrame(accumulated_rows_SpecAnnualDemand)
        df_SpecAnnualDemand = pd.concat([df_SpecAnnualDemand, new_rows_SpecAnnualDemand_df], ignore_index=True)
    
    if accumulated_rows_SpecDemandProfile:
        new_rows_SpecDemandProfile_df = pd.DataFrame(accumulated_rows_SpecDemandProfile)
        df_SpecDemandProfile = pd.concat([df_SpecDemandProfile, new_rows_SpecDemandProfile_df], ignore_index=True)
                                                                                                   

print('5 - Parameterize technologies.')
# HERE WE HAVE A FUNCTIONING DEMAND *DF* // still have to add the wide format
# ------------------------------------------------------------------------------
# THIS SECTION ONLY PARAMETERIZES TECHNOLOGIES:

Battery_Replacement = pd.ExcelFile(params['A2_extra_inputs'] + params['Xtra_Battery'])
Battery_Replacement_df = Battery_Replacement.parse( Battery_Replacement.sheet_names[0] ) # see all sheet names

if params['Use_Transport']:
    Fleet = pd.ExcelFile(os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Print_Fleet']))
    Fleet_df = Fleet.parse( Fleet.sheet_names[0] )
    Fleet_Groups = pickle.load( open( os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Pickle_Fleet_Groups']), "rb" ) )
    Fleet_Groups_Distance = {}
    Fleet_Groups_OR = {} # *OR* is occupancy rate




#
Parametrization = pd.ExcelFile(os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Print_Paramet']))
param_sheets = Parametrization.sheet_names # see all sheet names]
if 'growth_formula' in param_sheets:
    param_sheets.remove('growth_formula')
#
params_dict = {}
params_dict_new = {}
params_dict_new_natural = {}
params_columns_dict = {}
#
# Let us quickly obtain the parameter list found in this excel file:
overall_param_list = []
for s in range( len( param_sheets ) ):
    this_df = Parametrization.parse( param_sheets[s] )
    this_df = this_df.replace('\xa0', np.nan)                                        
    overall_param_list_raw = this_df[ 'Parameter' ].tolist()
    for p in range( len( list( set( overall_param_list_raw ) ) ) ):
        if list( set( overall_param_list_raw ) )[p] not in overall_param_list:
            overall_param_list.append( list( set( overall_param_list_raw ) )[p] )
#
overall_param_df_dict = {}
overall_param_df_dict.update( { 'InputActivityRatio':df_IAR } )
overall_param_df_dict.update( { 'OutputActivityRatio':df_OAR } )
overall_param_df_dict.update( { 'SpecifiedAnnualDemand':df_SpecAnnualDemand } )
overall_param_df_dict.update( { 'SpecifiedDemandProfile':df_SpecDemandProfile } )
for p in range( len( overall_param_list ) ):
    if overall_param_list[p] != 'OutputActivityRatio':
        overall_param_df_dict.update( { overall_param_list[p]:pd.DataFrame( columns = Wide_Param_Header ) } )
#*****************************************************************************
print('5.a. - Capacity limits.')
# Let us do the capacity limits for the group technologies
'''
Description: this section changes the units from *Demand_df* in % to Gpkm for the modes of transport
Alternatives: change the value of oar (occupancy rate) in time, here it is left constant
'''
Demand_df_new = deepcopy( Demand_df )
Demand_df_techs = Demand_df_new[ 'Fuel/Tech' ].tolist()

if params['Use_Transport']:
    groups_list = list( Fleet_Groups.keys() )
    # Initialize dictionaries to accumulate data for each parameter
    accumulated_data = {param: [] for param in ['TotalAnnualMaxCapacity', 'TotalTechnologyAnnualActivityLowerLimit']}
    for g in range( len( groups_list ) ):
        this_group_tech_index = Demand_df_techs.index( groups_list[g] )
        oar = Demand_df_new.loc[ this_group_tech_index, 'Ref.OAR.BY' ]
        closest_demand_index = None
        if isinstance(oar, (float, int)):                                 
            for d in range( this_group_tech_index ):
                if params['E6'] in Demand_df_new.loc[ d, params['fuel__tech'] ]:
                    closest_demand_index =  d
            # Assuming Fleet_Groups_OR is predefined and handled as per your logic
            Fleet_Groups_OR.update( { groups_list[g]:[ oar for y in range( len( time_range_vector ) ) ] } )
            if closest_demand_index is not None:  # Ensure closest_demand_index was found
                for y in range(len(time_range_vector)):
                    ref_demand_value = Demand_df_new.loc[closest_demand_index, time_range_vector[y]]
                    target_share = Demand_df.loc[this_group_tech_index, time_range_vector[y]]
                    adjusted_value = round(ref_demand_value * target_share / oar, 4)
                    Demand_df_new.loc[this_group_tech_index, time_range_vector[y]] = adjusted_value
    
                    # Preparing data for 'TotalAnnualMaxCapacity' and 'TotalTechnologyAnnualActivityLowerLimit'
                    for param in ['TotalAnnualMaxCapacity', 'TotalTechnologyAnnualActivityLowerLimit']:
                        accumulated_data[param].append({
                            'PARAMETER': param,
                                         
                                              
                            'Scenario': other_setup_params['Main_Scenario'],
                            'REGION': other_setup_params['Region'],
                            'TECHNOLOGY': groups_list[g],
                            'YEAR': time_range_vector[y],
                            'Value': deepcopy(adjusted_value)
                        })
                #
            Demand_df_new.loc[ this_group_tech_index, 'Target.Unit' ] = params['Gvkm']
        #
        elif oar == params['not_considered']: # Handling 'not considered', assumed for RAIL
    
            for y in range(len(time_range_vector)):
                ref_cap = Demand_df_new.loc[this_group_tech_index, time_range_vector[y]]
                for param in params['cap_params_Rail']:
                    accumulated_data[param].append({
                        'PARAMETER': param,
                        'Scenario': other_setup_params['Main_Scenario'],
                        'REGION': other_setup_params['Region'],
                        'TECHNOLOGY': groups_list[g],
                        'YEAR': time_range_vector[y],
                        'Value': deepcopy(round(ref_cap, 4))
                    })

    # Convert the accumulated data into DataFrames and append them to the corresponding DataFrames within overall_param_df_dict
    for param in accumulated_data:
        if accumulated_data[param]:  # Check if there is any data to append
            new_rows_df = pd.DataFrame(accumulated_data[param])
            overall_param_df_dict[param] = pd.concat([overall_param_df_dict[param], new_rows_df], ignore_index=True)




# Fill list to be part of a dataframe later on:
accumulated_data = {}
#
# Define variable to check how many Timeslices model has
timeslices_dict = {}
if params['xtra_scen']['Timeslice'] == 'Some':
    timeslices_dict.update( { 'Capacities':Parametrization.parse( 'Capacities' ) } )
    timeslices_dict = timeslices_dict['Capacities']['Timeslices'].unique().tolist()
else:
    timeslices_dict = []
timeslices_list_check = timeslices_dict
timeslices_list_check.sort()
#
print('5.b. - Remaining parameters.')
for s in range( len( param_sheets ) ):
    params_dict.update( { param_sheets[s]:Parametrization.parse( param_sheets[s] ) } )
    if param_sheets[s] != 'Yearsplit' and param_sheets[s] != 'DaySplit':
        this_df = params_dict[ param_sheets[s] ]
        #
        this_df = this_df.replace('\xa0', np.nan)                                        
        this_df_new = deepcopy( this_df )
        this_df_new_2 = deepcopy( this_df )
        #
        this_df.columns = this_df.columns.map(str)
        this_df_new.columns = this_df_new.columns.map(str)
        this_df_new_2.columns = this_df_new_2.columns.map(str)
        #
        params_columns_dict.update( { param_sheets[s]:this_df.columns.tolist() } )
        #
        df_index_list = this_df.index.tolist()
        df_index_list_with_data = []
        #
        this_df_tech_list = this_df['Tech' ].tolist()
        #
        for n in range( len( df_index_list ) ):
            this_tech = this_df.loc[ n, 'Tech' ]
            this_param = this_df.loc[ n, 'Parameter' ]
            
            if s != 0: # s == 0 does not need to change in time
                this_projection_mode = this_df.loc[ n, 'Projection.Mode' ]
                #-----------------------------------------
                if this_projection_mode == params['flat']:
                    for y in range( len( time_range_vector ) ):
                        this_df_new.loc[ n, time_range_vector[y] ] = round( this_df.loc[ n, params['initial_year'] ], 4 )
                        this_df_new_2.loc[ n, time_range_vector[y] ] = this_df_new.loc[ n, time_range_vector[y] ]
                #-----------------------------------------
                if this_projection_mode == params['perce_grow_incom_years']:
                    growth_param = this_df.loc[ n, 'Projection.Parameter' ]
                    for y in range( len( time_range_vector ) ):
                        value_field = this_df.loc[ n, time_range_vector[y] ]
                        if math.isnan(value_field) == True:
                            this_df_new.loc[ n, time_range_vector[y] ] = round( this_df_new.loc[ n, time_range_vector[y-1] ]*( 1 + growth_param/100 ), 4 )
                            this_df_new_2.loc[ n, time_range_vector[y] ] = this_df_new.loc[ n, time_range_vector[y] ]
                #-----------------------------------------
                if this_projection_mode == params['user_defined']:
                    for y in range( len( time_range_vector ) ):
                        this_df_new.loc[ n, str(time_range_vector[y]) ] = round( this_df.loc[ n, str(time_range_vector[y]) ], 4 )
                        this_df_new_2.loc[ n, str(time_range_vector[y]) ] = this_df_new.loc[ n, str(time_range_vector[y]) ]
                #-----------------------------------------
                if this_projection_mode == params['inter_Stated_value_proj_param']:
                    final_year_to_interpolate = this_df.loc[ n, 'Projection.Parameter' ]
                    #
                    x_coord_tofill, xp_coord_known, yp_coord_known = [], [], []
                    flatten_y_index = 0
                    for y in range( len( time_range_vector ) ):
                        value_field = this_df.loc[ n, time_range_vector[y] ]
                        #
                        if y != 0:
                            value_field_prior = this_df.loc[ n, time_range_vector[y-1] ]
                            if y > flatten_y_index and flatten_y_index != 0:
                                xp_coord_known.append( y )
                                yp_coord_known.append( flatten_y_value )
                            if math.isnan(value_field) == False and math.isnan(value_field_prior) == True and time_range_vector[y] > final_year_to_interpolate:
                                flatten_y_index = y
                                flatten_y_value = this_df.loc[ n, time_range_vector[y] ]
                                xp_coord_known.append( y )
                                yp_coord_known.append( flatten_y_value )
                            if time_range_vector[y] <= final_year_to_interpolate and math.isnan(value_field) == True:
                                xp_coord_known.append( y )
                                yp_coord_known.append( value_field_prior )
                                this_df.loc[ n, time_range_vector[y] ] = value_field_prior
                            if time_range_vector[y] <= final_year_to_interpolate and math.isnan(value_field) == False:
                                xp_coord_known.append( y )
                                yp_coord_known.append( value_field )
                            if time_range_vector[y] > final_year_to_interpolate and math.isnan(value_field) == True and flatten_y_index == 0:
                                x_coord_tofill.append( y )
                        #
                        else:
                            xp_coord_known.append( y )
                            yp_coord_known.append( value_field )
                        #
                    #
                    y_coord_filled = list( np.interp( x_coord_tofill, xp_coord_known, yp_coord_known ) )
                    interpolated_values = []
                    for coord in range( len( time_range_vector ) ):
                        if coord in xp_coord_known:
                            value_index = xp_coord_known.index(coord)
                            interpolated_values.append( float( yp_coord_known[value_index] ) )
                        elif coord in x_coord_tofill:
                            value_index = x_coord_tofill.index(coord)
                            interpolated_values.append( float( y_coord_filled[value_index] ) )
                    #
                    for y in range( len( time_range_vector ) ):
                        this_df_new.loc[ n, time_range_vector[y] ] = round( interpolated_values[y], 4 )
                        this_df_new_2.loc[ n, time_range_vector[y] ] = this_df_new.loc[ n, time_range_vector[y] ]
                    #
                #-----------------------------------------
                if this_projection_mode == params['zero']:
                    for y in range( len( time_range_vector ) ):
                        this_df_new.loc[ n, time_range_vector[y] ] = 0
                        this_df_new_2.loc[ n, time_range_vector[y] ] = this_df_new.loc[ n, time_range_vector[y] ]
                #
            #********************************************#
            if param_sheets[s] == params['veh_techs']:
                this_projection_mode = this_df.loc[ n, 'Projection.Mode' ]
                this_unit_introduced = this_df.loc[ n, 'Unit.Introduced' ] 
                this_unit_target = this_df.loc[ n, 'Unit' ] 
                if type(this_projection_mode) != str:
                    this_unit_target = ''
                if type(this_unit_introduced) != str:
                    this_unit_target = ''
                if type(this_unit_target) != str:
                    this_unit_target = ''
                # we must use fleet // demand // projections to calibrate the vehicle fleets
                if type( this_unit_introduced ) == str:
                
                    if params['relative'] in this_unit_introduced and this_projection_mode == params['user_rel_BY']:
                        ref_tech = this_unit_introduced.split(' ')[-1]
                        ref_tech_index = this_df_tech_list.index(ref_tech)
                        ref_value = this_df.loc[ ref_tech_index, time_range_vector[0] ]
                        for y in range( len( time_range_vector ) ):
                            if y == 0:
                                this_df_new.loc[ n, time_range_vector[y] ] = ref_value*this_df.loc[ n, time_range_vector[y] ]
                            else:
                                this_df_new.loc[ n, time_range_vector[y] ] = this_df_new.loc[ n, time_range_vector[0] ]*this_df.loc[ n, time_range_vector[y] ]                    
                            this_df_new_2.loc[ n, time_range_vector[y] ] = this_df_new.loc[ n, time_range_vector[y] ]
                    #
                    if params['relative'] in this_unit_introduced and this_projection_mode == params['flat']:
                        ref_tech = this_unit_introduced.split(' ')[-1]
                        ref_tech_index = this_df_tech_list.index(ref_tech)
                        ref_value = this_df.loc[ ref_tech_index, time_range_vector[0] ]
                        #
                        for y in range( len( time_range_vector ) ):
                            this_df_new.loc[ n, time_range_vector[y] ] = ref_value*this_df.loc[ n, time_range_vector[0] ]
                            this_df_new_2.loc[ n, time_range_vector[y] ] = this_df_new.loc[ n, time_range_vector[y] ]
                    #
                    if params['relative'] not in this_unit_introduced and this_projection_mode == params['user_rel_BY']:
                        for y in range( len( time_range_vector ) ): # these units are NOT relative to another tech, so we proceeed as follows:
                            if y == 0:
                                pass # nothing is necessary to do
                            else:
                                this_df_new.loc[ n, time_range_vector[y] ] = this_df.loc[ n, time_range_vector[0] ]*this_df.loc[ n, time_range_vector[y] ]
                                this_df_new_2.loc[ n, time_range_vector[y] ] = this_df_new.loc[ n, time_range_vector[y] ]
                    #
                    #**********************************#
                    # Now we need introduce the variable of kilometers, which is grabbed from the demand. Furthermore, we must pay attention to the restrictions.
                    groups_list = list( Fleet_Groups.keys() )
                    for g in range( len( groups_list ) ):
                        if this_tech in Fleet_Groups [ groups_list[g] ]:
                            this_group = groups_list[g]
                    #
                    tech_list_from_demand_df = Demand_df['Fuel/Tech'].tolist()
                    index_call_demand_df = tech_list_from_demand_df.index( this_group )
                    ref_km = Demand_df.loc[ index_call_demand_df, params['ref_km_by'] ]
                    
                    # print( n, this_tech, ref_km )
                    
                    if type( ref_km ) == float or type( ref_km ) == int:
                        #
                        if params['freight'] not in this_group:
                            this_km_list_change = Projections_sheet['Variation_km_Passenger'].tolist()
                        else:
                            this_km_list_change = Projections_sheet['Variation_km_Freight'].tolist()
                        #
                        this_km_list = [ ref_km ]
                        for y in range( len( this_km_list_change ) ):
                            this_km_list.append( round( this_km_list[-1]*( 1+this_km_list_change[y]/100 ), 4 ) ) # this should be added to a dataframe
                        #
                        Fleet_Groups_Distance.update( { this_tech:this_km_list } )
                        #
    
                        if params['Gvkm'] in this_unit_target and ( params['veh'] in this_unit_introduced or params['relative'] in this_unit_introduced ): # this acts on the cost parameters
    
                            for y in range( len( time_range_vector ) ): # we employ a conversion of units
                                this_df_new_2.loc[ n, time_range_vector[y] ] = round( 1000*this_df_new.loc[ n, time_range_vector[y] ]/this_km_list[y] , 4 )
                            this_df_new_2.loc[ n, 'Unit.Introduced' ] = this_df_new.loc[ n, 'Unit' ]
                        #
                        if params['Gvkm'] in this_unit_target and params['vehs'] in this_unit_introduced: # this acts on the capacity parameters
                            # print( 'got here' )
                            for y in range( len( time_range_vector ) ): # we employ a conversion of units
                                this_df_new_2.loc[ n, time_range_vector[y] ] = round( this_df_new.loc[ n, time_range_vector[y] ]*this_km_list[y]/(1e9), 4 )
                            this_df_new_2.loc[ n, 'Unit.Introduced' ] = this_df_new.loc[ n, 'Unit' ]
                        #
                    #
                #
                if this_param == 'TotalAnnualMaxCapacity' or this_param == 'TotalTechnologyAnnualActivityLowerLimit':
                    define_restriction = False
                    fleet_df_tech_list = Fleet_df['Techs'].tolist()
                    fleet_df_tech_index = fleet_df_tech_list.index( this_tech )
                    #
                    target_year = Fleet_df.loc[ fleet_df_tech_index, 'Target Year']
                    target_value = Fleet_df.loc[ fleet_df_tech_index, 'Target Value']
                    start_year = Fleet_df.loc[ fleet_df_tech_index, 'Start Year']
                    if Fleet_df.loc[ fleet_df_tech_index, params['target_type']] == params['hard']: # proceed
                        define_restriction = True
                    if Fleet_df.loc[ fleet_df_tech_index, params['target_type']] == params['lower'] and this_param == 'TotalTechnologyAnnualActivityLowerLimit': # proceed
                        define_restriction = True
                    if define_restriction == True:
                        demand_df_index = Demand_df_techs.index( this_group ) # this enables us to call the Gvkm value of the restriction
                        if start_year == params['continuos']: # means we need to take the residual capacity into consideration
                            this_residual_cap_row = this_df_new_2.loc[ ( this_df_new_2['Tech'] == this_tech ) & ( ( this_df_new_2['Parameter'] == 'ResidualCapacity' ) ) ] # units in Gvkm
                            this_main_cap_list = []
                            this_residual_cap_list = []
                            this_relative_cap_list = []
                            this_correct_capacity = []
                            for y in range( len( time_range_vector ) ):
                                this_residual_cap_list.append( this_residual_cap_row[ time_range_vector[y] ].tolist()[0] )
                                this_main_cap_list.append( Demand_df_new.loc[ demand_df_index, time_range_vector[y] ] )
                                this_relative_cap_list.append( this_residual_cap_list[-1] / this_main_cap_list[-1] )
                                if this_relative_cap_list[-1] < target_value/100:
                                    this_correct_capacity.append( round( this_main_cap_list[-1]*target_value/100, 4 ) )
                                else:
                                    this_correct_capacity.append( round( this_residual_cap_list[-1], 4 ) )
                                #
                                this_df_new_2.loc[ n, time_range_vector[y] ] = round( this_correct_capacity[-1], 4 )
                                #
                                new_row = {
                                    'PARAMETER': this_param,
                                    'Scenario': other_setup_params['Main_Scenario'],
                                    'REGION': other_setup_params['Region'],
                                    'TECHNOLOGY': this_tech,
                                    'YEAR': time_range_vector[y],
                                    'Value': deepcopy(round(this_correct_capacity[-1], 4))
                                }
    
                                # Check if this parameter already has a list in accumulated_data; if not, initialize it
                                if this_param not in accumulated_data:
                                    accumulated_data[this_param] = []
    
                                # Append the new row to the list for this parameter in accumulated_data
                                accumulated_data[this_param].append(new_row)
    
                        #
                        if type( start_year ) == int:
                            x_coord_tofill, xp_coord_known, yp_coord_known = [], [], []
                            for y in range( len( time_range_vector ) ):
                                if time_range_vector[y] < start_year:
                                    xp_coord_known.append( y )
                                    yp_coord_known.append( 0 )
                                if time_range_vector[y] >= start_year and time_range_vector[y] < target_year:
                                    x_coord_tofill.append( y )
                                if time_range_vector[y] == target_year:
                                    xp_coord_known.append( y )
                                    yp_coord_known.append( target_value/100 )
                            #
                            y_coord_filled = list( np.interp( x_coord_tofill, xp_coord_known, yp_coord_known ) )
                            interpolated_values = []
                            for coord in range( len( time_range_vector ) ):
                                if coord in xp_coord_known:
                                    value_index = xp_coord_known.index(coord)
                                    interpolated_values.append( float( yp_coord_known[value_index] ) )
                                elif coord in x_coord_tofill:
                                    value_index = x_coord_tofill.index(coord)
                                    interpolated_values.append( float( y_coord_filled[value_index] ) )
                            #
                            for y in range( len( time_range_vector ) ):
                                this_main_cap_list.append( Demand_df_new.loc[ demand_df_index, time_range_vector[y] ] )
                                this_df_new_2.loc[ n, time_range_vector[y] ] = round( this_main_cap_list[-1]*interpolated_values[y], 4 )
                                # Inside your condition where you're preparing this_dict_4_wide
                                new_row = {
                                    'PARAMETER': this_param,
                                    'Scenario': other_setup_params['Main_Scenario'],
                                    'REGION': other_setup_params['Region'],
                                    'TECHNOLOGY': this_tech,
                                    'YEAR': time_range_vector[y],
                                    'Value': deepcopy(round(this_main_cap_list[-1] * interpolated_values[y], 4))
                                }
    
                                # Check if this parameter already has a list in accumulated_data; if not, initialize it
                                if this_param not in accumulated_data:
                                    accumulated_data[this_param] = []
    
                                # Append the new row to the list for this parameter in accumulated_data
                                accumulated_data[this_param].append(new_row)                                
                #
            #
            #********************************************#
            # Remember to call: *this_tech*, *this_param*
            if s == 0:  # we store the parameters without years
                new_row = {
                    'PARAMETER': this_param,
                    'Scenario': other_setup_params['Main_Scenario'],
                    'REGION': other_setup_params['Region'],
                    'TECHNOLOGY': this_tech,
                    'Value': deepcopy(this_df_new_2.loc[n, 'Value'])
                }
                if this_param not in accumulated_data:
                    accumulated_data[this_param] = []
                accumulated_data[this_param].append(new_row)
    
            elif this_projection_mode not in ['EMPTY', 'Zero', '', None, 'According to demand'] and type(this_projection_mode) == str and param_sheets[s] != 'Other_Techs':
                if this_param in ['CapacityFactor']:
                    
                    if other_setup_params_timeslices != timeslices_list_check and other_setup_params['Timeslice'] == 'Some' and \
                        timeslices_list_check != []:
                        print('These variables are differents, so you need check A-O_Parametrization.xlsx sheet Timeslices and Config_MOMF_T1_A.yaml variable xtra_scen/Timeslices')
                        sys.exit()
                    elif other_setup_params['Timeslice'] == 'Some' and timeslices_list_check != []:
                        timeslices_list = timeslices_list_check
                    elif (other_setup_params['Timeslice'] == 'Some' and timeslices_list_check == []) or \
                        (other_setup_params['Timeslice'] == 'All' and timeslices_list_check != []):
                        print('Check the defintion of Timeslices, into A-O_Parametrization.xlsx sheet Timeslices and Config_MOMF_T1_A.yaml variable xtra_scen/Timeslice')
                        sys.exit()
                    elif other_setup_params['Timeslice'] == 'All':
                        timeslices_list = [other_setup_params['Timeslice']]
                    
                    timeslices_capacities = this_df_new_2['Timeslices'].tolist()
   
                    for y in range(len(time_range_vector)):
                        
                        # Condition when the model has one or more than one timeslice
                        if param_sheets[s] == 'Capacities' and timeslices_list_check != []:
                            # Prepare dictionary to accumulate data based on conditions
                            new_row = {
                                'PARAMETER': this_param,
                                'Scenario': other_setup_params['Main_Scenario'],
                                'REGION': other_setup_params['Region'],
                                'TECHNOLOGY': this_tech,
                                'YEAR': time_range_vector[y],
                                'Value': deepcopy(round(this_df_new_2.loc[n, str(time_range_vector[y])], 4)),
                                'TIMESLICE': timeslices_capacities[n]
                            }

                        elif param_sheets[s] != 'Capacities' and timeslices_list_check == []:
                            # Prepare dictionary to accumulate data based on conditions
                            new_row = {
                                'PARAMETER': this_param,
                                'Scenario': other_setup_params['Main_Scenario'],
                                'REGION': other_setup_params['Region'],
                                'TECHNOLOGY': this_tech,
                                'YEAR': time_range_vector[y],
                                'Value': deepcopy(round(this_df_new_2.loc[n, str(time_range_vector[y])], 4)),
                                'TIMESLICE': other_setup_params['Timeslice']
                            }   
                            
                        # Accumulate data for bulk append
                        if this_param not in accumulated_data:
                            accumulated_data[this_param] = []
                        accumulated_data[this_param].append(new_row)
                        
                elif this_param in ['VariableCost']:
                    
                    if len(other_setup_params['Mode_of_Operation'])>1:
                        modes_of_operation = other_setup_params['Mode_of_Operation']
                    
                        # for moo in range(len(modes_of_operation)):    
                        for y in range(len(time_range_vector)):
                            
                            # Condition when the model has one or more than one timeslice
                            if param_sheets[s] == 'VariableCost':
                                # Prepare dictionary to accumulate data based on conditions
                                new_row = {
                                    'PARAMETER': this_param,
                                    'Scenario': other_setup_params['Main_Scenario'],
                                    'REGION': other_setup_params['Region'],
                                    'TECHNOLOGY': this_tech,
                                    'YEAR': time_range_vector[y],
                                    'MODE_OF_OPERATION': this_df_new_2.loc[n, 'Mode.Operation'],
                                    'Value': deepcopy(round(this_df_new_2.loc[n, str(time_range_vector[y])], 4)),
                                }
                                
                                
                                # Accumulate data for bulk append
                                if this_param not in accumulated_data:
                                    accumulated_data[this_param] = []
                                accumulated_data[this_param].append(new_row) 
                                
                else:
                    for y in range(len(time_range_vector)):
                                             
                        # Prepare dictionary to accumulate data based on conditions
                        if this_param not in params['sheets_exceptions_to_fuel']:
                            new_row = {
                                'PARAMETER': this_param,
                                'Scenario': other_setup_params['Main_Scenario'],
                                'REGION': other_setup_params['Region'],
                                'TECHNOLOGY': this_tech,
                                'YEAR': time_range_vector[y],
                                'Value': deepcopy(round(this_df_new_2.loc[n, str(time_range_vector[y])], 4))
                            }
                            
                        elif this_param in params['sheets_exceptions_to_fuel']:
                            new_row = {
                                'PARAMETER': this_param,
                                'Scenario': other_setup_params['Main_Scenario'],
                                'REGION': other_setup_params['Region'],
                                'FUEL': this_tech,
                                'YEAR': time_range_vector[y],
                                'Value': deepcopy(round(this_df_new_2.loc[n, str(time_range_vector[y])], 4))
                            }
        
                        # Additional conditions for specific parameters
                        if this_param in ['VariableCost'] and len(other_setup_params['Mode_of_Operation'])<2:
                            new_row['MODE_OF_OPERATION'] = other_setup_params['Mode_of_Operation'][0]
        
                        # Accumulate data for bulk append
                        if this_param not in accumulated_data:
                            accumulated_data[this_param] = []
                        accumulated_data[this_param].append(new_row)
                        #
                    #
                #
            #
        #   
    #
    params_dict_new.update( { param_sheets[s]:this_df_new_2 } ) # this has the model ready for osemosys, but it may be less intuitive for the rest of the system
    params_dict_new_natural.update( { param_sheets[s]:this_df_new } ) # this has the model in natural terms
    #
# After the loop, convert accumulated data into DataFrames and append them in bulk
for param, rows in sorted(accumulated_data.items()):  # Sort for deterministic order
    if rows:  # Check if there are rows to append for this parameter
        new_rows_df = pd.DataFrame(rows)
        if param in overall_param_df_dict:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=FutureWarning)
                overall_param_df_dict[param] = pd.concat([overall_param_df_dict[param], new_rows_df], ignore_index=True)
        else:
            overall_param_df_dict[param] = new_rows_df                                                                                  
#
#------------------------------------------------------------------------------
print('6 - Define YearSplit and DaySplit')
#
df_Yearsplit = pd.DataFrame( columns = Wide_Param_Header )
# Initialize an empty list to accumulate dictionaries
accumulated_dicts_Yearsplit = []

if other_setup_params['Timeslice'] == 'Some' and other_setup_params_timeslices != []:
    timeslices_list = other_setup_params_timeslices
    this_df = params_dict[ 'Yearsplit' ]
    
    df_index_list = this_df.index.tolist()
    #
    for n in range( len( df_index_list ) ):
        this_ts = this_df.loc[ n, 'Timeslices' ]

        # Loop through the time range vector
        for y in range(len(time_range_vector)):
            # Create a dictionary for the current iteration
            this_dict_4_wide = {
                'PARAMETER': 'YearSplit',
                'Scenario': other_setup_params['Main_Scenario'],
                'TIMESLICE': this_ts,
                'YEAR': time_range_vector[y],
                'Value': deepcopy(this_df.loc[n, str(time_range_vector[y])])
            }
            # Add the dictionary to the list
            accumulated_dicts_Yearsplit.append(this_dict_4_wide)
    
elif other_setup_params['Timeslice'] == 'Some' and other_setup_params_timeslices == []:
    print('These variables have inconsistance variable xtra_scen/Timeslice and xtra_scen/Timeslices')
    sys.exit()
elif other_setup_params['Timeslice'] == 'All':
    timeslices_list = [other_setup_params['Timeslice']]

    for ts in range(len(timeslices_list)):
        # Loop through the time range vector
        for y in range(len(time_range_vector)):
            # Create a dictionary for the current iteration
            this_dict_4_wide = {
                'PARAMETER': 'YearSplit',
                'Scenario': other_setup_params['Main_Scenario'],
                'TIMESLICE': timeslices_list[ts],
                'YEAR': time_range_vector[y],
                'Value': 1
            }
            # Add the dictionary to the list
            accumulated_dicts_Yearsplit.append(this_dict_4_wide)

# Convert the accumulated list of dictionaries to a DataFrame
new_rows_Yearsplit_df = pd.DataFrame(accumulated_dicts_Yearsplit)

# Use pd.concat to append the new rows to the original DataFrame
with warnings.catch_warnings():
    warnings.simplefilter("ignore", category=FutureWarning)
    df_Yearsplit = pd.concat([df_Yearsplit, new_rows_Yearsplit_df], ignore_index=True)
overall_param_df_dict.update( { 'YearSplit':df_Yearsplit } )

#
df_DaySplit = pd.DataFrame( columns = Wide_Param_Header )
# Initialize an empty list to accumulate dictionaries
accumulated_dicts_DaySplit = []

if len(other_setup_params['DailyTimeBracket']) > 1:
    timeslices_list = other_setup_params_timeslices
    this_df = params_dict[ 'DaySplit' ]
    
    df_index_list = this_df.index.tolist()
    #
    for n in range( len( df_index_list ) ):
        this_dtb = this_df.loc[ n, 'DAILYTIMEBRACKET' ]

        # Loop through the time range vector
        for y in range(len(time_range_vector)):
            # Create a dictionary for the current iteration
            this_dict_4_wide = {
                'PARAMETER': 'DaySplit',
                'Scenario': other_setup_params['Main_Scenario'],
                'DAILYTIMEBRACKET': this_dtb,
                'YEAR': time_range_vector[y],
                'Value': deepcopy(this_df.loc[n, str(time_range_vector[y])])
            }
            # Add the dictionary to the list
            accumulated_dicts_DaySplit.append(this_dict_4_wide)
    
elif len(other_setup_params['DailyTimeBracket'])<2 and other_setup_params_timeslices == []:
    print('These variables have inconsistance variable xtra_scen/DailyTimeBracket')
    sys.exit()
elif other_setup_params['DailyTimeBracket'] == 1:
    timeslices_list = [other_setup_params['DailyTimeBracket'][0]]

    # Loop through the time range vector
    for y in range(len(time_range_vector)):
        # Create a dictionary for the current iteration
        this_dict_4_wide = {
            'PARAMETER': 'DaySplit',
            'Scenario': other_setup_params['Main_Scenario'],
            'DAILYTIMEBRACKET': other_setup_params['DailyTimeBracket'][0],
            'YEAR': time_range_vector[y],
            'Value': 1
        }
        # Add the dictionary to the list
        accumulated_dicts_DaySplit.append(this_dict_4_wide)

# Convert the accumulated list of dictionaries to a DataFrame
new_rows_DaySplit_df = pd.DataFrame(accumulated_dicts_DaySplit)

# Use pd.concat to append the new rows to the original DataFrame
with warnings.catch_warnings():
    warnings.simplefilter("ignore", category=FutureWarning)
    df_DaySplit = pd.concat([df_DaySplit, new_rows_DaySplit_df], ignore_index=True)
overall_param_df_dict.update( { 'DaySplit':df_DaySplit } )
#------------------------------------------------------------------------------
print('7 - Include emissions.')
Emissions = pd.ExcelFile(params['A2_extra_inputs'] + params['Xtra_Emi'])
# Emissions.sheet_names # see all sheet names // this only need thes wide format
Emissions_ghg_df = Emissions.parse( params['GHGs'] )
Emissions_ext_df = Emissions.parse( params['Externalities'] )
#
emissions_list = list( set( Emissions_ghg_df['Emission'].tolist() + Emissions_ext_df['External Cost'].tolist() ) )

if params['Use_OG_module']:
    Emissions_OG = pd.read_csv(os.path.join('OG_csvs_inputs', 'EMISSION.csv'))
    emissions_list = Emissions_OG['VALUE'].tolist()

#
df_Emissions = pd.DataFrame( columns = Wide_Param_Header )
these_e_modeofoperation = Emissions_ghg_df['Mode_Of_Operation'].tolist()
these_emissions = Emissions_ghg_df['Emission'].tolist()
these_e_techs = Emissions_ghg_df['Tech'].tolist()
these_e_values = Emissions_ghg_df['EmissionActivityRatio'].tolist()

accumulated_emission_data = []

for e in range(len(these_emissions)):
    this_emission = these_emissions[e]
    this_tech = these_e_techs[e]
    
    for y in range(len(time_range_vector)):
        new_row = {
            'PARAMETER': 'EmissionActivityRatio',
            'Scenario': other_setup_params['Main_Scenario'],
            'REGION': other_setup_params['Region'],
            'TECHNOLOGY': this_tech,
            'EMISSION': this_emission,
            'MODE_OF_OPERATION': these_e_modeofoperation[e],
            'YEAR': time_range_vector[y],
            'Value': these_e_values[e]
        }
        accumulated_emission_data.append(new_row)

if accumulated_emission_data:  # Ensure there's something to append
    new_emissions_df = pd.DataFrame(accumulated_emission_data)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=FutureWarning)
        df_Emissions = pd.concat([df_Emissions, new_emissions_df], ignore_index=True)
#
df_EmissionPenalty = pd.DataFrame( columns = Wide_Param_Header )
these_emissions_name = Emissions_ext_df['Emission'].tolist()
these_emissions = Emissions_ext_df['External Cost'].tolist()
these_e_techs = Emissions_ext_df['Tech'].tolist()
these_e_modeofoperation= Emissions_ext_df['Mode_Of_Operation'].tolist()
these_e_values = Emissions_ext_df['EmissionActivityRatio'].tolist()
these_e_penalty = Emissions_ext_df['EmissionsPenalty'].tolist()
this_emission_unique = []
# Initialize accumulators
accumulated_emission_data = []
accumulated_emission_penalty_data = []
emission_unique_set = set()  # For tracking unique emission-year combinations

for e in range(len(these_emissions)):
    this_emission = these_emissions_name[e]
    this_tech = these_e_techs[e]  # Assuming this is relevant for both emissions and penalties
    this_modeofoperation = these_e_modeofoperation[e]

    for y in range(len(time_range_vector)):
        # # Accumulate EmissionActivityRatio data
        # accumulated_emission_data = {
        #     'PARAMETER': 'EmissionActivityRatio',
        #     'Scenario': other_setup_params['Main_Scenario'],
        #     'REGION': other_setup_params['Region'],
        #     'TECHNOLOGY': this_tech,
        #     'EMISSION': this_emission,
        #     'MODE_OF_OPERATION': these_e_modeofoperation,
        #     'YEAR': time_range_vector[y],
        #     'Value': these_e_values[e]  # Assuming these_e_values corresponds to EmissionActivityRatio values
        # }

        # Check for uniqueness and accumulate EmissionsPenalty data if unique
        emission_year_key = f"{this_emission} {time_range_vector[y]}"
        if emission_year_key not in emission_unique_set:
            new_row = {
                'PARAMETER': 'EmissionsPenalty',
                'Scenario': other_setup_params['Main_Scenario'],
                'REGION': other_setup_params['Region'],
                'TECHNOLOGY': '',  # Empty as per the structure for penalties
                'EMISSION': this_emission,
                'YEAR': time_range_vector[y],
                'Value': these_e_penalty[e]  # Assuming these_e_penalty corresponds to EmissionsPenalty values
            } 
            accumulated_emission_penalty_data.append(new_row)

# # Bulk append for EmissionActivityRatio
# if accumulated_emission_data:
#     df_Emissions = pd.concat([df_Emissions, pd.DataFrame(accumulated_emission_data)], ignore_index=True)

# Bulk append for EmissionsPenalty
if accumulated_emission_penalty_data:
    # df_EmissionPenalty = pd.concat([df_EmissionPenalty, pd.DataFrame(accumulated_emission_penalty_data)], ignore_index=True)
    new_emissions_penalty_df = pd.DataFrame(accumulated_emission_penalty_data)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=FutureWarning)
        df_EmissionPenalty = pd.concat([df_EmissionPenalty, new_emissions_penalty_df], ignore_index=True)

# Update the overall_param_df_dict
overall_param_df_dict['EmissionActivityRatio'] = df_Emissions
overall_param_df_dict['EmissionsPenalty'] = df_EmissionPenalty                                                                                                                      
#
# Let us create the basis for the LTS:                                    
overall_param_df_dict_ndp = deepcopy(overall_param_df_dict)

# %%
additional_params_df = params_dict_new[ 'Other_Techs' ]
additional_params_list = additional_params_df[ 'Parameter' ].tolist()
additional_tech_list = additional_params_df[ 'Tech' ].tolist()
additional_fuel_list = additional_params_df[ 'Fuel' ].tolist()
# Initialize a dictionary to hold accumulators for each parameter
accumulated_data_by_param = {}

for p in range(len(additional_params_list)):
    this_param = additional_params_list[p]
    this_tech = additional_tech_list[p]
    this_fuel = additional_fuel_list[p] if additional_fuel_list[p] != 'none' else ''

    # Ensure the tech and fuel lists are updated
                      
     
    if this_tech not in All_Tech_list:
        All_Tech_list.append(this_tech)
    if this_fuel and this_fuel not in All_Fuel_list:
        All_Fuel_list.append(this_fuel)

    for y in range(len(time_range_vector)):
        new_row = {
            'PARAMETER': this_param,
            'Scenario': other_setup_params['Main_Scenario'],
            'REGION': other_setup_params['Region'],
            'TECHNOLOGY': this_tech,
            'FUEL': this_fuel,
            'YEAR': time_range_vector[y],
            'Value': deepcopy(additional_params_df.loc[p, time_range_vector[y]])
        }

        # Special handling for OutputActivityRatio
        if this_param == 'OutputActivityRatio':
            new_row['MODE_OF_OPERATION'] = other_setup_params['Mode_of_Operation']
            new_row['Value'] = deepcopy(round(additional_params_df.loc[p, time_range_vector[y]], 4))

        # Accumulate data based on parameter
        if this_param not in accumulated_data_by_param:
            accumulated_data_by_param[this_param] = []
        accumulated_data_by_param[this_param].append(new_row)

for param, rows in sorted(accumulated_data_by_param.items()):  # Sort for deterministic order
    if rows:  # Check if there are rows to append for this parameter
        new_rows_df = pd.DataFrame(rows)
        # Check if the parameter already exists in overall_param_df_dict_ndp, if not, initialize it
        if param not in overall_param_df_dict_ndp:
            overall_param_df_dict_ndp[param] = pd.DataFrame(columns=Wide_Param_Header)
        overall_param_df_dict_ndp[param] = pd.concat([overall_param_df_dict_ndp[param], new_rows_df], ignore_index=True)
#------------------------------------------------------------------------------
print('8 - Include Storage.')
Storages = pd.ExcelFile(params['A2_extra_inputs'] + params['Xtra_Storage'])
# Emissions.sheet_names # see all sheet names // this only need thes wide format
xtra_storage_fixed_hori_param = Storages.parse( params['xs_1'] )
xtra_storage_capitalcost = Storages.parse( params['xs_2'] )
xtra_storage_technology_sto = Storages.parse( params['xs_3'] )

#
df_StorageLevelStart = pd.DataFrame( columns = Wide_Param_Header )
df_OperationalLifeStorage = pd.DataFrame( columns = Wide_Param_Header )

storagelevelstart_data = []
operationallifestorage_data = []

this_df = xtra_storage_fixed_hori_param
df_index_list = this_df.index.tolist()
#
for n in range( len( df_index_list ) ):
    this_storage = deepcopy(this_df.loc[n, 'STORAGE'])
    this_parameter_storage = deepcopy(this_df.loc[n, 'Parameter'])
    these_s_values = deepcopy(this_df.loc[n, 'Value'])
    if this_parameter_storage == 'StorageLevelStart':
        new_row = {
            'PARAMETER': this_parameter_storage,
            'Scenario': other_setup_params['Main_Scenario'],
            'REGION': other_setup_params['Region'],
            'STORAGE': this_storage,
            'Value': these_s_values
        }
        storagelevelstart_data.append(new_row)
    
    elif this_parameter_storage == 'OperationalLifeStorage':
        new_row = {
            'PARAMETER': this_parameter_storage,
            'Scenario': other_setup_params['Main_Scenario'],
            'REGION': other_setup_params['Region'],
            'STORAGE': this_storage,
            'Value': these_s_values
        }
        operationallifestorage_data.append(new_row)

if storagelevelstart_data:  # Ensure there's something to append
    new_storagelevelstart_df = pd.DataFrame(storagelevelstart_data)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=FutureWarning)
        df_StorageLevelStart = pd.concat([df_StorageLevelStart, new_storagelevelstart_df], ignore_index=True)
if operationallifestorage_data:  # Ensure there's something to append
    new_operationallifestorage_df = pd.DataFrame(operationallifestorage_data)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=FutureWarning)
        df_OperationalLifeStorage = pd.concat([df_OperationalLifeStorage, new_operationallifestorage_df], ignore_index=True)

#
df_CapitalCostStorage = pd.DataFrame( columns = Wide_Param_Header )
df_ResidualStorageCapacity= pd.DataFrame( columns = Wide_Param_Header )

capitalcoststorage_data = []
residualstoragecapacity_data = []

this_df = xtra_storage_capitalcost
df_index_list = this_df.index.tolist()
#
for n in range( len( df_index_list ) ):
    this_storage = deepcopy(this_df.loc[n, 'STORAGE'])
    this_parameter = deepcopy(this_df.loc[n, 'Parameter'])
    this_projection_mode = deepcopy(this_df.loc[n, 'Projection.Mode'])
    # Loop through the time range vector
    for y in range(len(time_range_vector)):
        this_value_storage = deepcopy(this_df.loc[n, time_range_vector[y]])
        if this_projection_mode == 'EMPTY':
            continue
        # Create a dictionary for the current iteration
        if this_parameter == 'CapitalCostStorage':
            new_row = {
                'PARAMETER': this_parameter,
                'Scenario': other_setup_params['Main_Scenario'],
                'REGION': other_setup_params['Region'],
                'STORAGE': this_storage,
                'YEAR': time_range_vector[y],
                'Value': this_value_storage
            }
            # Add the dictionary to the list
            capitalcoststorage_data.append(new_row)

        elif this_parameter == 'ResidualStorageCapacity':
            new_row = {
                'PARAMETER': this_parameter,
                'Scenario': other_setup_params['Main_Scenario'],
                'REGION': other_setup_params['Region'],
                'STORAGE': this_storage,
                'YEAR': time_range_vector[y],
                'Value': this_value_storage
            }
            # Add the dictionary to the list
            residualstoragecapacity_data.append(new_row)

if capitalcoststorage_data:  # Ensure there's something to append
    new_capitalcoststorage_df = pd.DataFrame(capitalcoststorage_data)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=FutureWarning)
        df_CapitalCostStorage = pd.concat([df_CapitalCostStorage, new_capitalcoststorage_df], ignore_index=True)
        
if residualstoragecapacity_data:  # Ensure there's something to append
    new_residualstoragecapacity_df = pd.DataFrame(residualstoragecapacity_data)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=FutureWarning)
        df_ResidualStorageCapacity = pd.concat([df_ResidualStorageCapacity, new_residualstoragecapacity_df], ignore_index=True)

#
df_TechnologyToStorage = pd.DataFrame( columns = Wide_Param_Header )
df_TechnologyFromStorage = pd.DataFrame( columns = Wide_Param_Header )

technologytostorage_data = []
technologyfromstorage_data = []

this_df = xtra_storage_technology_sto
df_index_list = this_df.index.tolist()
#
for n in range( len( df_index_list ) ):
    this_tech = deepcopy(this_df.loc[n, 'TECHNOLOGY'])
    this_modeofoperation= deepcopy(this_df.loc[n, 'MODE_OF_OPERATION'])
    this_storage = deepcopy(this_df.loc[n, 'STORAGE'])
    this_parameter_storage = deepcopy(this_df.loc[n, 'Parameter'])
    this_value_storage = deepcopy(this_df.loc[n, 'Value.STORAGE'])
    if this_parameter_storage == 'TechnologyToStorage':
        # Create a dictionary for the current iteration
        new_row = {
            'PARAMETER': this_parameter_storage,
            'Scenario': other_setup_params['Main_Scenario'],
            'REGION': other_setup_params['Region'],
            'MODE_OF_OPERATION': this_modeofoperation,
            'TECHNOLOGY': this_tech,
            'STORAGE': this_storage,
            'Value': this_value_storage
        }
        # Add the dictionary to the list
        technologytostorage_data.append(new_row)

    elif this_parameter_storage == 'TechnologyFromStorage':
        # Create a dictionary for the current iteration
        new_row = {
            'PARAMETER': this_parameter_storage,
            'Scenario': other_setup_params['Main_Scenario'],
            'REGION': other_setup_params['Region'],
            'MODE_OF_OPERATION': this_modeofoperation,
            'TECHNOLOGY': this_tech,
            'STORAGE': this_storage,
            'Value': this_value_storage
        }
        # Add the dictionary to the list
        technologyfromstorage_data.append(new_row)

if technologytostorage_data:  # Ensure there's something to append
    new_technologytostorage_df = pd.DataFrame(technologytostorage_data)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=FutureWarning)
        df_TechnologyToStorage = pd.concat([df_TechnologyToStorage, new_technologytostorage_df], ignore_index=True)
if technologyfromstorage_data:  # Ensure there's something to append
    new_technologyfromstorage_df = pd.DataFrame(technologyfromstorage_data)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=FutureWarning)
        df_TechnologyFromStorage = pd.concat([df_TechnologyFromStorage, new_technologyfromstorage_df], ignore_index=True)



# Update the overall_param_df_dict
overall_param_df_dict['StorageLevelStart'] = df_StorageLevelStart
overall_param_df_dict['OperationalLifeStorage'] = df_OperationalLifeStorage  
overall_param_df_dict['CapitalCostStorage'] = df_CapitalCostStorage
overall_param_df_dict['ResidualStorageCapacity'] = df_ResidualStorageCapacity
overall_param_df_dict['TechnologyToStorage'] = df_TechnologyToStorage  
overall_param_df_dict['TechnologyFromStorage'] = df_TechnologyFromStorage  
# %%
#------------------------------------------------------------------------------
if params['xtra_scen']['Timeslice'] == 'Some':
    print('9 - Include Conversions parameters.')
    #
    df_Conversions = pd.DataFrame( columns = Wide_Param_Header )  
    accumulated_data_by_param_conversion = {}
    
    conversions_params = ['Conversionls', 'Conversionld', 'Conversionlh']
    
    for c in range(len(conversions_params)):
        this_param = conversions_params[c]
        
        if this_param == 'Conversionls':
            set_by_param = 'SEASON'
            name_by_set = 'Season'
        elif this_param == 'Conversionld':
            set_by_param = 'DAYTYPE'
            name_by_set = 'DayType'
        elif this_param == 'Conversionlh':
            set_by_param = 'DAILYTIMEBRACKET'
            name_by_set = 'DailyTimeBracket'
        
        count_values = 0
        for t in range(len(params['xtra_scen']['Timeslices'])):
            for l in range(len(params['xtra_scen'][name_by_set])):
                
                new_row = {
                    'PARAMETER': this_param,
                    'Scenario': other_setup_params['Main_Scenario'],
                    # 'REGION': other_setup_params['Region'],
                    'TIMESLICE': params['xtra_scen']['Timeslices'][t],
                    set_by_param: params['xtra_scen'][name_by_set][l], # Quantity of values
                    'Value': params[this_param][count_values] # List with that quantity
                }
                
                # Accumulate data based on parameter
                if this_param not in accumulated_data_by_param_conversion:
                    accumulated_data_by_param_conversion[this_param] = []
                accumulated_data_by_param_conversion[this_param].append(new_row)
                count_values += 1
    
        # Update the overall_param_df_dict and overall_param_df_dict_ndp
        overall_param_df_dict[this_param] = df_Conversions
        overall_param_df_dict_ndp[this_param] = df_Conversions
    
    # After the loop, convert accumulated data into DataFrames and append them in bulk
    for param, rows in sorted(accumulated_data_by_param_conversion.items()):  # Sort for deterministic order
        if rows:  # Check if there are rows to append for this parameter
            new_rows_df = pd.DataFrame(rows)
            # Check if the parameter already exists in overall_param_df_dict, if not, initialize it
            if param not in overall_param_df_dict:
                overall_param_df_dict[param] = pd.DataFrame(columns=Wide_Param_Header)  
            # Check if the parameter already exists in overall_param_df_dict_ndp , if not, initialize it
            elif param not in overall_param_df_dict_ndp:
                overall_param_df_dict_ndp[param] = pd.DataFrame(columns=Wide_Param_Header)
            overall_param_df_dict[param] = pd.concat([overall_param_df_dict[param], new_rows_df], ignore_index=True)   
            overall_param_df_dict_ndp[param] = pd.concat([overall_param_df_dict_ndp[param], new_rows_df], ignore_index=True) 
            #
        #
    #
# 
end_1 = time.time()   
time_elapsed_1 = -start1 + end_1
print( str( time_elapsed_1 ) + ' seconds /', str( time_elapsed_1/60 ) + ' minutes' )
print('*: For all effects, we have finished the processing tasks of this script. We must now print the results out.')
#***********************************************************************************

#---------------------------------
# Print updated demand DF (user)
writer_Demand_df_new = pd.ExcelWriter(os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Print_Dem_Completed']), engine='xlsxwriter')
Demand_df_new[params['initial_year']] = Demand_df_new[params['initial_year']].astype(float)
Demand_df_new = Demand_df_new.round( 4 )
Demand_df_new.to_excel( writer_Demand_df_new, sheet_name = params['A_O_Dem'], index=False)
writer_Demand_df_new.close()
#---------------------------------
# Print updated *parameterization* DF (user) // this is in Osemosys terms
writer_Param_df = pd.ExcelWriter(os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Print_Paramet_Completed']), engine='xlsxwriter')
param_sheets_print = list( params_dict_new.keys() )
for s in range( len( param_sheets_print ) ):
    this_df_print = params_dict_new[ param_sheets_print[s] ]
    this_df_print = this_df_print.round( 4 )
    this_df_print.to_excel( writer_Param_df, sheet_name = param_sheets_print[s], index=False)
writer_Param_df.close()
#---------------------------------
# Print updated *parameterization* DF (user) // this is in "natural" terms, i.e. the value of each one of the vehicles per unit
writer_Param_Natural_df = pd.ExcelWriter(os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Print_Paramet_Natural_Completed']), engine='xlsxwriter')
param_sheets_print = list( params_dict_new_natural.keys() )
for s in range( len( param_sheets_print ) ):
    this_df_print = params_dict_new_natural[ param_sheets_print[s] ]
    this_df_print = this_df_print.round( 4 )
    this_df_print.to_excel( writer_Param_Natural_df, sheet_name = param_sheets_print[s], index=False)
writer_Param_Natural_df.close()
#---------------------------------
# Print updated 'Activity Ratio' projections
writer_AR_Proj_df = pd.ExcelWriter(os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Print_Proj_Completed']), engine='xlsxwriter')
param_sheets_print = list( AR_Base_proj_df_new.keys() )
for s in range( len( param_sheets_print ) ):
    this_df_print = AR_Base_proj_df_new[ param_sheets_print[s] ]
    this_df_print = this_df_print.round( 4 )
    this_df_print.to_excel( writer_AR_Proj_df, sheet_name = param_sheets_print[s], index=False)
writer_AR_Proj_df.close()
#
#***********************************************************************************

#
#***********************************************************************************
#
lx = [  len( time_range_vector ), len( All_Tech_list ), len( other_setup_params['Timeslices'] ), len( All_Fuel_list ),
        len( emissions_list ), len( other_setup_params['Mode_of_Operation'] ), len( [ other_setup_params['Region'] ] ),
        len( other_setup_params['Season']), len( other_setup_params['DayType'] ),
        len( other_setup_params['DailyTimeBracket'] ), len( other_setup_params['Storage'] )]
mx = max( lx )
df_structure_year = time_range_vector + [ '' for n in range( mx-lx[0] ) ]
df_structure_tech = All_Tech_list + [ '' for n in range( mx-lx[1] ) ]
df_structure_timeslice = other_setup_params['Timeslices'] + [ '' for n in range( mx-lx[2] ) ]
df_structure_fuel = All_Fuel_list + [ '' for n in range( mx-lx[3] ) ]
df_structure_emission = emissions_list + [ '' for n in range( mx-lx[4] ) ]
df_structure_moo = other_setup_params['Mode_of_Operation'] + [ '' for n in range( mx-lx[5] ) ]
df_structure_region = [ other_setup_params['Region'] ] + [ '' for n in range( mx-lx[6] ) ]
df_structure_season = other_setup_params['Season'] + [ '' for n in range( mx-lx[7] ) ]
df_structure_daytype = other_setup_params['DayType'] + [ '' for n in range( mx-lx[8] ) ]
df_structure_dtb = other_setup_params['DailyTimeBracket'] + [ '' for n in range( mx-lx[9] ) ]
df_structure_storage = other_setup_params['Storage'] + [ '' for n in range( mx-lx[10] ) ]
#
df_structure = pd.DataFrame( columns = params['columns4'] )
df_structure['YEAR'] = df_structure_year
df_structure['TECHNOLOGY'] = df_structure_tech
df_structure['TIMESLICE'] = df_structure_timeslice
df_structure['FUEL'] = df_structure_fuel
df_structure['EMISSION'] = df_structure_emission
df_structure['MODE_OF_OPERATION'] = df_structure_moo
df_structure['REGION'] = df_structure_region
df_structure['DAYTYPE'] = df_structure_daytype
df_structure['DAILYTIMEBRACKET'] = df_structure_dtb
df_structure['SEASON'] = df_structure_season
df_structure['STORAGE'] = df_structure_storage # This should be changed when storage is added
writer_Structure_df = pd.ExcelWriter(params['Print_A2_Struct_List'], engine='xlsxwriter')
df_structure.to_excel( writer_Structure_df, sheet_name = params['lists'], index=False)
writer_Structure_df.close()
#
structure_dict = {
    'YEAR': df_structure_year,
    'TECHNOLOGY': df_structure_tech,
    'TIMESLICE': df_structure_timeslice,
    'FUEL': df_structure_fuel,
    'EMISSION': df_structure_emission,
    'MODE_OF_OPERATION': df_structure_moo,
    'REGION': df_structure_region,
    'DAYTYPE': df_structure_daytype,
    'DAILYTIMEBRACKET': df_structure_dtb,
    'SEASON': df_structure_season,
    'STORAGE': df_structure_storage
}
#
#***********************************************************************************
#
# Clean dicts
overall_param_df_dict = {
    k: v for k, v in overall_param_df_dict.items() if not pd.isna(k)
}
overall_param_df_dict_ndp = {
    k: v for k, v in overall_param_df_dict.items() if not pd.isna(k)
}
#
list_dicts = list( overall_param_df_dict.keys() )

path_main_scenario = os.path.join(params['A2_output_main_scen'], params['xtra_scen']['Main_Scenario'])

os.makedirs(path_main_scenario, exist_ok=True)

list_dicts = list(overall_param_df_dict.keys())
for d in range(len(list_dicts)):
    df_to_print = overall_param_df_dict[list_dicts[d]]    
    df_to_print.to_csv(
        os.path.join(path_main_scenario, list_dicts[d] + '.csv'),
        index=False,
        header=True
    )

for col_name, data_list in sorted(structure_dict.items()):  # Sort for deterministic order
    df = pd.DataFrame({'VALUE': data_list})
    file_path = os.path.join(path_main_scenario, f'{col_name}.csv')
    df.to_csv(file_path, index=False, header=True)

# Check if there are additional scenarios defined
if other_setup_params['Other_Scenarios']:
    
    # Loop through each additional scenario
    for scenario in other_setup_params['Other_Scenarios']:
        
        # Create the output folder for this scenario
        output_path = os.path.join(params['A2_output'], scenario)
        os.makedirs(output_path, exist_ok=True)

        # Process and save DataFrames for this scenario
        for name, df in sorted(overall_param_df_dict_ndp.items()):  # Sort for deterministic order
            # Replace the main scenario name with the current scenario
            df_scenario = df.replace({
                'Scenario': {
                    other_setup_params['Main_Scenario']: scenario
                }
            })
            # Save the modified DataFrame as a CSV file
            df_scenario.to_csv(
                os.path.join(output_path, f'{name}.csv'),
                index=False,
                header=True
            )

        # Save structure reference lists as individual CSVs
        for col_name, data_list in sorted(structure_dict.items()):  # Sort for deterministic order
            df_structure = pd.DataFrame({'VALUE': data_list})
            df_structure.to_csv(
                os.path.join(output_path, f'{col_name}.csv'),
                index=False,
                header=True
            )
#
#
end_2 = time.time()   
time_elapsed_2 = -start1 + end_2
print( str( time_elapsed_1 ) + ' seconds /', str( time_elapsed_2/60 ) + ' minutes' )
print('*: We just finished the printing of the results.')
# # %%
# # Print important pickles below:
# with open( os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Pickle_Fleet_Groups_Dist']), 'wb') as handle1:
#     pickle.dump(Fleet_Groups_Distance, handle1, protocol=pickle.HIGHEST_PROTOCOL)
# with open( os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Pickle_Fleet_Groups_OR']), 'wb') as handle2:
#     pickle.dump(Fleet_Groups_OR, handle2, protocol=pickle.HIGHEST_PROTOCOL)
# with open( os.path.join(params['A1_outputs'],params['A1_outputs'] + '_' + params['xtra_scen']['Main_Scenario'] + params['Pickle_Fleet_Groups_T2D']), 'wb') as handle3:
#     pickle.dump(Fuels_techs_2_dems, handle3, protocol=pickle.HIGHEST_PROTOCOL)
# #