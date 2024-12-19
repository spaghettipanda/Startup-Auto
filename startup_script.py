# Custom Modules
from shift_calc import *
from interface_input import *
from file_management import *

# [REMINDER: User \\ doubleslash for subfolders]

try:    
    # Get current date-time
    now = get_now()
    username = get_username()

    # Get Shift details as a string
    day_of_week = get_day_of_week(now)                  # Monday/Tuesday/Wednesday/Thursday/Friday/Saturday/Sunday
    today_date = get_date(now)                          # Today's date
    weekend_status = get_weekend_status(now)            # Weekend/Weekday
    day_or_night_shift = get_day_or_night_shift(now)    # Day/Night shift
    now_time = convert_to_time(now)                     # Current time

    # Add all the details into a single message for current the current shift
    shift = f'Day: {day_of_week} \nDate: {today_date} \nTime: {now_time} \n\nShift: {weekend_status} {day_or_night_shift} \n'
    
    home_directory = f'C:\\Users\\{username}'
    cfg_file_name = f'startup_config'
    cfg_dir = f'{home_directory}\\{cfg_file_name}'
    
    print('\nStarting script...')    
    print('\n======================================')
    print(colored('Use CTRL+C in the console window to cancel the script...', 'black', 'on_light_grey'))
    print('======================================\n')

    init = False

    # Apply based on settings
    if(file_exists(cfg_dir)):
        is_dashboards = str_to_bool(config_scan(cfg_dir, 'Dashboards'))
        print(is_dashboards)
        is_SAP = str_to_bool(config_scan(cfg_dir, 'SAP'))
        print(is_SAP)
        is_MOP = str_to_bool(config_scan(cfg_dir, 'MOP'))
        print(is_MOP)
        is_MonitorConfig = str_to_bool(config_scan(cfg_dir, 'MonitorConfig'))
        print(is_MonitorConfig)
        is_ExtraCustomConfig = str_to_bool(config_scan(cfg_dir, 'ExtraCustomConfig'))
        print(is_ExtraCustomConfig)
    else:
        message_box('O', f'A config file was not found...\n\nCreating one at {home_directory}', f'First Time?')
        print(f'Config does not exist\nCreating...')
        f = open(f'{cfg_dir}'.split('=')[0], 'w') # Create Config File
        f.write('# Configuration File\n')
        
        is_dashboards = message_box('Y/N', f'Do you have Dashboard shortcuts?', f'Dashboards')
        f.write(f'Dashboards={is_dashboards}\n')

        is_SAP = message_box('Y/N', f'Do you have after-hours SAP shortcuts? (PAS ONLY)', f'TROC Apps SAP')
        f.write(f'SAP={is_SAP}\n')

        is_MOP = message_box('Y/N', f'Do you have L1 MOP shortcuts?', f'L1 MOP')
        f.write(f'MOP={is_MOP}\n')

        is_MonitorConfig = message_box('Y/N', f'Do you want monitor configuration automation?\n\nNOTE: Requires MultiMonitorTool.exe; will be downloaded automatically if you select Yes.\n\nAlso requires some set-up (i.e. saving your current monitor configuration)\n\nSee QRG for more info...', f'Monitor Configuration')
        f.write(f'MonitorConfig={is_MonitorConfig}\n')
        
        is_ExtraCustomConfig = False
        f.write(f'ExtraCustomConfig={False}\n')

        f.close()
        init = True

    
    print(colored(f'Configuration Settings Found...\n', 'green', attrs=["bold"]))
    print(colored(f'Dashboard shortcuts', f'{bool_color(is_dashboards)}', attrs=["bold", "underline"]), ':', colored(f'{is_dashboards}', f'{bool_color(is_dashboards)}'))
    print(colored(f'SAP shortcuts:', f'{bool_color(is_SAP)}', attrs=["bold", "underline"]), colored(f'{is_SAP}', f'{bool_color(is_SAP)}'))
    print(colored(f'MOP shortcuts:', f'{bool_color(is_MOP)}', attrs=["bold", "underline"]), colored(f'{is_MOP}', f'{bool_color(is_MOP)}'))
    print(colored(f'Monitor Configs:', f'{bool_color(is_MonitorConfig)}', attrs=["bold", "underline"]), colored(f'{is_MonitorConfig}', f'{bool_color(is_MonitorConfig)}'))
    if(is_ExtraCustomConfig):
        print(colored(f'Extra Custom Config:', f'{bool_color(is_ExtraCustomConfig)}', attrs=["bold", "underline"]), colored(f'{is_ExtraCustomConfig}', f'{bool_color(is_ExtraCustomConfig)}'))
    
    print(colored(f'\n======================================', 'light_yellow', attrs=["dark"]))
    print(colored(f'\nREMINDER: \n\nTo change the script settings, edit the file directly or delete it at:\n', 'yellow'), colored(f'{cfg_dir}', 'light_yellow'))
    print(colored(f'\n======================================\n', 'light_yellow', attrs=["dark"]))

    # Directory variabes
    shortcut_path = f'{home_directory}\\Documents\\StartupShortcuts'# Path shortcut folder

    main = 'Main Shortcuts'                         # Main shortcut folder
    dashboard = 'Dashboard Shortcuts'               # Dashboard shortcut folder
    sap = 'SAP Shortcuts'                           # SAP shortcut folder
    mop = 'L1_MOP Shortcuts'                        # MOP shortcuts

    monitor_configs = 'MonitorConfig'  # Monitor configurations folder (Requires Multi Monitor Tool - See ### Below)
    extra_custom_config = 'ExtraCustomConfig' # Hidden - Extra Config file if needed
    monitor_setup = True

    custom_config = 'CustomConfig.cfg'
    default_config = 'MultiMonitorTool.cfg'

    # Append path to folder directories
    main = f'{shortcut_path}\\{main}'
    dashboards = f'{shortcut_path}\\{dashboard}'
    sap = f'{shortcut_path}\\{sap}'
    mop = f'{shortcut_path}\\{mop}'
    monitor_configs = f'{shortcut_path}\\{monitor_configs}'
    
    ### MULTI MONITOR TOOL ###
    multi_monitor_tool = f'{monitor_configs}\\MultiMonitorTool.exe' # https://www.nirsoft.net/utils/multi_monitor_tool.html (https://www.nirsoft.net/utils/multimonitortool-x64.zip)


    ## 1) Check SAP PM02 on night shifts and weekend shifts
    if(is_SAP):
        if(day_or_night_shift == 'Night Shift') or (weekend_status == 'Weekend'):
            
            # SAP Reminder
            sap_message = f'\n-------------------------------------\nREMINDER: SAP PM02 Today!\n-------------------------------------\n'
            shift = f'{shift} {sap_message}' 
    
    ## 2) Check if L1_MOP occurring on second shift of swing (Tuesday/Thursday/Saturday Day Shifts)
    if(is_MOP):
        if(day_or_night_shift == 'Day Shift'):
            if day_of_week in ('Tuesday', 'Thursday', 'Saturday'):
                
                # MOP Reminder
                mop_message = f'\n----------------------------------\nREMINDER: L1 MOP Today!\n----------------------------------\n'
                shift = f'{shift} {mop_message}'
    
    # Check if Multi Monitor Tool exists Prompt tool download if not
    if(is_MonitorConfig):
        if(not os.path.isfile(multi_monitor_tool)):
            zip = f'{monitor_configs}\\multimonitortool-x64.zip'
            if(os.path.isfile(zip)):
                extract_zip(zip, monitor_configs, 'MultiMonitorTool.exe')
                remove_file(zip)
            else:
                print(f'Downloading MultiMonitorTool.exe to: \n\n{monitor_configs}\n')
                create_folder(monitor_configs)
                download_to_dir('https://www.nirsoft.net/utils/multimonitortool-x64.zip', monitor_configs)
                extract_zip(zip, monitor_configs, 'MultiMonitorTool.exe')
                remove_file(zip)

        # Check multi monitor exists before proceeding
        if(not os.path.isfile(multi_monitor_tool)):
            monitor_setup = False

    # Show shift details
    print(colored(f'\n{shift}\n','red'))
    shift_details = message_box('O', shift, f'Hello! {get_data(3)}')
    
    ## 2) Monitor Configuration
    if(is_ExtraCustomConfig):
        print(colored(f'\n======================================', 'light_blue', attrs=["dark"]))
        print(colored(f'\nMONITOR CONFIGURATION\n', 'light_blue', attrs=["bold"]))
        # Configure monitor setup
        if(monitor_setup == True):
            custom = message_box('Y/N/C', 'Custom Config?', f'Hello! {get_data(3)}')
            if(custom==True):           # Yes - Custom Monitor Config
                if(not os.path.isfile(f'{monitor_configs}\\{custom_config}')):
                    message_box('error', f"""Monitor config file is missing for {custom_config}... 
                                    \n\nPlease save a configuration file as {custom_config} through the Multi Monitor application
                                    \n\nSave the config file to {monitor_configs}
                                    \n\nSee QRG for more info...""", f'Custom Monitor Configuration missing...')
                    monitor_setup = False
                else:
                    print(colored('\nApplying Custom Monitor Config...\n', 'light_green'))
                    monitor = execute_command(f'{multi_monitor_tool} /LoadConfig {monitor_configs}\\{custom_config}')
            elif(custom==False):        # No - Default Monitor Config
                if(not os.path.isfile(f'{monitor_configs}\\{default_config}')):
                    message_box('error', f"""Monitor config file is missing for default monitor configuration... 
                                    \n\nPlease save a configuration file as {default_config} through the Multi Monitor application
                                    \n\nSave the config file to {monitor_configs}
                                    \n\nSee QRG for more info...""", f'Default Monitor Configuration missing...')
                    monitor_setup = False
                else:
                    print(colored('\nApplying Default Monitor Config...\n', 'light_green'))
                    monitor = execute_command(f'{multi_monitor_tool} /LoadConfig {monitor_configs}\\{default_config}')
            else:                       # Cancelled - Apply no config
                print('\nKeeping current Monitor Config...\n') 
    else:
        if(is_MonitorConfig):
            if(not os.path.isfile(f'{monitor_configs}\\{default_config}')):
                message_box('error', f"""Monitor config file is missing for default monitor configuration... 
                                \n\nPlease save a configuration file as {default_config} through the Multi Monitor application
                                \n\nSave the config file to {monitor_configs}""", f'Default Monitor Configuration missing...')
            else:
                print(colored('\nApplying Default Monitor Config...\n', 'light_green'))
                monitor = execute_command(f'{multi_monitor_tool} /LoadConfig {monitor_configs}\\{default_config}')
            

    print(colored(f'\n======================================\n', 'light_blue', attrs=["dark"]))

    ## 3) Check if L1_MOP occurring on second shift of swing (Tuesday/Thursday/Saturday Day Shifts)
    if(is_MOP):
        if(day_or_night_shift == 'Day Shift'):
            if day_of_week in ('Tuesday', 'Thursday', 'Saturday'):
                
                # Traverse MOP directory
                traverse_dir(mop, init)
                print('')

    ## 3) Traverse main directory and open shortcuts
    traverse_dir(main, init)
    print('')

    ## 4) Traverse dashboards directory and open shortcuts
    if(is_dashboards):
        traverse_dir(dashboards, init)
        print('')

    ## 5) Traverse SAP directory and open shortcuts
    if(is_SAP):
        if(day_or_night_shift == 'Night Shift') or (weekend_status == 'Weekend'):    
            traverse_dir(sap, init)
            print('')

    # Script End
    print('Execution Finished')
    input(colored('Press any key to exit...', 'cyan' , attrs=["reverse"]))

# Exception catching
except ValueError as err:
    print(f'\ntype: {sys.exc_info()} \nerror:{err}')
    print(f'args: {err.args}')
    sys.exit()

# CTRL + C Pressed (KeyboardInterrupt)
except KeyboardInterrupt as err:
    print(colored(f'\n======================================', attrs=["reverse"]))
    print(colored(f'CTRL+C pressed! Exiting...............', attrs=["reverse"]))
    print(colored(f'======================================\n', attrs=["reverse"]))
    sys.exit()

except AttributeError as err:
    print(f'\ntype: {sys.exc_info()} \nerror:{err}')
    print(f'args: {err.args}')
    print(f'Line {sys.exc_info().tb_lineno}')
    print(f'\nIf you see this, it may be an issue with your startup_config file. Please delete it and try again... \n\nConfig Location: {cfg_dir}')
    sys.exit()

except Exception as err:
    print(f'\ntype: {sys.exc_info()} \nerror:{err}')
    print(f'args: {err.args}')
    sys.exit()