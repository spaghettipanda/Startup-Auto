# Custom Modules
from shift_calc import *
from interface_input import *
from directory_traversal import *

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
    
    init = False

    print('\nStarting script...')    
    print('\n======================================')
    print(colored('Use CTRL+C in the console window to cancel the script...', 'cyan'))
    print('======================================\n')

    # Apply based on settings
    if(file_exists(cfg_dir)):
        is_dashboards = str_to_bool(config_scan(cfg_dir, 'Dashboards'))
        is_SAP = str_to_bool(config_scan(cfg_dir, 'SAP'))
        is_MOP = str_to_bool(config_scan(cfg_dir, 'MOP'))
        is_MonitorConfigs = str_to_bool(config_scan(cfg_dir, 'MonitorConfigs'))
    else:
        message_box('O', f'A config file was not found...\n\nCreating one at {home_directory}', f'First Time?')
        print(f'Config does not exist\nCreating...')
        f = open(f'{cfg_dir}'.split('=')[0], 'w') # Create Config File
        f.write('# Configuration File\n')
        
        is_dashboards = message_box('Y/N', f'Do you have Dashboard shortcuts?', f'Dashboards')
        f.write(f'Dashboards={is_dashboards}\n')

        is_SAP = message_box('Y/N', f'Do you have after-hours SAP shortcuts (Weekends + Nights)?', f'After-hours SAP')
        f.write(f'SAP={is_SAP}\n')

        is_MOP = message_box('Y/N', f'Do you have L1 MOP shortcuts?', f'L1 MOP')
        f.write(f'MOP={is_MOP}\n')

        is_MonitorConfigs = message_box('Y/N', f'Do you want monitor configuration automation?', f'Monitor Configuration')
        f.write(f'MonitorConfigs={is_MonitorConfigs}\n')

        f.close()
        init = True
    

    print(f'Configuration Settings Found...\n')
    print(colored(f'Dashboard shortcuts:', 'white'), colored(f'{is_dashboards}', f'{bool_color(is_dashboards)}'))
    print(colored(f'SAP shortcuts:', 'white'), colored(f'{is_SAP}', f'{bool_color(is_SAP)}'))
    print(colored(f'MOP shortcuts:', 'white'), colored(f'{is_MOP}', f'{bool_color(is_MOP)}'))
    print(colored(f'Monitor Configs:', 'white'), colored(f'{is_MonitorConfigs}', f'{bool_color(is_MonitorConfigs)}'))
    
    print('\n======================================')
    print(colored(f'REMINDER: \n\nTo change the script settings, edit the file directly or delete it at:\n', 'yellow'), colored(f'{cfg_dir}', 'light_yellow'))
    print('======================================\n')

    # Directory variabes
    shortcut_path = f'{home_directory}\\Documents\\StartupShortcuts'# Path shortcut folder

    main = 'Main Shortcuts'                         # Main shortcut folder
    dashboard = 'Dashboard Shortcuts'               # Dashboard shortcut folder
    sap = 'SAP Shortcuts'                           # SAP shortcut folder
    mop = 'L1_MOP Shortcuts'                        # MOP shortcuts

    monitor_configs = 'MonitorConfigs'  # Monitor configurations folder (Requires Multi Monitor Tool - See ### Below)
    monitor_setup = True

    custom_config = 'CustomConfig.cfg'
    default_config = 'DefaultConfig.cfg'

    # Append path to folder directories
    main = f'{shortcut_path}\\{main}'
    dashboards = f'{shortcut_path}\\{dashboard}'
    sap = f'{shortcut_path}\\{sap}'
    mop = f'{shortcut_path}\\{mop}'
    monitor_configs = f'{shortcut_path}\\{monitor_configs}'
    
    ### MULTI MONITOR TOOL ###
    multi_monitor_tool = f'{monitor_configs}\\MultiMonitorTool.exe' # https://www.nirsoft.net/utils/multi_monitor_tool.html (https://www.nirsoft.net/utils/multimonitortool-x64.zip)

    ## 1) Traverse main directory
    traverse_dir(main)
    print('')

    ## 2) Traverse dashboards directory
    if(is_dashboards):
        traverse_dir(dashboards)
        print('')

    ## 3) SAP PM02 on night shifts and weekend shifts
    if(is_SAP):
        if(day_or_night_shift == 'Night Shift') or (weekend_status == 'Weekend'):
            
            # SAP Reminder
            sap_message = f'\n-------------------------------------\nREMINDER: SAP PM02 Today!\n-------------------------------------\n'
            shift = f'{shift} {sap_message}' 
            
            # Traverse SAP directory
            traverse_dir(sap)
            print('')

    ## 4) For L1_MOP occurring on second shift of swing (Tuesday/Thursday/Saturday Day Shifts)
    if(is_MOP):
        if(day_or_night_shift == 'Day Shift'):
            if day_of_week in ('Tuesday', 'Thursday', 'Saturday'):
                
                # MOP Reminder
                mop_message = f'\n----------------------------------\nREMINDER: L1 MOP Today!\n----------------------------------\n'
                shift = f'{shift} {mop_message}'
                
                # Traverse MOP directory
                traverse_dir(mop)
                print('')
            else:
                if(init):
                    check_dir(mop)
                    print('')
                else:
                    shift = shift + '----------'
    
    # Prompt tool download if Multi Monitor Tool isn't found
    if(is_MonitorConfigs):
        if(not os.path.isfile(multi_monitor_tool)):
            zip = f'{monitor_configs}\\multimonitortool-x64.zip'
            if(os.path.isfile(zip)):
                extract_zip(zip, monitor_configs, 'MultiMonitorTool.exe')
                remove_file(zip)
            else:
                download = message_box('Y/N', f'MultiMonitorTool.exe was not found at: \n\n{monitor_configs} \n\nDownload?', f'Missing Multi Monitor Tool...')
                if(download):
                    create_folder(monitor_configs)
                    download_to_dir('https://www.nirsoft.net/utils/multimonitortool-x64.zip', monitor_configs)
                    extract_zip(zip, monitor_configs, 'MultiMonitorTool.exe')
                    remove_file(zip)
                    
                else:
                    message_box('warning', f'MultiMonitorTool is needed to for monitor configuration...', f'Missing Multi Monitor Tool...')

        # Check multi monitor exists before proceeding
        if(not os.path.isfile(multi_monitor_tool)):
            monitor_setup = False

    # Show shift details
    shift_details = message_box('O', shift, f'Hello! {get_data(3)}')

    ## 5) Monitor Configuration
    if(is_MonitorConfigs):
        print('\n======================================\nMONITOR CONFIGURATION\n')
        # Check config files exists
        if(not os.path.isfile(f'{monitor_configs}\\{custom_config}')):
            monitor_setup = False
            message_box('error', f"""Monitor config file is missing for {custom_config}... 
                            \n\nPlease save a configuration file as {custom_config} through the Multi Monitor application
                            \n\nSave the config file to {monitor_configs}""", f'Custom Monitor Configuration missing...')
        if(not os.path.isfile(f'{monitor_configs}\\{default_config}')):
            monitor_setup = False
            message_box('error', f"""Monitor config file is missing for {default_config}... 
                            \n\nPlease save a configuration file as {default_config} through the Multi Monitor application
                            \n\nSave the config file to {monitor_configs}""", f'Default Monitor Configuration missing...')
        
        # Configure monitor setup
        if(monitor_setup == True):
            custom = message_box('Y/N/C', 'Laptop?', f'Hello! {get_data(3)}')
            if(custom==True):           # Yes - Custom Monitor Config
                print(colored('\nApplying Laptop Monitor Config...\n', 'light_green'))
                monitor = execute_command(f'{multi_monitor_tool} /LoadConfig {monitor_configs}\\{custom_config}')
            elif(custom==False):        # No - Default Monitor Config
                print(colored('\nApplying Default Monitor Config...\n', 'light_green'))
                monitor = execute_command(f'{multi_monitor_tool} /LoadConfig {monitor_configs}\\{default_config}')
            else:                       # Cancelled - Apply no config
                print('\nKeeping current Monitor Config...\n') 

    # Script End
    print('Execution Finished')
    input('Press any key to exit...')

# Exception catching
except ValueError as err:
    print(f'\ntype: {sys.exc_info()} \nerror:{err}')
    print(f'args: {err.args}')
    sys.exit()

# CTRL + C Pressed (KeyboardInterrupt)
except KeyboardInterrupt as err:
    print('\n======================================\n')
    print(f'CTRL+C pressed! Exiting...')
    print('\n======================================\n')
    sys.exit()

except Exception as err:
    print(f'\ntype: {sys.exc_info()} \nerror:{err}')
    print(f'args: {err.args}')
    sys.exit()