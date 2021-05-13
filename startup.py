import os
import time
import ctypes
from distutils.dir_util import copy_tree

print('Checking for updates...')
foreign_address = \
    os.path.expanduser('~').replace('\\', '/') + '/Barrett Benefits Group, Inc/BBG Admin - Shared/CenFill/updates'

# Check to see if the directory is there
if os.path.exists(foreign_address):
    # Acquire both the current version and the latest version in the shared folder
    foreign_check = open(foreign_address + '/version.txt', 'r')
    home_check = open('C:/Program Files/CenFill/version.txt', 'r')
    current_version = str(home_check.readline())
    new_version = str(foreign_check.readline())
    foreign_check.close()
    home_check.close()
    if new_version != current_version:
        print('Update found for version ' + new_version)
        if ctypes.windll.shell32.IsUserAnAdmin():
            print('Updating...')
            copy_tree(foreign_address, 'C:/Program Files/CenFill')
            print('Update complete! Exiting now... ')
        else:
            print('Please run CenFill again as an administrator to update. Program will exit now... ')
        time.sleep(4)
        exit()
    else:
        print('Update to date! Proceeding as normal...')
        import main


else:
    print('Directory not found. Proceeding as normal...')
    import main
