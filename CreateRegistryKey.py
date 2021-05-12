##########################################################################################################
# Create registry key which displays "Convert to PDF" on right-click context menu                        #
# File:     CreateRegistryKey.py                                                                         #
# Authors:  Antoine Dery                                                                                 #
# Thanks to Stephen Edwards for the file (https://github.com/seddie95/python_script_in_right_click_menu) #
########################################################################################################## 

import os
import sys
import winreg as reg

cwd = os.getcwd()                                                                       # Get current folder path
python_exe = sys.executable                                                             # Create executable

key_path = r'SOFTWARE\\Classes\\Word.Document.12\\shell\\Convert to PDF'                # Path to add registry key

key = reg.CreateKeyEx(reg.HKEY_CURRENT_USER, key_path)                                  # Create key named 'Convert to PDF'
reg.SetValue(key, '', reg.REG_SZ, '&Convert to PDF')                                    # Set value of key which is the displayed message on right click

sub_key = reg.CreateKey(key, r"command")                                                # Create subkey named 'command' 
reg.SetValue(sub_key, '', reg.REG_SZ, python_exe + f' "{cwd}\\PDFconverter.py" \"%1\"') # Link PDFConverter.py as value which execute the python script when clicked