import sys
import os
from PyInstaller.__main__ import run


# Specify the entry point of your application
entry_point = 'cgm_icon.py'

# Specify the name of the output directory where the built executable will be saved
output_dir = 'dist'

# Specify any additional PyInstaller options you want to use
pyinstaller_options = [
    '--onefile',        # Create a single executable file
    '--noconsole',      # Run the application without a console window
    '--icon=img/icon.ico',  # Specify the path to your application icon (replace with your own icon file)
    '--add-data=img/*.png;img',  # Include the image files in the 'img' directory
    # '--add-data=arial.ttf;.',     # Include the 'arial.ttf' font file in the current directory
]


# Create the PyInstaller command
command = [
    entry_point,
    *pyinstaller_options
]

# Run PyInstaller
run(command)

# Move the output file to the specified output directory
if sys.platform == 'win32':
    exe_name = entry_point[:-3] + '.exe'
    os.makedirs(output_dir, exist_ok=True)
    os.replace(exe_name, os.path.join(output_dir, exe_name))
else:
    print('Unsupported platform. Please build the executable manually.')
