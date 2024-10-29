import cx_Freeze
import sys
import os

base = None

if sys.platform == 'win32':
    base = 'Win32GUI'  # Use this for a GUI application, no console window

executables = [
    cx_Freeze.Executable(
        script='test.py',  # The main Python script
        base=base,
        icon='icon.ico'  # Specify your .ico file here
    )
]

cx_Freeze.setup(
    name='WresConvert1',  # Updated app name
    options={'build_exe': {
        'packages': ['os', 'tkinter', 'PIL', 'PyPDF2'], 
        'include_files': ['app_icon.webp', 'cropped-wrl-logo-new.png'],
        'build_exe': 'build/WresConvert1'  # Output folder where WresConvert1.exe will be saved
    }},
    executables=executables
)
