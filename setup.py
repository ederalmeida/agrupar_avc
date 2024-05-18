import cx_Freeze

build_exe_options = {
'include_msvcr': True,
"excludes": ["unittest"],
    "zip_include_packages": ["encodings"]
}

exe = [cx_Freeze.Executable('agrupar_avc.py',
                            base = 'Win32GUI',
                            target_name = 'agrupar_avc.exe')]

cx_Freeze.setup(name = 'agrupar_avc',
                options = {'build_exe': build_exe_options},
                executables = exe
)