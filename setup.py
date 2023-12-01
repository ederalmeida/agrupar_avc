import cx_Freeze
from apoio import versao_local as vl
import hashlib
import os.path

build_exe_options = {
'include_msvcr': True
}

exe = [cx_Freeze.Executable('agrupar_avc.py',
                            base = 'Win32GUI',
                            target_name = 'agrupar_avc.exe')]

cx_Freeze.setup(name = 'agrupar_avc',
                version=vl.v,
                options = {'build_exe': build_exe_options},
                executables = exe
)