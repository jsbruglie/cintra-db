from distutils.core import setup
import py2exe, sys

from shutil import copyfile

sys.argv.append('py2exe')

includes = []

setup(
    options = {'py2exe': {  'compressed':2,
                            'optimize': 2,
                            'includes':includes}},
    # comment if console output is desired
    windows = [{'script': "src\csv-tools.py"}],
    zipfile = None,
)

# Copy logo file
copyfile("src\logo.ico", "dist\logo.ico")
# Copy default preproc file
copyfile("src\preproc.csv", "dist\preproc.csv")
