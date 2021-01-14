##############################################################################
# Python header

__author__ = "Andrew Oriani"
__copyright__ = "Copyright 2020"
__credits__ = ["Andrew E. Oriani"]
__license__ = "BSD-3-Clause"
__version__ = "0.4"
__maintainer__ = "Andrew Oriani Schusterlab"
__email__ = "oriani@uchicago.edu"
__url__ = r'schusterlab.uchicago.edu'
__status__ = "Dev-Production"

##############################################################################
#Checks module compatibility
try:
    import glob
    del glob
except (ImportError, ModuleNotFoundError):
    print('Glob module not installed, install glob module using: $ conda  install -c conda-forge glob ')

try:
    import xlwings
    del xlwings
except (ImportError, ModuleNotFoundError):
    print('Xlwings module not installed, install glob module using: $ conda  install -c conda-forge xlwings ')

from . import  pyinvent
from .pyinvent import com_obj, structure, iPart, arc_pattern, circle_pattern 

__all__=['pyinvent', 'com_obj', 'structure', 'iPart', 'arc_pattern', 'circle_pattern']
