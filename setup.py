'''
Welcome to PyInventor! This is a Python wrapper for the Autodesk Inventor API
which is natively written in VBA. This package ONLY works on windows machines 
(or MacOS running windows in bootcamp) and will work on Inventor 2017 or later 
(although it is most thoroughly tested on Inventor 2019, which is recommended).
This package requires no dependencies outside of Python 3 through the Anaconda
distribution. This package can only create 3D parts (no assemblies) and still
lacks some of the 3D functionality (no lofts, 3D sketches, or chamfers, among 
others). To see the full functionality check the demos in the _Tutorial_Notebook
folder. Have fun and shoot me an email with questions:

~Andrew Oriani

'''

from pathlib import Path
from setuptools import setup, find_packages

here = Path(__file__).parent.absolute()

# Get the long description from the README file
with open(here / "README.md", encoding="utf-8") as f:
    long_description = f.read()

with open(here / "requirements.txt", encoding="utf-8") as f:
    requirements = f.read().splitlines()

doclines = __doc__.split('\n')

setup(name='PyInventor',
      version='0.4',
      description = doclines[0],
      long_description=long_description,
      long_description_content_type="text/markdown",
      author='Andrew E. Oriani',
      packages=find_packages(),
      author_email='oriani@uchicago.edu',
      maintainer='Andrew Oriani SchusterLab',
      license='BSD-3-Clause',
      classifiers=[
        "Intended Audience :: Developers",
        "Intended Audience :: Science/Research",
        "Operating System :: Microsoft :: Windows::Only",
        "Programming Language :: Python :: 3 :: Only",
        "Programming Language :: Python :: 3.5",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Topic :: Scientific/Engineering",
        "Environment :: Console"],
      python_requires=">=3.5, <4",
      # install_requires=['numpy', 'IPython'],
      install_requires=requirements
      )
