# -*- coding: utf-8 -*-
from distutils.core import setup
from setuptools import find_packages
import datetime

__version__ = "${VERSION}"
if "$" in __version__:
    __version__ = datetime.datetime.now().strftime("%Y.%m.%d.%H.%M")
print("Version:", __version__)

setup(
    name="beautifulexcel",
    version=__version__,
    description="BeautifulExcel is a python package that makes it easy and quick to save pandas dataframes in beautifully formatted excel files. BeautifulExcel is the Openpyxl for Data Scientists with a deadline.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    author="https://github.com/vanalmsick",
    url="https://github.com/vanalmsick/beautifulexcel",
    project_urls={
        "Issues": "https://github.com/vanalmsick/beautifulexcel/issues",
        "Documentation": "https://vanalmsick.github.io/beautifulexcel/",
        "Source Code": "https://github.com/vanalmsick/beautifulexcel",
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    packages=find_packages(),
    package_data={'beautifulexcel': ['*.yml', 'VERSION']}
    include_package_data=True,
)
