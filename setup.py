""" set up file for this library """
from setuptools import setup

setup(
    name='convert2pdf',
    version='0.0.1',
    description="""A Python3 application that converts multiple Office files into their PDF
                   versions automatically thus saving you the hassle of looking for online 
                   converters or converting them manually.""",
    url='https://github.com/aditeyabaral/Convert2PDF.git',
    author='aditeyabaral',
    license='unlicense',
    packages=['convert2pdf'],
    install_requires=[
        'img2pdf',
        'comtypes'
    ],
    zip_safe=False
)
