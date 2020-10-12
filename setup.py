""" set up file for this library """
import sys
from setuptools import setup

if sys.platform.startswith('win32'):
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
else:
    print('Sorry, your os does not support this library!')
    sys.exit(1)
