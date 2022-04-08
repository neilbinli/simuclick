import os
import pathlib
import re

from setuptools import setup, find_packages

setup(
    name='simuclick',
    version="0.1",
    description='',
    url='',
    packages=find_packages(include=['src']),
    python_requires='==3.7.6',
    install_requires=[
        'pyperclip==1.8.2',
        'xlrd==2.0.1',
        'pyautogui==0.9.53',
        'opencv-python==4.5.5.64',
        'pillow==9.1.0',
        'pydirectinput==1.0.4'
    ],
)
