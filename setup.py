# -*- coding: UTF-8 -*-
#.Data:2018/5/19
from setuptools import setup

setup(name='Python_Vector_CANoe',
      version='0.1',
      description='Control Vector CANoe API by Python',
      url='https://github.com/hmq2018/Python-Vector-CANoe',
      author='Hz',
      author_email='weld83@126.com',
      license='MIT',
      packages=['Python_Vector_CANoe'],
      install_requires=[
          'pywin32==301',
      ],
      zip_safe=False)