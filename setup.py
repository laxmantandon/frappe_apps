# -*- coding: utf-8 -*-
from setuptools import setup, find_packages

with open('requirements.txt') as f:
	install_requires = f.read().strip().split('\n')

# get version from __version__ variable in ssipl_import/__init__.py
from ssipl_import import __version__ as version

setup(
	name='ssipl_import',
	version=version,
	description='Stock Entry Import with required masters',
	author='Laxman',
	author_email='laxmantandon@gmail.com',
	packages=find_packages(),
	zip_safe=False,
	include_package_data=True,
	install_requires=install_requires
)
