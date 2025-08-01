from setuptools import setup, find_packages

setup(
    name='financial_data_parser',
    version='0.1',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=[
        'pandas',
        'openpyxl',
        'numpy'
    ]
)