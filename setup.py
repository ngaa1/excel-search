from setuptools import setup, find_packages

setup(
    name='excel-search',
    version='1.0.0',
    description='Excel文件内容搜索工具',
    author='Your Name',
    packages=find_packages(),
    install_requires=[
        'openpyxl',
        'xlrd',
        'fuzzywuzzy',
        'python-Levenshtein'
    ],
    entry_points={
        'console_scripts': [
            'excel-search=main:main'
        ]
    }
)