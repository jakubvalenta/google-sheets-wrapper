from codecs import open
from os import path

from setuptools import find_packages, setup

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='google-sheets-wrapper',
    version='1.0.0',
    description=(
        'A library wrapping Google Sheets API to make some operations' 'easier'
    ),
    long_description=long_description,
    url='https://github.com/jakubvalenta/google-sheets-wrapper',
    author='Jakub Valenta',
    author_email='jakub@jakubvalenta.cz',
    license='Apache Software License',
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: Apache Software License',
        'Topic :: Software Development',
        'Programming Language :: Python :: 3',
    ],
    keywords='',
    packages=find_packages(),
    install_requires=['google-api-python-client'],
)
