from pathlib import Path

from setuptools import find_packages, setup

setup(
    name='google-sheets-wrapper',
    version='2.0.0',
    description=(
        'A library wrapping Google Sheets API to make some operations easier'
    ),
    long_description=(Path(__file__).parent / 'README.md').read_text(),
    url='https://github.com/jakubvalenta/google-sheets-wrapper',
    author='Jakub Valenta',
    author_email='jakub@jakubvalenta.cz',
    license='Apache Software License',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Topic :: Software Development',
        'License :: OSI Approved :: Apache Software License',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
    ],
    packages=find_packages(),
    install_requires=[
        'google-api-python-client',
        'google-auth-httplib2',
        'google-auth-oauthlib',
    ],
)
