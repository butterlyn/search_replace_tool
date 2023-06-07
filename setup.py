from setuptools import setup, find_packages

with open('requirements.txt') as f:
    requirements = f.read().splitlines()

setup(
    name='glossary_formatter',
    version='0.0.1',
    packages=find_packages(),
    install_requires=requirements,
    entry_points={
        'console_scripts': [
            'glossary_formatter=glossary_formatter:main',
        ],
    },
    # Metadata
    author='Nicholas Butterly',
    author_email='nicholas.butterly@aemo.com.au',
)
