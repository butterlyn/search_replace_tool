from setuptools import setup, find_packages

with open("README.md", "r") as fh:
    long_description = fh.read()

with open("requirements.txt", "r") as f:
    requirements = f.read().splitlines()

setup(
    name="search_replace_tool",
    version="0.0.2",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=requirements,
    author="Nicholas Butterly",
    author_email="nicholas.butterly@aemo.com.au",
    description="A tool for searching and replacing text in .docx files",
    long_description=long_description,
    long_description_content_type="text/markdown",
    license="MIT",
    keywords="search replace tool",
    url="https://github.com/yourusername/search_replace_tool",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
