import os
import subprocess
import shutil
from distutils.dir_util import copy_tree
from setuptools import setup
from setuptools.command.install import install


class InstallCommand(install):
    """Custom install command to build the executable using PyInstaller"""

    def run(self):
        # Run the default install command
        install.run(self)

        # Install dependencies from requirements.txt
        subprocess.run(["pip", "install", "-r", "requirements.txt"])

        from tqdm import tqdm

        tqdm.write("Installing...")

        # Build the executable using PyInstaller
        subprocess.run(["pyinstaller", "--onefile", "src/search_replace.py"])

        # Copy the contents of the data folder to the dist directory
        dist_dir = os.path.join(self.install_lib, "search_replace_tool")
        data_dir = os.path.join(os.getcwd(), "data")
        copy_tree(data_dir, os.path.join(dist_dir, "data"))

        # Copy the README.md file to the dist directory
        shutil.copy2("README.md", dist_dir)
        tqdm.write("Installation complete")


with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="search_replace_tool",
    version="0.1.0",
    author="Nicholas Butterly",
    author_email="nicholas.butterly@aemo.com.au",
    long_description=long_description,
    long_description_content_type="text/markdown",
    py_modules=["search_replace"],
    cmdclass={
        "install": InstallCommand,
    },
)
