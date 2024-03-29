@REM @echo off

@REM REM Creating a temporary conda environment for installation...
@REM conda create -n search_replace_tool_temp_environment python=3.10 -y

@REM REM Activating the tempoaray conda environment...
@REM conda activate search_replace_tool_temp_environment

@REM REM Installing pip...
@REM conda install pip -y

@REM REM Installing dependencies from requirements.txt...
@REM pip install -r requirements.txt

@REM REM Creating the dist directory if it doesn't exist...
@REM if not exist dist mkdir dist

@REM REM Building the executable using PyInstaller...
@REM pyinstaller --onefile src/search_replace.py -y

@REM REM Copying the README.md file to the dist directory...
@REM copy README.md dist\

@REM REM Copying the contents of the data folder to the dist directory...
@REM xcopy data dist\ /s /e /y

@REM REM Deactivate the tempoaray conda environment...
@REM conda deactivate

@REM REM Removing the tempoaray conda environment...
@REM conda remove -n search_replace_tool_temp_environment --all -y

rmdir /s dist -y & C:/Users/NButterly/AppData/Local/miniconda3/Scripts/activate & conda create -n search_replace_tool_temp_environment python=3.10 -y & conda activate search_replace_tool_temp_environment & conda install pip -y & pip install -r requirements.txt & if not exist dist mkdir dist & pyinstaller --onefile src/search_replace.py -y & copy README.md dist\ & xcopy data dist\ /s /e /y & conda deactivate & conda remove -n search_replace_tool_temp_environment --all -y & rmdir /s build -y & del search_replace.spec -y