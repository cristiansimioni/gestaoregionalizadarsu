@echo off
@setlocal enableextensions
@cd /d "%~dp0"

color 1F
set current_dir=%cd%
cd %current_dir%

rem --Install python
echo Instalando Python [AGUARDE]
python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
echo Instalando Python [OK]
echo.


rem --Refresh Environmental Variables
call refresh-env.cmd

echo Checando a versao instalada [AGUARDE]
python --version
echo Checando a versao instalada [OK]
echo.

echo Configurando dependencias [AGUARDE]
cd ../src/combinations/
python -m pip install --upgrade pip
pip install -r requirements.txt
echo Configurando dependencias [OK]

echo.
echo Pronto. A ferramenta Gestao Regionalizada RSU ja pode ser utilizada! :D
echo.

pause