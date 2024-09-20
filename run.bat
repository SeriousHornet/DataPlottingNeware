@echo off
cd /d %~dp0
set "folder_path=%~dp0"

:: Loop through each Excel file in the folder
for %%f in ("%folder_path%\*.xlsx") do (
	echo Processing %%~nxf
:: Call the Python script and pass the Excel file as an argument
	python MasterPlotter.py "%%f"
	)

pause
