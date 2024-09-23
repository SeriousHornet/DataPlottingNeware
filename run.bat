@echo off
cd /d %~dp0
set "folder_path=%~dp0"

:: Create "Plot Data" folder if it doesn't exist
if not exist "%folder_path%\Plot Data" (
    mkdir "%folder_path%\Plot Data"
)

:: Loop through each Excel file in the folder
for %%f in ("%folder_path%\*.xlsx") do (
	echo Processing %%~nxf
	:: Call the Python script and pass the Excel file and the "Plot Data" folder as arguments
	python MasterPlotter.py "%%f" "%folder_path%\Plot Data"
)

pause