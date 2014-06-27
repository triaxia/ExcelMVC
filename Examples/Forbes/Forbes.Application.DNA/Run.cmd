REM This command launches Excel with the Forbes2000.xlsx and ExcelMvc.AddinDna.xll

pushd "%~dp0"

set addin="Forbes.Application.xll"
if exist "C:\Program Files (x86)\." (
if exist "C:\Program Files\Microsoft Office\Office15\." (
set addin="Forbes.Application (x64).xll"
))

START EXCEL /x %addin% "Forbes2000.xlsx"
popd
