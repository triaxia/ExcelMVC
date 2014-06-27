pushd "%~dp0"

set addin="Forbes.Application.DNA.xll"

if exist "C:\Program Files (x86)\." (

if exist "C:\Program Files\Microsoft Office\Office15\Excel.Exe" (
  set addin="Forbes.Application.DNA (x64).xll"

))

START EXCEL /x %addin% "Forbes2000.xlsx"
popd
