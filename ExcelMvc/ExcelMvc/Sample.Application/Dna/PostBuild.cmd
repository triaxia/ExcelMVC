REM This command is called by the project's post build event. It simply copies the sample dlls
REM to where the xll is and then packs the lot into a DNA xll.

REM %1 "$(TargetDir)"
REM %2 "$(TargetName)"
REM %3 $(ConfigurationName)

pushd "%~dp0"

set out=%~1Dna\
if exist "%out%\." (
  rmdir /S /Q "%out%"
)
mkdir "%out%"

copy "..\..\..\ExcelDna\ExcelDnaPack.exe" "%out%"
copy "..\..\..\ExcelDna\ExcelDna.xll" "%out%ExcelMvc.AddinDna.xll"
rem copy "..\..\..\ExcelDna\ExcelDna64.xll" "%out%ExcelMvc.AddinDna.xll"

copy "Start.cmd" "%out%"
copy "ExcelMvc.AddinDna.dna" "%out%"

copy "..\..\ExcelMvc\bin\%~3\ExcelMvc.dll" "%out%"
copy "..\..\ExcelMvc.AddinDna\bin\%~3\ExcelMvc.AddinDna.dll" "%out%"
copy "..\..\ExcelMvc.AddinDna\bin\%~3\ExcelDna.Integration.dll" "%out%"

copy "..\..\Sample.Models\bin\%~3\Sample.Models.dll" "%out%"
copy "..\..\Sample.Views\bin\%~3\Sample.Views.dll" "%out%"
copy "..\..\Sample.Application\bin\%~3\Sample.Application.dll" "%out%"

copy "..\..\Sample.Models\Forbes.csv" "%out%"
copy "..\..\Sample.Views\Forbes2000.xlsx" "%out%"

if exist "%~1%~2.dll.config" (
	copy "%~1%~2.dll.config" "%~1%~2.xll.config"
)

cd "%out%"
ExcelDnaPack.exe "ExcelMvc.AddinDna.dna" /Y

del "*.dll"
del "*.exe"
del "*.dna"

popd