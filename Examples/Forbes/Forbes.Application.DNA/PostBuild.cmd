REM This command is called by the project's post build event. It simply copies the sample dlls
REM to where the xll is and then packs the lot into a DNA xll.

REM %1 "$(TargetDir)"
REM %2 "$(TargetName)"
REM %3 "$(OutDir)"

pushd "%~dp0"

set out=%~1Dna\
if exist "%out%\." (
  rmdir /S /Q "%out%"
)
mkdir "%out%"

copy "Run.cmd" "%out%"
copy "Sample.Application.dna" "%out%"

copy "..\..\..\..\ExcelMvc\ExcelMvc\%~3ExcelMvc.dll" "%out%"
copy "..\..\..\..\ExcelMvc\ExcelMvc.AddinDna\%~3ExcelMvc.AddinDna.dll" "%out%"
copy "..\..\..\..\ExcelMvc\ExcelMvc.AddinDna\%~3ExcelDna.Integration.dll" "%out%"

copy "..\..\Sample.Models\%~3Sample.Models.dll" "%out%"
copy "..\..\Sample.Views\%~3Sample.Views.dll" "%out%"
copy "..\..\Sample.Application\%~3Sample.Application.dll" "%out%"
copy "..\..\Sample.Application.dna" "%out%"

copy "..\..\Sample.Models\Forbes.csv" "%out%"
copy "..\..\Sample.Views\Forbes2000.xlsx" "%out%"

copy "..\..\Sample.Application\%~3Sample.Application.dll.config" "%out%Sample.Application.xll.config"

copy "..\..\packages\Excel-DNA.0.32.0\tools\ExcelDnaPack.exe" "%out%"
copy "..\..\packages\Excel-DNA.0.32.0\tools\ExcelDna.xll" "%out%Sample.Application.xll"
copy "..\..\packages\Excel-DNA.0.32.0\tools\ExcelDna64.xll" "%out%Sample.Application (x64).xll"

cd "%out%"

rem x86
ExcelDnaPack.exe "Sample.Application.dna" /Y
del "Sample.Application.xll"
rename "Sample.Application-packed.xll" "Sample.Application.xll"

rem x64
rename "Sample.Application.dna" "Sample.Application (x64).dna"
ExcelDnaPack.exe "Sample.Application (x64).dna" /Y
del "Sample.Application (x64).xll"
rename "Sample.Application (x64)-packed.xll" "Sample.Application (x64).xll"
copy "Sample.Application.xll.config" "Sample.Application (x64).xll.config

del "*.dll"
del "*.exe"
del "*.dna"

popd