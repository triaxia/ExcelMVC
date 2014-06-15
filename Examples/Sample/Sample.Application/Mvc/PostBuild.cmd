REM This command is called by the project's post build event. It simply renames dll to
REM xll and copies the sample dlls to where the xll is so that they can be loaded by
REM the Addin automatically.

REM %1 "$(TargetDir)"
REM %2 "$(TargetName)"
REM %3 $(OutDir)
REM %4 "$(ConfigurationName)"

pushd "%~dp0"

set out=%~1Mvc\
if exist "%out%\." (
  rmdir /S /Q "%out%"
)
mkdir "%out%"

copy "Run.cmd" "%out%"

copy "..\..\..\..\ExcelMvc\ExcelMvc\%~3ExcelMvc.dll" "%out%"

copy "..\..\..\..\ExcelMvc\ExcelMvc.Addin\bin\%~4\ExcelMvc.Addin.xll" "%out%ExcelMvc.Addin.xll"
copy "..\..\..\..\ExcelMvc\ExcelMvc.Addin\bin\%~4 (x64)\ExcelMvc.Addin (x64).xll" "%out%ExcelMvc.Addin (x64).xll"

copy "..\..\Sample.Models\%~3Sample.Models.dll" "%out%"
copy "..\..\Sample.Views\%~3Sample.Views.dll" "%out%"
copy "..\..\Sample.Application\%~3Sample.Application.dll" "%out%"
copy "..\..\Sample.Application\%~3Sample.Application.dll.config" "%out%"

copy "..\..\Sample.Models\Forbes.csv" "%out%"
copy "..\..\Sample.Views\Forbes2000.xlsx" "%out%"

popd