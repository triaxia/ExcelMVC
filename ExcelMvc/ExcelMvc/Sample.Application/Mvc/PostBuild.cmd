REM This command is called by the project's post build event. It simply renames dll to
REM xll and copies the sample dlls to where the xll is so that they can be loaded by
REM the Addin automatically.

REM %1 "$(TargetDir)"
REM %2 "$(TargetName)"
REM %3 $(ConfigurationName)

pushd "%~dp0"

set out=%~1Mvc\
if exist "%out%\." (
  rmdir /S /Q "%out%"
)
mkdir "%out%"

copy "Start.cmd" "%out%"

copy "..\..\ExcelMvc\bin\%~3\ExcelMvc.dll" "%out%"
copy "..\..\ExcelMvc.Addin\bin\%~3\ExcelMvc.Addin.dll" "%out%ExcelMvc.Addin.xll"

copy "..\..\Sample.Models\bin\%~3\Sample.Models.dll" "%out%"
copy "..\..\Sample.Views\bin\%~3\Sample.Views.dll" "%out%"
copy "..\..\Sample.Application\bin\%~3\Sample.Application.dll" "%out%"

copy "..\..\Sample.Models\Forbes.csv" "%out%"
copy "..\..\Sample.Views\Forbes2000.xlsx" "%out%"

popd