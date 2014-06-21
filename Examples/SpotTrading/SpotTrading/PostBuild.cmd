REM This command is called by the project's post build event. It simply renames dll to
REM xll and copies the sample dlls to where the xll is so that they can be loaded by
REM the Addin automatically.

REM %1 "$(TargetDir)"
REM %2 "$(ConfigurationName)"

pushd "%~dp0"

copy "..\..\..\ExcelMvc\ExcelMvc.Addin\bin\%~2\ExcelMvc.Addin.xll" "%~1ExcelMvc.Addin.xll"
copy "..\..\..\ExcelMvc\ExcelMvc.Addin\bin\%~2 (x64)\ExcelMvc.Addin (x64).xll" "%~1ExcelMvc.Addin (x64).xll"

popd