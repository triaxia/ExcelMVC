REM This command is called by the project's post build event. It simply renames dll to
REM xll and copies the sample dlls to where the xll is so that they can be loaded by
REM the Addin automatically.

REM %1 "$(ProjectDir)"
REM %2 "$(TargetDir)"

pushd .

copy "%~1..\Forbes.Models\Forbes.csv" "%~2\*.*"
copy "%~1..\Forbes.Views\Forbes2000.xlsx" "%~2\*.*"

popd