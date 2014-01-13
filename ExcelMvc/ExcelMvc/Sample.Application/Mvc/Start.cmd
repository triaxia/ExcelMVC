REM This command launches Excel with the Forbes2000.xlsx and ExcelMvc.Addin.xll

pushd "%~dp0"

if exist ".\PostBuild.cmd" (
START Excel "..\bin\Debug\Mvc\ExcelMvc.Addin.xll" "..\bin\Debug\Mvc\Forbes2000.xlsx""
) else (
START Excel "ExcelMvc.Addin.xll" "Forbes2000.xlsx""
)

popd
