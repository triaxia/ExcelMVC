REM This command launches Excel with the Forbes2000.xlsx and ExcelMvc.Addin.xll

pushd "%~dp0"
START Excel "..\bin\Debug\Mvc\ExcelMvc.Addin.xll" "..\bin\Debug\Mvc\Forbes2000.xlsx""

popd
