REM This command launches Excel with the Forbes2000.xlsx and ExcelMvc.AddinDna.xll

pushd "%~dp0"

if exist ".\PostBuild.cmd" (
START Excel "..\bin\Debug\Dna\Sample.Application.xll" "..\bin\Debug\Dna\Forbes2000.xlsx"
) else (
START Excel "Sample.Application.xll" "Forbes2000.xlsx"
)

popd
