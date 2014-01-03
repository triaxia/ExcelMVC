REM This command launches Excel with the Forbes2000.xlsx and ExcelMvc.AddinDna.xll

pushd "%~dp0"
START Excel "..\bin\Debug\Dna\ExcelMvc.AddinDna-packed.xll" "..\bin\Debug\Dna\Forbes2000.xlsx""

popd
