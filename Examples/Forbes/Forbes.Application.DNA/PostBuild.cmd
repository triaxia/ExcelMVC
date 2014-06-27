REM %1 "$(ProjectDir)"
REM %2 "$(TargetDir)"

pushd .

-----------------------------------------------------------------------------------------------
REM copy the workbook and data file the target directory

copy "%~1..\Forbes.Models\Forbes.csv" "%~2*.*"
copy "%~1..\Forbes.Views\Forbes2000.xlsx" "%~2*.*"

-----------------------------------------------------------------------------------------------
REM copy DNA pack files to the target directory

copy "%~1..\packages\Excel-DNA.0.32.0\tools\ExcelDnaPack.exe" "%~2"
copy "%~1..\packages\Excel-DNA.0.32.0\tools\ExcelDna.Integration.dll" "%~2"
copy "%~1..\packages\Excel-DNA.0.32.0\tools\ExcelDna.xll" "%~2Forbes.Application.DNA.xll"
copy "%~1..\packages\Excel-DNA.0.32.0\tools\ExcelDna64.xll" "%~2Forbes.Application.DNA (x64).xll"

-----------------------------------------------------------------------------------------------
REM pack x86 addin

ExcelDnaPack.exe "Forbes.Application.DNA.dna" /Y
del "Forbes.Application.DNA.xll"
rename "Forbes.Application.DNA-packed.xll" "Forbes.Application.DNA.xll"
copy "Forbes.Application.DNA.dll.config" "Forbes.Application.DNA.xll.config"

-----------------------------------------------------------------------------------------------
REM pack x64 addin

rename "Forbes.Application.DNA.dna" "Forbes.Application.DNA (x64).dna"
ExcelDnaPack.exe "Forbes.Application.DNA (x64).dna" /Y
del "Forbes.Application.DNA (x64).xll"
rename "Forbes.Application.DNA (x64)-packed.xll" "Forbes.Application.DNA (x64).xll"
copy "Forbes.Application.DNA.dll.config" "Forbes.Application.DNA (x64).xll.config"

-----------------------------------------------------------------------------------------------
REM clean up unwanted files
del "*.dna"
del "*.exe"
del "*.pdb"
del "*.dll"
del "*.xml"
del "*.dll.config"

popd