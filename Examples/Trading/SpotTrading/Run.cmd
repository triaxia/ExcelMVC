pushd .

cd "%~dp0"

REM set addin=ExcelMvc.Addin.x86.xll
set addin=ExcelMvc.Addin.x64.xll

REM Use full path for Workbook argments. ExcelMvc gets upset with relative Workbook
REM paths... 
START Excel /X "%cd%\%addin%" "%cd%\Views\SpotTrading.xlsx"

popd
