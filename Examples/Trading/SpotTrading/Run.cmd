pushd .

cd "%~dp0"

REM set addin=SpotTrading.xll
set addin=SpotTrading64.xll

REM Use full path for Workbook argments. ExcelMvc gets upset with relative Workbook
REM paths... 
START Excel /X "%cd%\%addin%" "%cd%\Views\SpotTrading.xlsx"

popd
