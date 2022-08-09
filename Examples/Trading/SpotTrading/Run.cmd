pushd .

cd "%~dp0"

set addin=ExcelMvc.Addin.x86.xll
REM set addin=ExcelMvc.Addin.x64.xll

REM workbooks passed to Excel require FULL path
START Excel /X "%cd%\%addin%" "%cd%\Views\SpotTrading.xlsx"

popd
