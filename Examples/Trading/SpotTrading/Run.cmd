pushd .

cd "%~dp0"

set addin=ExcelMvc.Addin.xll
REM set addin=ExcelMvc.Addin (x64).xll

START Excel /X "%cd%\%addin%" "%cd%\Views\SpotTrading.xlsx"

popd
