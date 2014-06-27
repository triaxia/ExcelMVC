rem ------------------------------------------------------------------------------------
rem Packages ExcelMVC

pushd .

cd "%~dp0"

set out=release
if exist "%out%\." (
  rmdir /S /Q "%out"
)

mkdir %out%

xcopy /Y /S /R "..\ExcelMvc\ExcelMvc\bin\Release\*.*" "%out%\Bin\"

xcopy /Y /S /R "..\Examples\SpotTrading\SpotTrading\bin\Release\net35\*.*" "%out%\Samples\Trading\"
xcopy /Y /S /R "..\Examples\Forbes\Forbes.Application.MVC\bin\Release\*.*"  "%out%\samples\Forbes.MVC\"
xcopy /Y /S /R "..\Examples\Forbes\Forbes.Application.DNA\bin\Release\*.*"  "%out%\samples\Forbes.DNA\"

xcopy /Y /S /R "..\ExcelMvc\*.*" "%out%\source\ExcelMvc\"
xcopy /Y /S /R "..\Examples\*.*" "%out%\source\Examples\"

call clean.cmd "%cd%\ExcelMvc"
call clean.cmd "%cd%\Examples"

popd

pause

goto :eof
