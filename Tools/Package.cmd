rem ------------------------------------------------------------------------------------
rem Packages ExcelMVC

pushd .

cd "%~dp0"

set out=release
if exist "%out%\." (
  rmdir /S /Q "%out"
)

mkdir %out%

xcopy /Y /S /R "..\ExcelMvc\ExcelMvc\bin\Release\*.*" "%out%\bin\"
xcopy /Y /S /R "..\Examples\SpotTrading\SpotTrading\bin\Release\net35\*.*" "%out%\samples\trading\"
xcopy /Y /S /R "..\Examples\Forebes\Forebes.Application\bin\Release\net35\Mvc\*.*"  "%out%\samples\forbes\"

xcopy /Y /S /R "..\ExcelMvc\*.*" "%out%\source\ExcelMvc\"
xcopy /Y /S /R "..\Examples\*.*" "%out%\source\Examples\"

call clean.cmd "%cd%\ExcelMvc"
call clean.cmd "%cd%\Examples"

popd

pause

goto :eof
