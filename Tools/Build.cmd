rem ------------------------------------------------------------------------------------
rem builds ExcelMVC solutions

pushd .

cd "%~dp0"

call "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\Tools\vsvars32.bat"
msbuild "..\Examples\Sample\Sample.sln" /t:Rebuild /p:Configuration="Release Package" /p:Platform="Any CPU" /clp:ErrorsOnly
msbuild "..\Examples\SpotTrading\SpotTrading.sln" /t:Rebuild /p:Configuration="Release Package" /p:Platform="Any CPU" /clp:ErrorsOnly
msbuild "..\ExcelMvc\ExcelMvc.sln" /t:Rebuild /p:Configuration="Release Package" /p:Platform="Any CPU" /clp:ErrorsOnly

popd

pause

goto :eof
