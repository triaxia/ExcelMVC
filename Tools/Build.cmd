@echo off
REM ----------------------------------------------------------------------------------------------------------------------------------
REM Builds ExcelMVC solutions
REM set /p version=Enter the build version (i.j.k):

pushd "%~dp0"

call "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\Tools\vsvars32.bat"
msbuild "..\Examples\Forebes\Forbes.sln" /t:Rebuild /p:Configuration="Release" /p:Platform="Any CPU" /clp:ErrorsOnly
msbuild "..\Examples\SpotTrading\SpotTrading.sln" /t:Rebuild /p:Configuration="Release Package" /p:Platform="Any CPU" /clp:ErrorsOnly
msbuild "..\ExcelMvc\ExcelMvc.sln" /t:Rebuild /p:Configuration="Release Package" /p:Platform="Any CPU" /clp:ErrorsOnly

call ::package_release "%cd%" "release"

popd

pause

goto :eof

REM ----------------------------------------------------------------------------------------------------------------------------------
REM Packages ExcelMVC release
:package_release

pushd "%~1"

set out=%~2
if exist "%out%\." (
  rmdir /S /Q "%out"
)

mkdir %out%

xcopy /Y /S /R "..\ExcelMvc\ExcelMvc\bin\Release\*.*" "%out%\bin\"

xcopy /Y /S /R "..\Examples\Trading\SpotTrading\bin\Release\*.*" "%out%\Samples\Trading\"
xcopy /Y /S /R "..\Examples\Forbes\Forbes.Application.MVC\bin\Release\*.*"  "%out%\Samples\Forbes.MVC\"
xcopy /Y /S /R "..\Examples\Forbes\Forbes.Application.DNA\bin\Release\*.*"  "%out%\Samples\Forbes.DNA\"

xcopy /Y /S /R "..\ExcelMvc\*.*" "%out%\Source\ExcelMvc\"
xcopy /Y /S /R "..\Examples\*.*" "%out%\Source\Examples\"

@echo clean up build files
call :clean_dir "%out%\Source\ExcelMvc"
call :clean_dir "%out%\Source\Examples"

popd

goto :eof


REM ----------------------------------------------------------------------------------------------------------------------------------
REM Cleans up build files
:clean_dir

pushd "%~1"

call :clean_loop "%~1"

popd
goto :eof

:clean_loop

pushd "%~1"

if exist "*.sdf" del /Q *.sdf
if exist "*.vsmdi" del /Q *.vsmdi 
if exist "*.vspscc" del /Q *.vspscc 
if exist "*.vssscc" del /Q *.vssscc 
if exist "*.obj" del /Q *.obj 
if exist "*.cache" del /Q *.cache 
if exist "bin\." rmdir /S /Q bin 
if exist "obj\." rmdir /S /Q obj 
if exist "TestResults\." rmdir /S /Q TestResults 
if exist "Debug\." rmdir /S /Q Debug  
if exist "Release\." rmdir /S /Q Release 
if exist "packages\." rmdir /S /Q packages 
if exist "ipch\." rmdir /S /Q ipch 

for /D %%x in (*) do call :clean_loop "%%x"

popd

goto :eof