rem ------------------------------------------------------------------------------------
rem cleans up compilation result files

pushd .

cd "%~dp0"

call :loop "%cd%\..\ExcelMvc"
call :loop "%cd%\..\Examples"

popd
goto :eof

:loop

pushd .
cd "%~1"

if exist "*.sdf" del /Q *.vsmdi
if exist "*.vsmdi" del /Q *.vsmdi
if exist "*.vspscc" del /Q *.vspscc
if exist "*.vssscc" del /Q *.vssscc
if exist "*.obj" del /Q *.obj
if exist "*.cache" del /Q *.cache
if exist "bin\." rmdir /S /Q bin
if exist "obj\." rmdir /s /Q obj
if exist "TestResults\." rmdir /S /Q TestResults
if exist "Debug\." rmdir /S /Q Debug
if exist "Release\." rmdir /S /Q Release

for /D %%x in (*) do call :loop "%%x"

popd

goto :eof