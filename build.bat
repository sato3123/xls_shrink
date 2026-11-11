@echo off
cd /d "%~dp0"
if exist build_done.flag del build_done.flag
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist bin mkdir bin
"%CSC%" /nologo /utf8output /codepage:65001 /target:winexe /platform:anycpu /optimize+ /out:bin\XlsImageShrink.exe /r:System.dll /r:System.Core.dll /r:Microsoft.CSharp.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:System.Web.Extensions.dll /r:Microsoft.VisualBasic.dll src\*.cs > build.log 2>&1
set RC=%ERRORLEVEL%
echo EXITCODE=%RC% >> build.log
ver >> build.log
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >> build.log 2>&1
if not "%RC%"=="0" goto done
if "%1"=="notest" goto done
start "" /wait bin\XlsImageShrink.exe --selftest "%~dp0selftest"
echo SELFTEST_EXIT=%ERRORLEVEL% >> build.log
:done
echo done> build_done.flag
