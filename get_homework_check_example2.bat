echo off & color 0A

set DIR="%cd%"
echo DIR=%DIR%
echo %0
echo %~f1

echo start python %~dp0\homework_check2.py %~f1
start python %~dp0\homework_check2.py %~f1
::pause