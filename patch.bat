@echo off
@cls
cd %1
%2
patcher.exe x -y %3 %4>c:\windows\temp\data.txt
pause