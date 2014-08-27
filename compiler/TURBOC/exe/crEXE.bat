@echo off
cd\
cd %1
cls
%3
cd bin
..\bin\tcc.exe -v -y -M  -I..\include -L..\lib -o%2.obj -ec:\windows\temp\cpp\%2.exe c:\windows\temp\cpp\%2.cpp>c:\windows\temp\msg.txt
copy c:\windows\temp\msg.txt c:\windows\temp\%2.txt
pause
