@echo off
cd\
cd %1
%3
bin\sc.exe -cpp c:\windows\temp\cpp\%2.cpp c:\windows\temp\cpp\test.exe>c:\windows\temp\%2.txt
