;echo %Windir%\ 
echo off
regsvr32 /u "%Windir%\Downloaded Program Files\TRContactFinder.dll"
del "%Windir%\Downloaded Program Files\TRContactFinder.dll"
