@echo off

call clean

echo eVB crashes when the app is compiled with the /make argument, please compile the app manually.
echo Place the application in the project root directory or Setup\App.
"C:\Program Files (x86)\Microsoft eMbedded Tools\EVB\EVB3.EXE" "Telephone Alphabet.ebp"

move *.vb Setup\App
copy Assets\App\*.* Setup\App
copy Assets\Windows\*.* Setup\Windows

"C:\Program Files (x86)\Microsoft eMbedded Tools\EVB\cabwiz.exe" "%~dp0Setup\Phonetic.INF" /cpu "CEF" "Mips 4000 (1K) v2.10" "SH 4 (4K) v2.10" "Arm 1100 (4K) v2.10" "SH 3 (1K) v2.10"

del Setup\*.DAT
move Setup\*.cab Setup\CD1