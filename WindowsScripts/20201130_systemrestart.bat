@echo off
echo Windows will restart after a few seconds.
echo !! If u wan't restart now, pls close this window!!

rem wating...
@ping 127.0.0.1 -n 30 > nul
echo 30
@ping 127.0.0.1 -n 20 > nul
echo 20
@ping 127.0.0.1 -n 10 > nul
echo 10
@ping 127.0.0.1 -n 5 > nul
echo 5
@ping 127.0.0.1 -n 2 > nul
echo .

echo System will restart after 30s. 
shutdown -r
