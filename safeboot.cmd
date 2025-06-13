@echo off
echo ................................................
echo Press 1, 2, 3 to select your task, or 4 to EXIT.
echo Run this .bat with admin privileges            .
echo ................................................
echo.
echo 1 - Safe Boot Minimal
echo 2 - Safe Boot with Network
echo 3 - Normal Boot
echo 4 - Exit
echo.
SET /P M=Type 1, 2, 3, 4 then press ENTER:
IF %M%==1 GOTO safe
IF %M%==2 GOTO safenet
IF %M%==3 GOTO normal
IF %M%==4 GOTO exit

:safe
bcdedit /set {default} safeboot minimal
goto reboot

:safenet
bcdedit /set {default} safeboot network
goto reboot

:normal
bcdedit /deletevalue {default} safeboot
goto reboot

:reboot
shutdown -r -f -t 4
exit

:exit
exit

source: https://pastebin.com/DV