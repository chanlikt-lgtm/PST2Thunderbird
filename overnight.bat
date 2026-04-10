@echo off
echo ============================================
echo  PST Overnight Job — %date% %time%
echo ============================================
echo.

echo [1/4] Waiting for current reconvert (backup+11 PSTs) to finish...
:WAIT
tasklist /FI "PID eq 20784" 2>NUL | find "20784" >NUL
if %ERRORLEVEL%==0 (
    timeout /t 30 /nobreak >NUL
    goto WAIT
)
echo Done waiting.
echo.

echo [2/4] Running sort_folders.py (add YYYY-MM prefixes)...
py -3.11 "E:\claude\Pst2Thunder\sort_folders.py"
echo.

echo [3/4] Running reconvert --all (fix Date headers on all 120 PSTs)...
py -3.11 "E:\claude\Pst2Thunder\reconvert.py" --all
echo.

echo [4/4] Running sort_folders.py again (prefix any new folders)...
py -3.11 "E:\claude\Pst2Thunder\sort_folders.py"
echo.

echo ============================================
echo  ALL DONE — %date% %time%
echo  Restart Thunderbird to see all changes.
echo ============================================
