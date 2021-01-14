@echo off
TITLE Chrome Debug mode
echo -----------------------------------------------------------------------------------------------------------------------
echo 							INSTRUCTIONS
echo -----------------------------------------------------------------------------------------------------------------------
echo 1. Make sure chrome browser and chrome driver are installed in your system. 
echo --Link for chrome driver: https://chromedriver.chromium.org/
echo --(Note: Version of chrome browser and chrome driver has to match, if not python script will crash.)
echo 2. Add chrome browser's location to PATH.
echo 3. Sign in to google (only for the first time).
echo 4. Open meet and enter the meeting id.
echo 5. After all the participants have entered the secret key, run the python script to make note of all the participants.
echo --(Make sure python 3.5+ is installed in your pc).
echo -----------------------------------------------------------------------------------------------------------------------
echo This window will close as soon as you exit from chrome.
echo Starting chrome...

chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenum\ChromeProfile"

pause
