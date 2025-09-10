@echo off
echo Changing directory to the application folder...
cd /d "%~dp0"

echo Checking and installing dependencies (this may take a moment)...
npm install

echo Starting the application server...
echo Your browser should open to http://localhost:3000 shortly.
start http://localhost:3000
npm start
