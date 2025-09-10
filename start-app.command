#!/bin/bash
# Change directory to the script's location
cd "$(dirname "$0")"

echo "Checking and installing dependencies (this may take a moment)..."
npm install

echo "Starting the application server..."
echo "Your browser should open to http://localhost:3000 shortly."
open http://localhost:3000
npm start
