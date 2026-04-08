# Nubra Insti Excel Plugin

This Excel plugin is built for Nubra's institutional clients to log in, add instruments, and build an Open Interest dashboard inside Excel.

## What This Does

- Launches the Nubra Excel plugin using `NubraInstiExcelLauncher.exe`
- Lets Insti clients log in with their Nubra credentials
- Verifies access using MPIN
- Loads the F&O stocks cache
- Adds instruments to the live tracker
- Helps build and refresh the OI dashboard in Excel

## Download And Start

1. Download this project folder to your Windows machine.
2. Keep all files together in the same folder. Do not move the `.exe` file away from the other project files.
3. Double-click `NubraInstiExcelLauncher.exe`.
4. Wait for the launcher to start the local Nubra Excel plugin services.
5. Open Microsoft Excel.
6. Open the Nubra Excel add-in/task pane when prompted by the launcher setup.

## Login Flow For Nubra Insti Clients

1. In the Excel add-in, enter your institutional credentials:
   - Exchange Client Code
   - Client Code
   - Username
   - Password
2. Click `Login Insti`.
3. Enter your 4-digit MPIN.
4. Click `Verify MPIN`.
5. Once logged in, the add-in will show that your session is active.

## Build Your OI Dashboard

1. Click `Load F&O Stocks Cache` to load available stocks with futures data.
2. Search for a stock in `Track Stock`.
3. Click `Add To Live Tracker` to add the instrument to the live tracker sheet.
4. Use the interval tracker controls to define market time slots.
5. Click `Create Interval Sheet` to prepare interval-based tracking.
6. Click `Refresh Latest Interval` to update the latest live interval data.

## Important Notes

- This plugin is intended only for Nubra institutional clients.
- You must have valid Nubra Insti login credentials and MPIN access.
- The launcher and supporting files should remain in the same folder for the setup to work correctly.
- Excel desktop on Windows is recommended.
