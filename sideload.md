# Sideloading Office.js Add-ins on Excel for Mac

## Why the "Add" Button is Greyed Out
- Personal Microsoft accounts often block manifest uploads.
- Excel for Mac uses a file-based sideload method instead of the in-app "Add" button.
- Admin rights on macOS do not override Office account restrictions.

## Steps to Sideload a Manifest
1. **Check Account Type**
   - Sign in with a Microsoft 365 business or education account.
   - Personal accounts may not allow manifest uploads.

2. **Clear Add-in Cache**
   - Quit Excel and other Office apps.
   - Navigate to:
     ```
     ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
     ```
   - Delete cached add-in data.

3. **Place Manifest in Sideload Folder**
   - Copy your manifest XML file into:
     ```
     /Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef
     ```
   - Replace `<username>` with your macOS account name.

4. **Restart Excel**
   - Open Excel again.
   - Your add-in should appear under *Developer Add-ins*.

## For Production Deployment
- Use **Centralized Deployment** in Microsoft 365 admin center.
- Ensures add-in availability across devices and avoids sideload quirks.