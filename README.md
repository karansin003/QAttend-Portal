# QAttend Attendance Management System

This package contains the separated source files from the current QAttend project.

## Files
- `index.html` - UI/HTML
- `style.css` - styling/responsive layout
- `script.js` - Firebase Auth/Firestore logic and request flows
- `firestore.rules` - Firestore security rules

## Run locally
Open the folder with VS Code and use Live Server (recommended), then open the local URL.
Do not open `index.html` directly with `file://` because Firebase modules and browser security can cause issues.

## Firebase
The Firebase project configuration already present in `script.js` is retained. Publish `firestore.rules` to the same Firebase project before testing request submission.

## Important
`index.html` references `logo.png`. Keep your existing QAttend `logo.png` in this same folder.
