QAttend Admin Corrected Localhost Build

Core files:
- index.html
- style.css
- script.js
- firestore.rules
- firebase.json
- logo.png
- Section_Template.xlsx
- seed.html / seed.js / seed-backup.js

This build restores the Admin layout and wires the Admin dashboard, Manage Sections,
Manage Users, Reports, Requests, course management, and the existing CR/attendance UI.
The existing visual theme is retained. The Excel template is included because the
section request form references it.

Run locally:
  python3 -m http.server 5500
Then open:
  http://localhost:5500

Firebase:
  Publish firestore.rules before testing Firestore operations.
  Create Firebase Authentication users separately. The Admin user profile must exist
  in users/{admin-email} with role: "admin".
