# QAttend Structured Pages

- `login.html` - dedicated login page, Forgot Password, Contact Admin and Request New Section.
- `admin.html` - admin portal.
- `cr.html` - CR portal.
- `style.css` - single shared stylesheet.
- `script.js` - shared Firebase/auth/attendance/request/admin logic.
- `logo.png` - QAttend logo.
- `firebase.json` - Firebase Hosting configuration.
- `firestore.rules` - Firestore security rules.
- `index.html` - redirects to `login.html`.

The existing UI/features are retained. The only structural change is page separation and role-based routing after login.

## Local test
Use a local static server (for example VS Code Live Server) rather than opening HTML with `file://`.
Open `login.html`.

## Firebase Hosting
Deploy the folder as the Hosting public directory. The Firebase project configuration remains the same.
