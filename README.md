# QAttend Portal v2

QAttend is a Firebase-based university attendance portal with separate Admin and Class Representative (CR) roles.

## New role workflow

### CR
- Login -> directly opens the assigned section attendance page.
- Assigned section is automatic and cannot be changed.
- Mark Present/Absent, Mark All Present/Absent, Save, Download and Share.
- Requests menu lets the CR send student/subject change requests to Admin.
- CR never receives Admin access.

### Admin
- Login -> Admin Dashboard.
- Dashboard shows total students, subjects, active CRs and pending requests.
- Manage Students, Subjects and CR assignments.
- Review CR requests and Approve / Reject them.
- Approving a request applies the requested student/subject change.
- Admin can also mark attendance for any section.

## Firebase data layout

- `users/{email}` -> `{ role: "admin" | "cr", section: "SECTION-1" }`
- `sections/{section}/students/{studentId}` -> `{ qid, name }`
- `sections/{section}/subjects/{subjectId}` -> `{ name }`
- `sections/{section}/attendance/{recordId}` -> saved attendance
- `requests/{requestId}` -> CR change request + approval status

## Files

- `index.html`
- `style.css`
- `script.js`
- `firestore.rules`
- `firebase.json`
- `logo.png`
- `seed-backup.js`

## Deploy

Publish the security rules first, then deploy Hosting:

```bash
firebase deploy
```

CR/Admin role separation is enforced in Firestore Rules. The client-side UI is only one layer of the protection.
