# QAttend v2 Setup

1. Keep the Firebase project configuration in `script.js` aligned with your Firebase project.
2. Publish `firestore.rules` in Firebase Console -> Firestore Database -> Rules.
3. Create the Admin Firebase Auth user.
4. Create `users/{admin-email}` with `role: "admin"`.
5. Log in as Admin.
6. Create CR Firebase Auth users using lowercase emails.
7. From Admin -> Manage System -> Assign CR, assign each CR to a section.
8. CR login now opens its assigned attendance section automatically.
9. CR requests for student/subject changes go to Admin under Requests.
10. Admin Approve applies the requested change; Reject leaves the data unchanged.

Important:
- Do not give CR users Admin role.
- Keep CR emails lowercase to match the user document IDs.
- Delete/replace old Firestore rules before testing request creation, because the new `requests` collection needs the new rule block.
