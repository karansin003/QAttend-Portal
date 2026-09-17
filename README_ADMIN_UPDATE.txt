QAttend Admin Panel Update

Files:
- index.html
- style.css
- script.js
- firestore.rules

What was added (Admin only):
1. Admin Dashboard stats and request center remain available.
2. Requests table with All / Pending / Students / Subjects / New Sections / Contact Admin filters.
3. Individual View / Approve / Reject actions.
4. Checkbox selection with Accept Selected / Reject Selected.
5. Request details modal.
6. Approving Add/Delete Student or Subject applies the requested Firestore change.
7. Approving New Section creates section metadata, imports the students/subjects stored in the request, and assigns the CR profile.
8. Rejecting a request only changes request status and does not modify section data.
9. Manage Sections: add/edit/delete dynamic section metadata. Existing fixed system sections are protected from deletion here.
10. Manage Users: view, add/update profile, activate/deactivate users. Firebase Auth account creation is intentionally not done from client-side admin UI.
11. Reports: CSV export for attendance, requests, students and subjects. Reports load only when clicked.
12. Contact Admin login-page submissions are also saved as requests, while preserving the existing mailto behavior.

Important:
- Existing login/CR/attendance styling is not intentionally changed by this update.
- Publish firestore.rules before testing public Contact Admin/New Section requests.
- Firebase Authentication passwords cannot safely be created from browser-only client code. Create Auth accounts through Firebase Console or a trusted backend/admin workflow.
