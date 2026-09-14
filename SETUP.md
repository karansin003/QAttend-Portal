# QAttend — Setup Steps (do these in order)

## Rule for every login: use all-lowercase emails
Firebase does not force lowercase. If casing is inconsistent between the Auth
login and the role assignment, the app will treat the account as unassigned.
Always create accounts and assign roles using lowercase email addresses.

---

## Step 1 — Turn on Firestore (if not already on)
Firebase Console → your project → Build → Firestore Database → Create database.
Choose production mode. Any region.

## Step 2 — Publish the security rules
Firebase Console → Firestore Database → Rules tab → paste the entire contents
of `firestore.rules` → Publish.

Without this step, your database has no real access control — the app will
still work, but the Admin/CR separation is fake.

## Step 3 — Create the first Admin account
This one step has to be done manually, because the app needs an existing
Admin before it can create anything.

1. Firebase Console → Authentication → Users → Add user.
   Enter your email (lowercase) and a password.
2. Firebase Console → Firestore Database → Data → Start collection.
   - Collection ID: `users`
   - Document ID: your email (exact same lowercase email as step 1)
   - Field: `role` (string) = `admin`
   - Save.

## Step 4 — Deploy the site files and log in
Upload `index.html`, `style.css`, `script.js`, `seed.html`, `seed.js`,
`logo.png`, `firebase.json` to Firebase Hosting as usual. Open the site,
log in with the Admin account from Step 3.

## Step 5 — Run the one-time data import
While logged in as Admin, open `yoursite.com/seed.html` and click
"Run Import" once. This loads the full roster (539 students across
Section 1-8, AIML-1, AIML-2) and the 10 subjects into the database.
After it finishes, you can delete `seed.html` and `seed.js` — they're not
needed again and aren't linked from the main app.

There's no roster for a **CSCQ** section yet — that data wasn't in the sheet
you gave me. Add those students manually through the Admin panel once you
have the list, the same way you'd add any student.

9 rows in the sheet were exact duplicate Q.IDs — those were skipped
automatically so the same student doesn't get added twice.

## Step 6 — Onboard each CR
1. Firebase Console → Authentication → Add user → create their login
   (lowercase email + password you give them).
2. In the app, log in as Admin → "Assign CR" panel → enter that same
   lowercase email and pick their section → Assign CR.
3. Give the CR their email + password. When they log in, they'll only see
   their own section, and only get Mark / Save / Download — no add/delete
   controls.

To remove a CR's access later: Admin panel → "Remove" next to their name.
That revokes app access but does **not** delete their Firebase Auth login —
do that separately in the Console if you want the login gone entirely.

---

## What Admin can do
- Switch between any section
- Add / delete students in any section
- Add / delete subjects (shared across all sections)
- Assign / remove CRs
- Mark, save, and download attendance for any section

## What CR can do
- See only their assigned section (fixed, can't switch)
- Mark, save, and download attendance for that section
- Cannot add or delete students or subjects — those buttons don't
  exist for CR, and even if someone bypassed the UI, Firestore Rules
  would block the write

## Requests and Activity Logs
The app stores CR change requests in `requests` and Admin-visible activity history in `activityLogs`. After deploying the latest files, publish the matching `firestore.rules` in Firebase Console before testing these features.


## New Section setup

From Admin -> Update / Manage Data -> Add New Section, enter the class/mentor/CR details, download the Excel template, fill the Students and Subjects sheets, upload it, preview the counts, then create the section.

New sections are stored under the top-level `sections/{sectionId}` document and their students/subjects live in the existing subcollections. The Firestore rules include Admin create/update/delete access to section metadata and CR read access to the metadata for the CR's assigned section.

The CR account is provisioned from the Admin browser using a secondary Firebase Auth instance so the Admin session is not signed out. For a brand-new CR email, a temporary password is generated. The invitation email opens in Gmail as a pre-filled draft because the current app has no server-side email service.
