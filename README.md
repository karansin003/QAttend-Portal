<div align="center">

<img src="logo.png" alt="QAttend Logo" width="120" height="126" />

# QAttend

### Attendance Management System — Quantum University

A Firebase-powered attendance portal with dedicated **Login**, **Admin**, and **Class Representative (CR)** interfaces.

[![Firebase](https://img.shields.io/badge/Firebase-Auth%20%2B%20Firestore-FFCA28?logo=firebase&logoColor=black)](https://firebase.google.com/)
[![Hosting](https://img.shields.io/badge/Firebase-Hosting-orange?logo=firebase)](https://firebase.google.com/products/hosting)
[![HTML5](https://img.shields.io/badge/HTML5-E34F26?logo=html5&logoColor=white)](https://developer.mozilla.org/en-US/docs/Web/HTML)
[![CSS3](https://img.shields.io/badge/CSS3-1572B6?logo=css3&logoColor=white)](https://developer.mozilla.org/en-US/docs/Web/CSS)
[![JavaScript](https://img.shields.io/badge/JavaScript-ES%20Modules-F7DF1E?logo=javascript&logoColor=black)](https://developer.mozilla.org/en-US/docs/Web/JavaScript)
[![License](https://img.shields.io/badge/License-Private-lightgrey)]()

</div>

---

## 📖 Table of Contents

- [Overview](#-overview)
- [Roles at a Glance](#-roles-at-a-glance)
- [Project Structure](#-project-structure)
- [Tech Stack](#-tech-stack)
- [Authentication Flow](#-authentication-flow)
- [Feature Highlights](#-feature-highlights)
  - [Login Page](#login-page)
  - [Admin Portal](#admin-portal)
  - [CR Portal](#cr-portal)
- [Firestore Data Model](#-firestore-data-model)
- [Excel Import/Export](#-excel-importexport)
- [Getting Started](#-getting-started)
- [First Admin Setup](#-first-admin-setup)
- [One-Time Data Import](#-one-time-data-import)
- [EmailJS Configuration](#-emailjs-configuration)
- [Deployment](#-deployment)
- [Security Notes](#-security-notes)
- [Troubleshooting](#-troubleshooting)
- [Testing Checklist](#-testing-checklist)

---

## 🧭 Overview

**QAttend** manages university courses, sections, students, subjects, Class
Representatives (CRs), attendance records, requests, and activity history —
all backed by **Firebase Authentication** and **Cloud Firestore**.

The project is split into role-specific pages so every user only sees the
interface relevant to them:

| Page | Role | Purpose |
|---|---|---|
| `login.html` | Guest | Login, Forgot Password, Contact Admin, Request New Section |
| `admin.html` | Admin | Full system management |
| `cr.html` | Class Representative | Attendance for their assigned section only |

---

## 👥 Roles at a Glance

<table>
<tr>
<td valign="top" width="33%">

### 🔑 Guest
- Login
- Forgot Password
- Contact Admin
- Request a New Section

</td>
<td valign="top" width="33%">

### 🛡️ Admin
- Dashboard & stats
- Manage Courses / Sections
- Manage Students / Subjects
- Requests & approvals
- Manage Users / CR assignment
- Attendance (any section)
- Activity Logs & Reports (CSV)

</td>
<td valign="top" width="33%">

### 🎓 Class Representative (CR)
- Fixed, assigned section only
- Mark / Save / Download attendance
- View previous records
- Send change requests
- Help & Support

</td>
</tr>
</table>

---

## 🗂 Project Structure

```text
QAttend/
│
├── index.html              # Redirects to login.html
├── login.html               # Guest entry point
├── admin.html                # Admin portal
├── cr.html                    # CR portal
│
├── style.css                 # Shared styling for all pages
├── script.js                 # Firebase Auth, Firestore, Excel, admin/CR logic
│
├── logo.png                  # QAttend logo
├── firebase.json             # Firebase Hosting configuration
├── firestore.rules           # Firestore security rules
│
├── seed.html                 # One-time roster/subject import UI
├── seed.js                   # One-time import script
├── seed-backup.js            # Backup copy of import script
│
├── Section_Template.xlsx     # Excel template (Students / Subjects / Instructions)
└── README.md
```

---

## ⚙️ Tech Stack

**Frontend**
- HTML5, CSS3, JavaScript (ES Modules)
- Responsive, mobile-tested layout

**Backend / Cloud**
- Firebase Authentication (Email/Password)
- Cloud Firestore
- Firebase Hosting

**Libraries & Services**
- Firebase Web SDK `12.18.0`
- [SheetJS](https://sheetjs.com/) — Excel import/export
- [EmailJS](https://www.emailjs.com/) — direct email sending (no mail-client popup)

---

## 🔐 Authentication Flow

```text
login.html
     │
     ▼
Firebase Authentication
     │
     ▼
Read users/{email}
     │
     ├── role = admin ──► Admin Portal
     │
     └── role = cr ─────► CR Portal
```

- Uses browser-local Firebase Auth persistence — login survives page reloads.
- Admin/CR pages are **not public**; access is enforced through Firebase Auth
  **and** the user's Firestore profile — never rely on hiding UI elements alone.
- ⚠️ **All account emails must be lowercase.** Firebase does not enforce
  casing, and a mismatch between the Auth login email and the Firestore role
  document will leave the account unassigned.

---

## ✨ Feature Highlights

### Login Page

- QAttend branding, email/password fields, show/hide password
- **Forgot Password** — Firebase `sendPasswordResetEmail()` flow
- **Contact Admin** modal — Name, Email, Mobile, Message → stored in
  `requests` (`type: contact_admin`, `status: pending`) and sent via EmailJS
- **Request a New Section** modal — collects CR & Mentor details, course,
  section, semester, year, plus an Excel upload for students/subjects

### Admin Portal

```text
Dashboard → Manage Courses → Manage Sections → Manage Students
→ Manage Subjects → Requests → Manage Users → Attendance → Reports → Logs → Profile
```

- **Dashboard** — live stats (courses, sections, students, pending requests),
  recent requests, and activity feed
- **Requests table** — All / Pending / Students / Subjects / New Sections /
  Contact Admin filters, with View / Approve / Reject and bulk selection
- Approving **Add/Delete Student or Subject** applies the change directly
- Approving **New Section** creates the section, imports its roster/subjects,
  and assigns the requesting CR
- Rejecting a request only updates its status — section data is untouched
- **Manage Sections** — add/edit/delete section metadata (fixed system
  sections are protected from deletion)
- **Manage Users** — add/update profiles, activate/deactivate, assign/remove CRs
  *(Firebase Auth accounts are intentionally created via Console, not client code)*
- **Reports** — CSV export for attendance, requests, students, subjects
  (loaded on demand)

### CR Portal

- Locked to their **assigned section only**
- Mark, save, and download attendance
- View previous attendance records
- **Send Request** for Add/Delete Student or Subject (Admin-approved)
- Cannot add/delete students or subjects directly — enforced both in the UI
  and in Firestore Security Rules

---

## 🗄 Firestore Data Model

```text
users/{email}                     → { role: "admin" | "cr", section?, ... }
courses/{courseId}                → { name, code, department, duration }
sections/{sectionId}
   ├── students/{studentId}       → { QID, name, ... }
   └── attendance/...
subjects/{subjectId}              → shared across sections
requests/{requestId}              → { type, status, ... }
activityLogs/{logId}              → Admin-visible history
```

> Publish `firestore.rules` **before** testing Contact Admin, New Section, or
> any Firestore-backed feature — without it, access control is not enforced.

---

## 📊 Excel Import/Export

Used for bulk-creating a new section's roster and subject list.

```text
Download Template → Fill Students → Fill Subjects
→ Upload Excel → Validate → Preview → Submit Request → Admin Request Center
```

`Section_Template.xlsx` contains three sheets:

| Sheet | Columns |
|---|---|
| **Students** | `QID` \| `Student Name` |
| **Subjects** | `Subject Name` |
| **Instructions** | Rules for QIDs, names, subjects, duplicates, file format |

Duplicate QIDs and invalid rows are automatically flagged/skipped during validation.

---

## 🚀 Getting Started

Do **not** open the project with `file://` — use a local HTTP server.

**Option 1 — VS Code Live Server**
1. Open the project folder in VS Code
2. Right-click `login.html` → **Open with Live Server**

**Option 2 — Python**
```bash
python3 -m http.server 5500
```
Then open:
```text
http://localhost:5500/login.html
```

---

## 🏁 First Admin Setup

1. **Enable Firestore** — Firebase Console → Build → Firestore Database → Create database (production mode, any region)
2. **Publish security rules** — Firestore → Rules tab → paste `firestore.rules` → Publish
3. **Create the first Admin account** *(manual, one-time)*:
   - Authentication → Users → Add user → lowercase email + password
   - Firestore → Data → new collection `users`, document ID = the same lowercase email, field `role` (string) = `admin`
4. **Deploy site files** — `index.html`, `style.css`, `script.js`, `seed.html`, `seed.js`, `logo.png`, `firebase.json`
5. Log in with the Admin account

---

## 📥 One-Time Data Import

While logged in as Admin, open `yoursite.com/seed.html` and click **Run Import**.

- Loads the full roster (539 students across Sections 1–8, AIML-1, AIML-2) and the 11 subjects
- Safe to run more than once — only adds what's missing, never duplicates
- After it finishes, `seed.html` / `seed.js` can be deleted — they're not linked from the main app
- No roster exists yet for a **CSCQ** section — add those students manually via the Admin panel once available

---

## ✉️ EmailJS Configuration

In `script.js`:

```javascript
const EMAILJS_CONFIG = {
    publicKey: "YOUR_EMAILJS_PUBLIC_KEY",
    serviceId: "YOUR_EMAILJS_SERVICE_ID",
    templateId: "YOUR_EMAILJS_TEMPLATE_ID",
    adminEmail: "sonusin8672@gmail.com"
};
```

Replace only the `YOUR_...` placeholders with values from your EmailJS
project. Leave `adminEmail` unless the recipient is intentionally changing.

---

## ☁️ Deployment

```bash
firebase login
firebase use <your-project>
firebase deploy --only hosting          # site files
firebase deploy --only firestore:rules  # if rules changed
firebase deploy --only hosting,firestore
```

If the repo is connected to Firebase Hosting via GitHub Actions, pushing to
the configured branch triggers deployment automatically:

```bash
git add .
git commit -m "Update QAttend"
git push
```

---

## 🔒 Security Notes

- Never commit Firebase Admin SDK service-account keys, EmailJS secrets, or
  any server-side credentials to this repo
- Firestore Security Rules + Firebase Authentication are the real access
  control — never rely on hiding buttons in HTML/CSS
- Never replace rules with `allow read, write: if true;`

---

## 🛠 Troubleshooting

| Symptom | Check |
|---|---|
| Firestore permission errors | `firestore.rules` deployed via `firebase deploy --only firestore:rules` |
| Login works, wrong portal opens | `users/{email}` document — correct `role`, and `section` for CRs |
| Contact Admin email not sending | EmailJS config in `script.js` — values shouldn't start with `YOUR_` |
| Excel won't load | Internet access — SheetJS loads from CDN on demand |
| Changes not appearing after deploy | Hard refresh — HTML/JS/CSS are served with no-cache headers |

---

## ✅ Testing Checklist

<details>
<summary><strong>Login</strong></summary>

- [ ] Admin login  
- [ ] CR login  
- [ ] Wrong password  
- [ ] Forgot password  
- [ ] Logout
</details>

<details>
<summary><strong>Contact Admin</strong></summary>

- [ ] Opens before login  
- [ ] Submits successfully  
- [ ] Firestore request appears  
- [ ] Email service sends mail
</details>

<details>
<summary><strong>New Section Request</strong></summary>

- [ ] Template downloads  
- [ ] Students/Subjects sheets fill correctly  
- [ ] Upload validates counts  
- [ ] Preview appears  
- [ ] Admin sees the request
</details>

<details>
<summary><strong>Admin</strong></summary>

- [ ] Dashboard loads  
- [ ] Course CRUD  
- [ ] Section create + Excel import  
- [ ] Student/Subject management  
- [ ] CR assignment  
- [ ] Request approve/reject  
- [ ] Attendance  
- [ ] Logs
</details>

<details>
<summary><strong>CR</strong></summary>

- [ ] Assigned section loads  
- [ ] Attendance marks & saves  
- [ ] Previous records load  
- [ ] Send Request modal works  
- [ ] Logout
</details>

<details>
<summary><strong>Mobile</strong></summary>

- [ ] All modals usable on small screens  
- [ ] Admin sidebar responsive  
- [ ] No horizontal overflow
</details>

---

<div align="center">

**QAttend** · Built for structured course, section, student, subject, CR and attendance management with Firebase 🔥

</div>
