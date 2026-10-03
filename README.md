<div align="center">

<img src="logo.png" alt="QAttend Logo" width="120" height="126" />

# QAttend

### Attendance Management System — Quantum University

A Firebase-powered attendance portal with dedicated **Login**, **Admin**, **Class Representative (CR)**, and **Teacher** interfaces with full **Offline-First PWA** support.

[![Firebase](https://img.shields.io/badge/Firebase-Auth%20%2B%20Firestore-FFCA28?logo=firebase&logoColor=black)](https://firebase.google.com/)
[![Hosting](https://img.shields.io/badge/Firebase-Hosting-orange?logo=firebase)](https://firebase.google.com/products/hosting)
[![PWA](https://img.shields.io/badge/PWA-Offline%20First-5A0FC8?logo=pwa&logoColor=white)]()
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
- [Authentication & Role Routing](#-authentication--role-routing)
- [Offline-First & PWA Architecture](#-offline-first--pwa-architecture)
- [Feature Highlights](#-feature-highlights)
  - [Login Page](#login-page)
  - [Admin Portal](#admin-portal)
  - [CR Portal](#cr-portal)
  - [Teacher Portal](#teacher-portal)
- [Firestore Data Model & Security Rules](#-firestore-data-model--security-rules)
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

**QAttend** manages university courses, sections, students, subjects, Class Representatives (CRs), Teachers, attendance records, requests, and activity history — all backed by **Firebase Authentication**, **Cloud Firestore**, and an **offline-first Service Worker**.

The project is split into role-specific pages so every user only sees the interface relevant to them:

| Page | Role | Purpose |
|---|---|---|
| `login.html` | Guest | Login, Forgot Password, Contact Admin, Request New Section |
| `admin.html` | Admin | Full system management, course/section setup, CR & Teacher assignments |
| `cr.html` | Class Representative | Attendance for their single assigned section only |
| `teacher.html` | Teacher | Attendance for their assigned `(section, subject)` pairs only |

---

## 👥 Roles at a Glance

<table>
<tr>
<td valign="top" width="25%">

### 🔑 Guest
- Login
- Modern SVG eye password toggle
- Forgot Password
- Contact Admin
- Request a New Section

</td>
<td valign="top" width="25%">

### 🛡️ Admin
- Live Dashboard & stats
- Courses & Sections CRUD
- Manage Students & Subjects
- Requests approval/rejection
- Assign / Manage CRs
- Assign / Manage Teachers
- Attendance for any section
- Activity Logs & Reports (CSV)

</td>
<td valign="top" width="25%">

### 🎓 Class Representative
- Locked to single assigned section
- Mark / Save / Download attendance
- View section records history
- Send change requests to Admin
- Offline attendance support
- Help & Support

</td>
<td valign="top" width="25%">

### 👨‍🏫 Teacher
- Locked to assigned `(section, subject)` pairs
- Auto-lands on assigned section+subject
- Multi-assignment switcher dropdown
- Locked subject display (no switching)
- Records history filtered by assigned subject
- Full offline attendance marking
- Offline auto-sync on reconnect

</td>
</tr>
</table>

---

## 🗂 Project Structure

```text
QAttend/
│
├── index.html              # Redirects to login.html
├── login.html               # Guest entry point (login, forgot pwd, contact admin, request section)
├── admin.html                # Admin portal (full system management, user & role assignments)
├── cr.html                    # CR portal (section attendance)
├── teacher.html               # Teacher portal (subject & section locked attendance)
│
├── style.css                 # Shared responsive stylesheet for all pages & roles
├── script.js                 # Unified Firebase Auth, Firestore, PWA, SheetJS, & role logic
├── service-worker.js         # Service Worker for offline app shell caching & PWA support
│
├── logo.png                  # QAttend logo
├── firebase.json             # Firebase Hosting headers and rules configuration
├── firestore.rules           # Security rules for collections, roles, and attendance writes
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
- HTML5, Vanilla CSS3, JavaScript (ES Modules)
- Modern SVG icons (clean eye and eye-off password view toggles)
- Fully responsive layout for Desktop, Tablet, and Mobile

**Offline & PWA**
- Service Worker (`service-worker.js`) caching local App Shell (`CACHE_VERSION: qattend-v2`)
- Firestore IndexedDB local cache (`persistentLocalCache` + `persistentMultipleTabManager`)
- Auth persistence (`browserLocalPersistence`)
- Profile caching (`localStorage: qattend-profile`) for instant offline launch
- Automatic background write synchronization on network reconnection

**Backend & Cloud**
- Firebase Authentication (Email/Password)
- Cloud Firestore
- Firebase Hosting

**Libraries & Services**
- Firebase Web SDK `12.18.0`
- [SheetJS](https://sheetjs.com/) — On-demand Excel import/export (XLSX)
- [EmailJS](https://www.emailjs.com/) — Direct email delivery for admin requests

---

## 🔐 Authentication & Role Routing

```text
login.html
     │
     ▼
Firebase Authentication
     │
     ▼
Read users/{email}
     │
     ├── role = "admin" ───► admin.html (Admin Portal)
     ├── role = "cr" ──────► cr.html (CR Portal)
     └── role = "teacher" ─► teacher.html (Teacher Portal)
```

- **Separate Page Routing**: Each portal page validates user role upon loading. If a user visits the wrong page or attempts unauthorized access, they are automatically redirected to their dedicated page.
- **Offline Authentication**: When offline, Firebase Auth restores the session from IndexedDB, and the app reads the cached profile from `localStorage` (`qattend-profile`), allowing immediate access to `admin.html`, `cr.html`, or `teacher.html` without internet connectivity.
- ⚠️ **All account emails must be lowercase.** Firebase casing mismatches between Auth and Firestore documents can leave users unassigned.

---

## ⚡ Offline-First & PWA Architecture

1. **Service Worker App Shell**:
   - `service-worker.js` caches `index.html`, `login.html`, `admin.html`, `cr.html`, `teacher.html`, `style.css`, `script.js`, and `logo.png`.
   - Navigation fallback serves cached pages seamlessly when offline.
2. **Firestore Offline Persistence**:
   - Firestore's IndexedDB persistence maintains all viewed student rosters, subjects, and records locally.
   - Attendance markings saved offline are queued locally and automatically synchronized with Firestore when internet connectivity returns.
3. **Auto-Reconnect Refresh**:
   - The app listens for `qattend:network-restored` and refreshes active views automatically once pending sync completes.

---

## ✨ Feature Highlights

### Login Page

- Clean QAttend branding with university subtitle
- Email and password input with a **sleek SVG eye / eye-off toggle button** (open eye to reveal, slashed eye to hide)
- **Forgot Password** — Firebase `sendPasswordResetEmail()` workflow
- **Contact Admin** modal — Stored in `requests` collection and notified via EmailJS
- **Request a New Section** modal — Section, Mentor, Course details with Excel template upload

### Admin Portal

- **Live Dashboard**: Total counts for courses, sections, students, subjects, CRs, and pending requests
- **Manage Users**:
  - **Assign / Manage CR**: Assign a student CR to a specific section; view existing CRs with one-click revocation.
  - **Assign / Manage Teacher**: Assign teachers to `(section, subject)` pairs. The Subject dropdown automatically populates from active subjects in that section. Supports assigning multiple sections/subjects to the same teacher.
  - **Remove Teacher Assignment**: Remove individual section+subject pairs; removing the last assignment automatically deletes the teacher document.
- **Manage Courses & Sections**: Complete CRUD operations with Excel batch import.
- **Request Center**: Approve or reject CR and visitor requests with single or bulk actions.
- **Activity Logs**: Searchable audit log tracking logins, student edits, CR/Teacher assignments, and attendance saves with color-coded badges.

### CR Portal

- Locked to their single assigned section
- Mark attendance (Present, Absent, All Present toggle, date picker, notes)
- Save attendance, download Excel spreadsheet (`.xlsx`), and share records
- View previous attendance records for their section
- Submit student/subject change requests to Admin

### Teacher Portal (`teacher.html`)

- **Strict Subject & Section Scoping**:
  - **Single Assignment**: Auto-selects the assigned section and locks the subject field.
  - **Multiple Assignments**: Shows a clean assignment selector above the attendance table (`"Section Label — Subject Name"`). Selecting one switches the roster and locks the corresponding subject.
- **Non-Editable Subject**: Subject dropdown is replaced with a locked display badge (`#teacherSubjectDisplay`) to prevent accidental or unauthorized subject switching.
- **Filtered Attendance History**: The previous records section only displays attendance records matching the teacher's assigned subject in that section.
- **No Unauthorized Controls**: Administrative controls (student/subject management, section creation, CR/Teacher assignment) remain hidden.
- **Offline Attendance**: Mark and save attendance even without network; synced automatically when back online.

---

## 🗄 Firestore Data Model & Security Rules

### Data Shapes (`users/{email}`)

```json
// Admin
{ "role": "admin" }

// Class Representative
{ "role": "cr", "section": "SECTION-1" }

// Teacher
{
  "role": "teacher",
  "assignments": [
    { "section": "SECTION-1", "subject": "Design and Analysis of Algorithm" },
    { "section": "SECTION-3", "subject": "Python" }
  ],
  "sections": ["SECTION-1", "SECTION-3"]
}
```

### Security Rules Highlights (`firestore.rules`)

- **Self Profile Read**: `request.auth.token.email == email`
- **Admin Full Access**: `isAdmin()` grants full read/write across all collections.
- **CR Scoping**: Read students/subjects in `mySection()`; create/update/delete attendance only in `mySection()`.
- **Teacher Scoping**:
  - Read students and subjects in assigned sections via `isAssignedToSection(sectionId)`.
  - Read attendance in assigned sections.
  - Create and update attendance **only if** `request.resource.data.subject` matches an assigned subject in that section:
    ```javascript
    isTeacher() && isAssignedTo(sectionId, request.resource.data.subject)
    ```
  - Delete attendance only for their assigned subject.

---

## 📊 Excel Import/Export

- **Export**: Generates standard attendance spreadsheets formatted with Student QID, Name, Status (Present/Absent), Date, Subject, Section, and Notes using [SheetJS](https://sheetjs.com/).
- **Import**: `Section_Template.xlsx` allows bulk importing students and subjects during section creation requests.

---

## 🚀 Getting Started

Do **not** open the project using `file://` protocol — use a local HTTP server to allow ES modules and Service Worker registration.

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

1. **Enable Firestore Database** in Firebase Console (Production Mode).
2. **Deploy Security Rules**:
   ```bash
   firebase deploy --only firestore:rules
   ```
3. **Create the First Admin Account**:
   - Authentication → Add user → lowercase email + password
   - Firestore → Collection `users` → Document ID = `<email>` → `{ "role": "admin" }`
4. **Deploy Application**:
   ```bash
   firebase deploy --only hosting
   ```
5. Log in at `login.html` with your Admin credentials.

---

## 📥 One-Time Data Import

While logged in as Admin, navigate to `yoursite.com/seed.html` and click **Run Import** to populate initial student rosters and subjects. This script is idempotent and safe to run once.

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

---

## ☁️ Deployment

```bash
# Login to Firebase CLI
firebase login

# Select Project
firebase use <your-project>

# Deploy all services
firebase deploy

# Or deploy specific targets
firebase deploy --only hosting
firebase deploy --only firestore:rules
```

---

## 🔒 Security Notes

- Never commit private keys, Service Account JSONs, or environment secrets to git.
- Access control is strictly enforced in `firestore.rules` at the database level, preventing client-side tampering of roles or attendance subjects.
- Always enforce lowercase email addresses across Auth and Firestore.

---

## 🛠 Troubleshooting

| Symptom | Check |
|---|---|
| Firestore permission errors | Ensure latest `firestore.rules` are published via Firebase CLI |
| Login redirects to wrong page | Verify role and section/assignments in `users/{email}` |
| Teacher sees no subjects/sections | Verify `assignments` array has valid `{ section, subject }` objects |
| Offline mode not loading | Confirm Service Worker registered (`qattend-v2`) in browser Application tab |
| Excel import fails | Verify sheet column names match `Section_Template.xlsx` |

---

## ✅ Testing Checklist

<details>
<summary><strong>Admin</strong></summary>

- [ ] Dashboard counters and live stats load
- [ ] Create, edit, and delete courses & sections
- [ ] Assign CR to section & remove CR
- [ ] Assign Teacher to section & subject
- [ ] Assign same Teacher to additional section+subject
- [ ] Remove individual Teacher assignment
- [ ] Approve / Reject section and student change requests
- [ ] View activity logs with color-coded badges
</details>

<details>
<summary><strong>Teacher</strong></summary>

- [ ] Single-assignment Teacher lands directly on attendance screen
- [ ] Multiple-assignment Teacher sees assignment dropdown switcher
- [ ] Subject is locked to assigned subject (not editable)
- [ ] Admin controls and CR request buttons are removed
- [ ] Mark, save, and download attendance works normally
- [ ] Saved-records list shows only their own subject's history
- [ ] Offline attendance marking and saving works without internet
- [ ] Unauthorized attendance write to unassigned subject is blocked by Firestore rules
</details>

<details>
<summary><strong>Class Representative (CR)</strong></summary>

- [ ] Assigned section loads automatically
- [ ] Can mark, save, download, and share attendance
- [ ] Previous records load for their section
- [ ] Change request modal sends requests to Admin
- [ ] Offline attendance marks and queues writes
</details>

<details>
<summary><strong>Authentication & UI</strong></summary>

- [ ] Password toggle switches between clean Eye and Eye-Off SVG icons
- [ ] Forgot password email triggers reset email
- [ ] Mobile responsive layout with drawer navigation
- [ ] Offline PWA indicator displays network status
</details>

---

<div align="center">

**QAttend** · Built for structured course, section, student, subject, CR, and Teacher attendance management with Firebase 🔥

</div>
