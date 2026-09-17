# QAttend Attendance Management System

> A Firebase-powered attendance management portal for Quantum University
> with separate Login, Admin, and Class Representative (CR) interfaces.

------------------------------------------------------------------------

## 1. Project Overview

**QAttend** is an attendance management system designed to manage
university courses, sections, students, subjects, Class Representatives
(CRs), attendance records, requests, and activity history.

The project is structured into separate pages so that each role gets
only the interface and features relevant to that role.

### Main roles

-   **Guest / Login User**
    -   Login
    -   Forgot Password
    -   Contact Admin
    -   Request a New Section
-   **Admin**
    -   Dashboard
    -   Course management
    -   Section management
    -   Student management
    -   Subject management
    -   Section / CR requests
    -   Attendance management
    -   Activity logs
    -   Profile
    -   Logout
-   **Class Representative (CR)**
    -   Assigned-section attendance
    -   Previous attendance records
    -   Send Request
    -   Help & Support
    -   Logout

------------------------------------------------------------------------

# 2. Project Structure

``` text
QAttend/
│
├── index.html
├── login.html
├── admin.html
├── cr.html
│
├── style.css
├── script.js
│
├── logo.png
├── firebase.json
├── firestore.rules
├── Section_Template.xlsx        # optional/static template if included
└── README.md
```

### File responsibilities

  -----------------------------------------------------------------------
  File                                Purpose
  ----------------------------------- -----------------------------------
  `index.html`                        Entry page that redirects users to
                                      `login.html`

  `login.html`                        Login, Forgot Password, Contact
                                      Admin and New Section Request

  `admin.html`                        Admin Portal

  `cr.html`                           CR Portal

  `style.css`                         Shared styling for all pages

  `script.js`                         Firebase Authentication, Firestore,
                                      Excel processing, requests,
                                      attendance and admin/CR logic

  `logo.png`                          QAttend logo

  `firebase.json`                     Firebase Hosting configuration

  `firestore.rules`                   Firestore security rules

  `README.md`                         Project documentation
  -----------------------------------------------------------------------

The current pages already reference the common stylesheet, and the
Login/Admin/CR pages use page-specific `data-page` values for
role/page-aware JavaScript. The project also keeps Firebase and Excel
functionality in the shared JavaScript file.

------------------------------------------------------------------------

# 3. Technology Stack

### Frontend

-   HTML5
-   CSS3
-   JavaScript ES Modules
-   Responsive design

### Backend / Cloud

-   Firebase Authentication
-   Cloud Firestore
-   Firebase Hosting

### Additional libraries/services

-   Firebase Web SDK `12.18.0`
-   SheetJS for Excel import/export
-   EmailJS for direct email sending

Firebase is initialized in `script.js`, and the current project uses
Firebase Authentication and Firestore from the Firebase CDN.

------------------------------------------------------------------------

# 4. Authentication

QAttend uses Firebase Email/Password Authentication.

### Login flow

``` text
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

The application uses browser-local Firebase Auth persistence so a valid
login can remain available across page navigation/reloads.

### Important

The Admin and CR pages should not be treated as public application
pages. Access is checked through Firebase Authentication and the user's
Firestore profile.

------------------------------------------------------------------------

# 5. Login Page

`login.html` contains:

-   QAttend branding
-   Email field
-   Password field
-   Show/hide password
-   Login
-   Forgot Password
-   Need access? Contact Admin
-   Request a new section

The Login page also contains the Contact Admin and New Section request
modals.

------------------------------------------------------------------------

# 6. Forgot Password

The Forgot Password feature uses Firebase's password reset email flow.

Basic flow:

``` text
Enter email
   ↓
Firebase sendPasswordResetEmail()
   ↓
Password reset email
   ↓
User changes password
```

------------------------------------------------------------------------

# 7. Contact Admin

The Contact Admin form is available before login.

### Fields

-   Name
-   Email
-   Mobile Number
-   Message

The request is stored in the Firestore `requests` collection with:

``` text
type: contact_admin
status: pending
```

The configured Admin recipient is:

``` text
sonusin8672@gmail.com
```

### Email

Direct email sending is intended to use EmailJS.

The current `script.js` contains placeholders for:

``` text
EMAILJS public key
EMAILJS service ID
EMAILJS template ID
```

These values must be configured before direct email sending will work.

The application is specifically designed to send the email directly
through the email service rather than opening Gmail or another
mail-compose window.

------------------------------------------------------------------------

# 8. Request a New Section

A guest can request a new academic section before logging in.

### Section information

The form collects:

-   CR Q.ID
-   CR Name
-   CR Mobile
-   CR Email
-   Mentor Name
-   Mentor Mobile
-   Course
-   Section
-   Semester
-   Year

### Excel workflow

``` text
Download Template
       ↓
Fill Students
       ↓
Fill Subjects
       ↓
Upload Excel
       ↓
Validate
       ↓
Preview
       ↓
Submit Request
       ↓
Admin Request Center
```

The current Excel generator creates:

### Sheet 1: Students

``` text
QID | Student Name
```

### Sheet 2: Subjects

``` text
Subject Name
```

### Sheet 3: Instructions

Contains rules for entering QIDs, student names, subjects, duplicate
QIDs and file format.

The current JavaScript also validates duplicate QIDs and invalid student
rows and prepares student/subject preview data.

------------------------------------------------------------------------

# 9. Admin Portal

`admin.html` is the Admin interface.

### Main Admin areas

``` text
Dashboard
Manage Courses
Manage Sections
Manage Students
Manage Subjects
Section Requests
Attendance
Logs
Profile
Logout
```

------------------------------------------------------------------------

# 10. Admin Dashboard

The Dashboard provides an overview of the system.

### Main statistics

-   Total Courses
-   Total Sections
-   Total Students
-   Pending Requests

### Dashboard sections

-   Recent Section Requests
-   Activity Logs

The dashboard data is loaded from Firestore rather than being intended
as static demo data.

------------------------------------------------------------------------

# 11. Manage Courses

Admin can manage academic courses.

### Course fields

-   Course Name
-   Course Code
-   Department
-   Duration

Course data is stored under:

``` text
courses/{courseId}
```

Typical operations:

``` text
Create Course
Edit Course
Delete Course
View Course
```

------------------------------------------------------------------------

# 12. Manage Sections

Admin can create and manage sections belonging to courses.

### Section information

-   Course
-   Section
-   Semester
-   Year
-   Mentor Name
-   Mentor Mobile

A section can also be populated from an Excel file so that students and
subjects do not have to be entered manually one by one.

### Section creation flow

``` text
Select Course
     ↓
Add Section
     ↓
Enter Section Details
     ↓
Upload Excel
     ↓
Preview Students + Subjects
     ↓
Create Section
```

------------------------------------------------------------------------

# 13. Students

Students belong to a section.

Firestore structure:

``` text
sections
   └── {sectionId}
        └── students
             └── {studentId}
```

Student data includes:

``` text
QID
Name
```

Admin can:

-   View students
-   Add students
-   Edit students
-   Delete students

------------------------------------------------------------------------

# 14. Subjects

Subjects also belong to a section.

Firestore structure:

``` text
sections
   └── {sectionId}
        └── subjects
             └── {subjectId}
```

Subject data includes:

``` text
name
```

Admin can:

-   View subjects
-   Add subjects
-   Delete subjects

------------------------------------------------------------------------

# 15. CR Assignment

A CR is assigned to a section.

The user profile stored under `users/{email}` contains role and section
information.

Example profile structure:

``` text
users/{crEmail}

role: cr
email: cr@example.com
qid: ...
name: ...
mobile: ...
section: ...
sectionLabel: ...
course: ...
semester: ...
year: ...
mentorName: ...
mentorMobile: ...
```

The project uses a secondary Firebase Authentication instance during CR
provisioning so that creating a CR account does not unnecessarily sign
the logged-in Admin out.

------------------------------------------------------------------------

# 16. CR Portal

`cr.html` is dedicated to the Class Representative.

### CR features

-   Mark Attendance
-   Previous Records
-   Send Request
-   Help & Support
-   Logout

The CR is restricted to the section assigned in the CR's Firestore
profile.

------------------------------------------------------------------------

# 17. CR Send Request

The CR hamburger menu includes:

``` text
Mark Attendance
Previous Records
Send Request
Help & Support
Logout
```

### Send Request options

``` text
Add Student
Delete Student
Add Subject
Delete Subject
```

### Add Student

Fields:

-   Student Q.ID
-   Student Name
-   Reason

### Delete Student

Fields:

-   Select Student
-   Reason

### Add Subject

Fields:

-   Subject Name
-   Reason

### Delete Subject

Fields:

-   Select Subject
-   Reason

All CR change requests are submitted as pending requests for Admin
review.

------------------------------------------------------------------------

# 18. Request Approval System

Requests are stored in:

``` text
requests/{requestId}
```

Request records can contain information such as:

``` text
type
status
requestedBy
requestedByRole
section
sectionLabel
reason
createdAt
```

### Request types

``` text
contact_admin
add_section
add_student
delete_student
add_subject
delete_subject
```

### Admin workflow

``` text
CR / Guest submits request
          ↓
       Pending
          ↓
    Admin reviews
       ↙     ↘
  Approve     Reject
     ↓           ↓
Apply change   No data change
     ↓
Status = approved
```

For academic change requests, the intended behavior is that the
underlying student/subject/section change happens only after Admin
approval.

------------------------------------------------------------------------

# 19. Attendance

Attendance is stored under the selected section.

Firestore structure:

``` text
sections
   └── {sectionId}
        └── attendance
             └── {recordId}
```

### Attendance workflow

``` text
Select Section
      ↓
Select Date
      ↓
Select Subject
      ↓
Load Students
      ↓
Mark Present / Absent
      ↓
Save Attendance
```

CRs work only with their assigned section.

Admin can manage attendance across sections.

------------------------------------------------------------------------

# 20. Previous Records

Saved attendance records can be viewed from the attendance/records area.

Typical record information includes:

-   Section
-   Subject
-   Date
-   Student attendance
-   Present/Absent status

The project also contains Excel download/export functionality for
attendance records.

------------------------------------------------------------------------

# 21. Excel Attendance Export

The project uses SheetJS for Excel processing.

To improve initial page speed, the Excel library is loaded lazily when
Excel functionality is actually required rather than being loaded
immediately on every page.

The attendance export keeps the student row order and S.No. associated
with the student's current row/order.

The project also supports the attendance shortcut row used by the
existing Excel workflow.

------------------------------------------------------------------------

# 22. Activity Logs

Activity logs are stored in:

``` text
activityLogs/{logId}
```

Examples of logged actions include:

``` text
Logged in
Logged out
ADD_STUDENT
EDIT_STUDENT
DELETE_STUDENT
ADD_SUBJECT
DELETE_SUBJECT
Attendance actions
Request actions
```

The Admin can read the activity history.

Regular users can append their own activity entries but cannot read or
modify the complete Admin activity history.

------------------------------------------------------------------------

# 23. Firestore Data Model

The main Firestore collections are:

``` text
users
courses
sections
requests
activityLogs
```

Section subcollections:

``` text
sections/{sectionId}/students
sections/{sectionId}/subjects
sections/{sectionId}/attendance
```

### Simplified structure

``` text
users/
  admin@example.com
  cr@example.com

courses/
  BTECH-CSE

sections/
  BTECH-CSE-CSE-1
      students/
      subjects/
      attendance/

requests/
  requestId

activityLogs/
  logId
```

------------------------------------------------------------------------

# 24. Firestore Security

The current rules distinguish between Admin, CR and guest access.

### Admin

Admin has management access to:

-   Users
-   Courses
-   Sections
-   Students
-   Subjects
-   Attendance
-   Requests
-   Activity Logs

### CR

CR can access the section assigned to the CR.

CR can:

-   Read assigned section data
-   Work with attendance for assigned section
-   Submit change requests
-   Read their own submitted requests

CR cannot directly modify students or subjects.

### Guest

Before login, the rules permit only the intended request creation flows:

-   New Section request
-   Contact Admin request

Guest users cannot read academic data.

The current Firestore rules explicitly implement these role restrictions
and request permissions.

------------------------------------------------------------------------

# 25. Firebase Configuration

`firebase.json` configures Firebase Hosting.

The current Hosting configuration uses:

``` text
target: qattend
public: .
```

HTML, JavaScript and CSS are configured with no-cache headers, while
common image formats are cached for one week.

### Important

Do not replace Firestore rules with:

``` text
allow read, write: if true;
```

That would remove the intended access restrictions.

Always deploy the tested `firestore.rules`.

------------------------------------------------------------------------

# 26. EmailJS Configuration

Open:

``` text
script.js
```

Find:

``` javascript
const EMAILJS_CONFIG = {
    publicKey: "YOUR_EMAILJS_PUBLIC_KEY",
    serviceId: "YOUR_EMAILJS_SERVICE_ID",
    templateId: "YOUR_EMAILJS_TEMPLATE_ID",
    adminEmail: "sonusin8672@gmail.com"
};
```

Replace only:

``` text
YOUR_EMAILJS_PUBLIC_KEY
YOUR_EMAILJS_SERVICE_ID
YOUR_EMAILJS_TEMPLATE_ID
```

with the values from your EmailJS project.

Do not change:

``` text
adminEmail
```

unless the Admin recipient is intentionally changed.

------------------------------------------------------------------------

# 27. Local Development

Do not open the project using:

``` text
file://
```

Use a local HTTP server.

### Option 1: VS Code Live Server

1.  Open the QAttend folder in VS Code.
2.  Install/use Live Server.
3.  Right-click `login.html`.
4.  Select **Open with Live Server**.
5.  Open the displayed localhost address.

### Option 2: Python

From the project folder:

``` bash
python3 -m http.server 5500
```

Then open:

``` text
http://localhost:5500/login.html
```

------------------------------------------------------------------------

# 28. Local Testing Checklist

Before deploying, test:

### Login

-   [ ] Admin login
-   [ ] CR login
-   [ ] Wrong password
-   [ ] Forgot password
-   [ ] Logout

### Contact Admin

-   [ ] Open before login
-   [ ] Fill form
-   [ ] Submit
-   [ ] Firestore request appears
-   [ ] Email service sends mail

### New Section

-   [ ] Open before login
-   [ ] Download template
-   [ ] Fill Students sheet
-   [ ] Fill Subjects sheet
-   [ ] Upload file
-   [ ] Student count appears
-   [ ] Subject count appears
-   [ ] Preview appears
-   [ ] Submit request
-   [ ] Admin can see request

### Admin

-   [ ] Dashboard
-   [ ] Course create/edit/delete
-   [ ] Section create
-   [ ] Excel import
-   [ ] Student management
-   [ ] Subject management
-   [ ] CR assignment
-   [ ] Request approval
-   [ ] Request rejection
-   [ ] Attendance
-   [ ] Logs

### CR

-   [ ] Assigned section loads
-   [ ] Subject selection works
-   [ ] Attendance saves
-   [ ] Previous records work
-   [ ] Hamburger opens
-   [ ] Send Request appears
-   [ ] Add Student request
-   [ ] Delete Student request
-   [ ] Add Subject request
-   [ ] Delete Subject request
-   [ ] Logout

### Mobile

-   [ ] Login page
-   [ ] Contact Admin modal
-   [ ] New Section modal
-   [ ] CR Send Request modal
-   [ ] Admin sidebar
-   [ ] Buttons visible
-   [ ] Close/Cancel/Submit buttons visible
-   [ ] No horizontal overflow

------------------------------------------------------------------------

# 29. Deployment to Firebase Hosting

The current project is configured for Firebase Hosting.

Typical deployment:

``` bash
firebase login
firebase use <your-project>
firebase deploy --only hosting
```

If Firestore rules were changed:

``` bash
firebase deploy --only firestore:rules
```

Or deploy both:

``` bash
firebase deploy --only hosting,firestore
```

### Important

If the GitHub repository is connected to Firebase Hosting through GitHub
Actions, pushing changes to the configured branch can trigger the
Firebase Hosting deployment workflow.

Before pushing:

``` bash
git status
git add .
git commit -m "Update QAttend"
git push
```

Then check the GitHub Actions workflow for the deployment result.

------------------------------------------------------------------------

# 30. Recommended Project Workflow

Use this development order:

``` text
1. Local Login Testing
        ↓
2. Firebase Authentication
        ↓
3. Firestore Rules
        ↓
4. Admin Dashboard
        ↓
5. Course Management
        ↓
6. Section Management
        ↓
7. Student / Subject Management
        ↓
8. CR Assignment
        ↓
9. CR Attendance
        ↓
10. CR Send Request
        ↓
11. Admin Request Approval
        ↓
12. Activity Logs
        ↓
13. Excel Testing
        ↓
14. Mobile Testing
        ↓
15. GitHub Push
        ↓
16. Firebase Hosting Deployment
```

------------------------------------------------------------------------

# 31. Troubleshooting

### Firebase Rules Error

Check:

``` text
firestore.rules
```

Make sure the rules are deployed:

``` bash
firebase deploy --only firestore:rules
```

### Login works but wrong page opens

Check the user's Firestore profile:

``` text
users/{email}
```

Verify:

``` text
role: admin
```

or:

``` text
role: cr
```

For a CR, also verify the assigned:

``` text
section
```

### Contact Admin email does not send

Check EmailJS configuration in `script.js`.

If the values still start with:

``` text
YOUR_
```

EmailJS is not configured.

### Excel does not load

Check internet access because SheetJS is loaded on demand from the
SheetJS CDN.

### Changes do not appear after deployment

Hard refresh the browser and check Firebase Hosting deployment status.
The current Hosting configuration deliberately uses no-cache headers for
HTML, JS and CSS.

------------------------------------------------------------------------

# 32. Security Notes

Never put:

-   Firebase Admin SDK service-account private keys
-   Service-account JSON files
-   EmailJS private secrets
-   Other server-side secrets

inside the public frontend project.

Firebase client configuration is intended for the web application, but
access control must still be enforced through Firestore Security Rules
and Firebase Authentication.

Never rely only on hiding buttons in HTML/CSS for security.

------------------------------------------------------------------------

# 33. Project Status

The project is organized around:

-   Separate Login page
-   Separate Admin page
-   Separate CR page
-   Shared CSS
-   Shared JavaScript
-   Firebase Authentication
-   Firestore
-   Role-based access
-   Attendance management
-   Excel import/export
-   Section requests
-   CR change requests
-   Admin approval workflow
-   Activity logging
-   Firebase Hosting

The current project documentation identifies the same page separation
and shared `style.css` / `script.js` structure.

------------------------------------------------------------------------

# 34. Quick Reference

### Start locally

``` bash
python3 -m http.server 5500
```

Open:

``` text
http://localhost:5500/login.html
```

### Deploy Hosting

``` bash
firebase deploy --only hosting
```

### Deploy Firestore Rules

``` bash
firebase deploy --only firestore:rules
```

### Deploy both

``` bash
firebase deploy --only hosting,firestore
```

------------------------------------------------------------------------

## QAttend

**Attendance Management System**

Built for structured course, section, student, subject, CR and
attendance management with Firebase.
