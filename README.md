# 📚 QAttend Portal

<div align="center">

<img src="logo.png" alt="QAttend Logo" width="120"/>

### Smart Attendance Management System for Quantum University

A secure, role-based web application designed to simplify student attendance management for administrators and Class Representatives.

</div>

---

## 📌 About The Project

**QAttend Portal** is a web-based attendance management system developed to make the process of recording, managing, and exporting student attendance faster and more organized.

The system provides different access levels for **Administrators** and **Class Representatives (CRs)**. Administrators have complete control over students, subjects, sections, and CR assignments, while CRs can securely manage attendance only for their assigned section.

The application uses **Firebase Authentication** for secure login and **Cloud Firestore** for storing students, subjects, user roles, and attendance records.

---

# ✨ Features

## 🔐 Secure Authentication

* Firebase Authentication
* Email and password login
* Role-based access control
* Secure Admin and CR accounts
* Automatic session handling
* Logout functionality

---

## 👨‍💼 Admin Features

Administrators have complete control over the system.

### Admin can:

* ✅ Access all sections
* 👨‍🎓 Add students
* 🗑️ Delete students
* 📚 Add subjects
* 🗑️ Delete subjects
* 👨‍🏫 Assign Class Representatives
* ❌ Remove Class Representatives
* 📋 Mark attendance for any section
* 💾 Save attendance records
* 📊 View previous attendance records
* 🗑️ Delete attendance records
* 📥 Download attendance reports
* 📤 Share attendance reports

---

## 👨‍🎓 Class Representative Features

Class Representatives have restricted access for security.

### CR can:

* 📌 Access only their assigned section
* 👀 View students in their section
* 📚 View subjects for their section
* ✅ Mark students as Present or Absent
* ⚡ Mark all students as Present
* ❌ Mark all students as Absent
* 💾 Save attendance
* 📂 View previous attendance records
* 📥 Download attendance reports
* 📤 Share attendance reports

### CR cannot:

* ❌ Switch sections
* ❌ Add students
* ❌ Delete students
* ❌ Add subjects
* ❌ Delete subjects
* ❌ Assign CRs

These restrictions are enforced not only through the user interface but also through **Firestore Security Rules**.

---

# 📊 Attendance Management

QAttend makes attendance management simple and efficient.

### Features include:

* 📅 Select attendance date
* 📚 Select subject
* 👨‍🎓 View complete student list
* 🟢 Mark students as Present
* 🔴 Mark students as Absent
* ⚡ Mark all Present
* ❌ Mark all Absent
* 🔍 Search students by Name or Q.ID
* 📊 Real-time attendance statistics
* 💾 Save attendance records
* 🔄 Load previously saved attendance

The system also prevents accidental incomplete attendance submission.

> ⚠️ Attendance cannot be saved until all students have been marked.

---

# 📁 Attendance Records

All attendance records are stored securely in **Cloud Firestore**.

Each attendance record contains:

* 🏫 University
* 🎓 Course
* 📌 Section
* 📚 Subject
* 📅 Date
* 👨‍🎓 Student attendance status
* 👤 Marked by
* 🕒 Last updated time

Users can:

* 👀 View attendance records
* 📤 Share attendance reports
* 📥 Download attendance reports
* 🗑️ Delete records

---

# 📥 Excel Report Generation

QAttend Portal can generate attendance reports in **Excel format**.

Each report contains:

| Field        | Description           |
| ------------ | --------------------- |
| University   | Quantum University    |
| Course       | B.Tech                |
| Section      | Student Section       |
| Semester     | Semester 5            |
| Subject      | Selected Subject      |
| Date         | Attendance Date       |
| Q.ID         | Student University ID |
| Student Name | Name of Student       |
| Attendance   | Present / Absent      |

Attendance data can also be shared directly using the device's native sharing functionality.

---

# 🛠️ Technology Stack

### Frontend

* HTML5
* CSS3
* JavaScript

### Backend & Database

* Firebase Authentication
* Cloud Firestore

### Hosting

* Firebase Hosting

### Additional Libraries

* SheetJS (XLSX) for Excel report generation

---

# 🏗️ Project Structure

```text
qattend-portal/
│
├── index.html
├── style.css
├── script.js
│
├── logo.png
│
├── firestore.rules
├── firebase.json
├── .firebaserc
│
├── seed.html
├── seed.js
├── seed-backup.js
│
├── SETUP.md
│
├── .github/
│   └── workflows/
│       ├── firebase-hosting-merge.yml
│       └── firebase-hosting-pull-request.yml
│
└── .gitignore
```

---

# 🔐 Role-Based Access Control

QAttend Portal uses Firestore Security Rules to protect application data.

### Admin Access

```text
Students     → Read / Write
Subjects     → Read / Write
Attendance   → Read / Write
Users        → Read / Write
```

### Class Representative Access

```text
Students     → Read Only (Assigned Section)
Subjects     → Read Only (Assigned Section)
Attendance   → Read / Write (Assigned Section)
Users        → Own Profile Only
```

This means security is enforced at the database level and not just through hidden UI elements.

---

# 🚀 Getting Started

## 1️⃣ Clone the Repository

```bash
git clone https://github.com/YOUR_USERNAME/qattend-portal.git
```

Navigate to the project directory:

```bash
cd qattend-portal
```

---

## 2️⃣ Create a Firebase Project

Go to the Firebase Console and create a new project.

Enable:

* Firebase Authentication
* Email/Password Authentication
* Cloud Firestore
* Firebase Hosting

---

## 3️⃣ Configure Firebase

Open `script.js` and update the Firebase configuration with your own Firebase project credentials.

```javascript
const firebaseConfig = {
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_PROJECT.firebaseapp.com",
  projectId: "YOUR_PROJECT_ID",
  storageBucket: "YOUR_PROJECT.appspot.com",
  messagingSenderId: "YOUR_SENDER_ID",
  appId: "YOUR_APP_ID"
};
```

---

## 4️⃣ Configure Firestore Security Rules

Open:

```text
firestore.rules
```

Copy the rules into:

```text
Firebase Console
→ Firestore Database
→ Rules
```

Then click:

```text
Publish
```

---

# 👨‍💼 Creating the First Admin

The first administrator must be created manually.

### Step 1

Go to:

```text
Firebase Console
→ Authentication
→ Users
→ Add User
```

Create an account using an email and password.

⚠️ Use lowercase email addresses consistently.

---

### Step 2

Go to:

```text
Firestore Database
→ Data
```

Create a collection:

```text
users
```

Create a document using the Admin's email as the document ID.

Example:

```text
users
└── admin@example.com
```

Add:

```text
role: "admin"
```

Now the user can log in as an administrator.

---

# 👨‍🏫 Assigning a Class Representative

### Step 1

Create the CR account in:

```text
Firebase Authentication
```

### Step 2

Log into QAttend as Admin.

### Step 3

Open the:

```text
Assign CR
```

panel.

### Step 4

Enter:

* CR Email
* Assigned Section

The CR will then only have access to their assigned section.

---

# 📦 Importing Student Data

QAttend includes a data import system.

Open:

```text
seed.html
```

While logged in as an Admin.

Click:

```text
Run Import
```

This imports the available student roster and subjects into Firestore.

The import data includes students across multiple sections.

> ⚠️ The seed/import files are intended for initial database setup.

---

# 🌐 Deployment

The project is configured for **Firebase Hosting**.

Install Firebase CLI:

```bash
npm install -g firebase-tools
```

Login to Firebase:

```bash
firebase login
```

Initialize Firebase if required:

```bash
firebase init
```

Deploy the project:

```bash
firebase deploy
```

---

# 🔄 GitHub Actions

The repository includes GitHub Actions workflows for Firebase Hosting.

```text
.github/workflows/
```

Available workflows include:

* Firebase Hosting deployment for Pull Requests
* Firebase Hosting deployment after merging changes

This helps automate deployment through GitHub.

---

# 🔒 Security

QAttend Portal implements database-level security using Firestore Rules.

The system ensures:

* 🔐 Only authenticated users can access protected data
* 👨‍💼 Admins can manage all sections
* 👨‍🏫 CRs can access only their assigned section
* 🚫 CRs cannot modify students or subjects
* 📁 Attendance data is restricted by section
* 👤 Users can access their own role profile

---

# 📱 User Workflow

```text
User
  │
  ▼
Login
  │
  ├───────────────┐
  ▼               ▼
Admin             Class Representative
  │                       │
  ▼                       ▼
Manage System         Assigned Section Only
  │                       │
  ├── Students            ├── Mark Attendance
  ├── Subjects            ├── Save Attendance
  ├── CR Management       ├── View Records
  └── Attendance          └── Download / Share
```

---

# 📸 Screenshots

You can add screenshots of the application here.

```markdown
![Login Page](screenshots/login.png)

![Dashboard](screenshots/dashboard.png)

![Attendance Panel](screenshots/attendance.png)

![Admin Panel](screenshots/admin.png)
```

---

# 🔮 Future Improvements

Some possible improvements for future versions:

* 📊 Attendance analytics dashboard
* 📈 Student attendance percentage
* ⚠️ Low attendance warnings
* 👨‍🎓 Individual student attendance history
* 📱 Improved mobile experience
* 🔔 Notifications and reminders
* 📧 Email reports
* 📊 Advanced analytics
* 📅 Monthly attendance reports
* 🔍 Advanced filtering
* 👨‍🏫 Teacher accounts
* 📷 QR-based attendance
* 🧠 AI-based attendance insights

---

# 🤝 Contributing

Contributions are welcome.

1. Fork the repository
2. Create a new branch

```bash
git checkout -b feature/YourFeature
```

3. Make your changes
4. Commit your changes

```bash
git commit -m "Add new feature"
```

5. Push the branch

```bash
git push origin feature/YourFeature
```

6. Open a Pull Request

---

# 👨‍💻 Developer

**Karan Kumar**

Computer Science Student | Java Developer | Software Developer

* 💻 Passionate about building real-world projects
* 📚 Currently improving programming and software development skills
* 🚀 Focused on creating practical applications

---

# ⭐ Support

If you found this project useful, consider giving the repository a ⭐.

It helps support the project and motivates further development.

---

<div align="center">

### Made with ❤️ for better attendance management

**QAttend Portal**

</div>
