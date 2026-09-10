import { initializeApp } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js";

import {
    getAuth,
    signInWithEmailAndPassword,
    signOut,
    onAuthStateChanged,
    sendPasswordResetEmail
} from "https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js";

import {
    getFirestore,
    collection,
    doc,
    addDoc,
    setDoc,
    getDoc,
    getDocs,
    deleteDoc,
    query,
    orderBy
} from "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";


// FIREBASE CONFIG

const firebaseConfig = {
    apiKey: "AIzaSyA-NTLzWwa9R94_jN1cZMrBGwcMPJVVU2w",
    authDomain: "quantumuniveristy.firebaseapp.com",
    projectId: "quantumuniveristy",
    storageBucket: "quantumuniveristy.firebasestorage.app",
    messagingSenderId: "1043452018213",
    appId: "1:1043452018213:web:698ee52884812e0ad209cc",
    measurementId: "G-HSC3Z9RCGR"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);


// LAZY-LOAD SHEETJS (Excel library)
// Only fetched the first time Save/Download/Share is actually used — not on
// every page load. This is a large library, and loading it upfront made the
// login page noticeably slower on weak connections for no benefit, since
// most visits to the page are just to log in and mark attendance.
let xlsxLoadPromise = null;

function ensureXLSXLoaded() {
    if (window.XLSX) return Promise.resolve();

    if (!xlsxLoadPromise) {
        xlsxLoadPromise = new Promise(function (resolve, reject) {
            const script = document.createElement("script");
            script.src = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";
            script.onload = resolve;
            script.onerror = function () {
                xlsxLoadPromise = null; // allow retrying on next attempt
                reject(new Error("Could not load the Excel library. Check your internet connection."));
            };
            document.head.appendChild(script);
        });
    }

    return xlsxLoadPromise;
}


// FIXED SECTION LIST — edit here if names change
const SECTIONS = [
    { id: "SECTION-1", label: "Section 1" },
    { id: "SECTION-2", label: "Section 2" },
    { id: "SECTION-3", label: "Section 3" },
    { id: "SECTION-4", label: "Section 4" },
    { id: "SECTION-5", label: "Section 5" },
    { id: "SECTION-6", label: "Section 6" },
    { id: "SECTION-7", label: "Section 7" },
    { id: "SECTION-8", label: "Section 8" },
    { id: "AIML-1", label: "AIML - 1" },
    { id: "AIML-2", label: "AIML - 2" },
    { id: "CSCQ", label: "CSCQ" },
    { id: "DATA-SCIENCE", label: "Data Science" },
    { id: "FULL-STACK-DEV", label: "Full Stack Development" },
    { id: "CLOUD-TECH-INFOSEC", label: "Cloud Technology & Information Security" }
];

function sectionLabelOf(id) {
    const match = SECTIONS.find(section => section.id === id);
    return match ? match.label : id;
}

// Compares two names word by word — first name first, then middle name(s),
// then last name — rather than treating the whole name as one long string.
// This is the same end result as a plain string compare for most names, but
// is explicit and predictable regardless of spacing quirks in the data.
function compareNames(nameA, nameB) {
    const aWords = nameA.trim().split(/\s+/);
    const bWords = nameB.trim().split(/\s+/);
    const wordCount = Math.max(aWords.length, bWords.length);

    for (let i = 0; i < wordCount; i++) {
        const aWord = aWords[i] || "";
        const bWord = bWords[i] || "";
        const comparison = aWord.localeCompare(bWord);
        if (comparison !== 0) return comparison;
    }

    return 0;
}


// STATE

let profile = null;          // { role: "admin" | "cr", section: string|null }
let activeSection = null;    // section id currently being viewed/managed
let students = [];           // [{ id, qid, name, status }]
let subjects = [];           // [{ id, name }]
let recordsCache = {};       // { recordId: full record data } — used so Share
                              // can build the file instantly without an extra
                              // Firestore read, which otherwise delays the
                              // navigator.share() call past the click and
                              // makes some browsers reject it (NotAllowedError).


// HTML ELEMENTS

const loginPage = document.getElementById("loginPage");
const appPage = document.getElementById("appPage");

const loginForm = document.getElementById("loginForm");
const emailInput = document.getElementById("email");
const passwordInput = document.getElementById("password");
const togglePassword = document.getElementById("togglePassword");
const loginMessage = document.getElementById("loginMessage");

const forgotPasswordLink = document.getElementById("forgotPasswordLink");
const resetPasswordBox = document.getElementById("resetPasswordBox");
const resetEmail = document.getElementById("resetEmail");
const sendResetBtn = document.getElementById("sendResetBtn");
const resetMessage = document.getElementById("resetMessage");
const backToLoginLink = document.getElementById("backToLoginLink");

const logoutBtn = document.getElementById("logoutBtn");
const loggedUser = document.getElementById("loggedUser");
const roleBadge = document.getElementById("roleBadge");

const welcomeTitle = document.getElementById("welcomeTitle");
const welcomeSubtitle = document.getElementById("welcomeSubtitle");
const sectionLabel = document.getElementById("sectionLabel");

const sectionSwitchBox = document.getElementById("sectionSwitchBox");
const sectionSelect = document.getElementById("sectionSelect");
const workArea = document.getElementById("workArea");
const noSectionMessage = document.getElementById("noSectionMessage");

const adminPanel = document.getElementById("adminPanel");
const adminOnlyEls = document.querySelectorAll(".admin-only");

const addStudentForm = document.getElementById("addStudentForm");
const newStudentQid = document.getElementById("newStudentQid");
const newStudentName = document.getElementById("newStudentName");

const addSubjectForm = document.getElementById("addSubjectForm");
const newSubjectName = document.getElementById("newSubjectName");
const subjectManageList = document.getElementById("subjectManageList");

const addCrForm = document.getElementById("addCrForm");
const newCrEmail = document.getElementById("newCrEmail");
const newCrSection = document.getElementById("newCrSection");
const crManageList = document.getElementById("crManageList");

const studentList = document.getElementById("studentList");
const subjectSelect = document.getElementById("subjectSelect");
const attendanceDate = document.getElementById("attendanceDate");
const searchStudent = document.getElementById("searchStudent");

const totalStudents = document.getElementById("totalStudents");
const headerTotalStudents = document.getElementById("headerTotalStudents");
const presentStudents = document.getElementById("presentStudents");
const absentStudents = document.getElementById("absentStudents");

const saveBtn = document.getElementById("saveBtn");
const downloadBtn = document.getElementById("downloadBtn");
const markAllPresentBtn = document.getElementById("markAllPresentBtn");
const markAllAbsentBtn = document.getElementById("markAllAbsentBtn");

const recordsList = document.getElementById("recordsList");
const refreshRecordsBtn = document.getElementById("refreshRecordsBtn");


// TODAY DATE
attendanceDate.value = new Date().toISOString().split("T")[0];


// PASSWORD SHOW / HIDE
togglePassword.addEventListener("click", function () {
    if (passwordInput.type === "password") {
        passwordInput.type = "text";
        togglePassword.textContent = "🙈";
    } else {
        passwordInput.type = "password";
        togglePassword.textContent = "👁";
    }
});


// LOGIN
loginForm.addEventListener("submit", async function (event) {
    event.preventDefault();

    const email = emailInput.value.trim();
    const password = passwordInput.value;

    loginMessage.textContent = "Logging in...";

    try {
        await signInWithEmailAndPassword(auth, email, password);
        loginMessage.textContent = "";
    } catch (error) {
        console.error(error);
        loginMessage.textContent = "Invalid email or password.";
    }
});


// FORGOT PASSWORD — show the reset box, hide the login form
forgotPasswordLink.addEventListener("click", function (event) {
    event.preventDefault();

    resetEmail.value = emailInput.value.trim(); // carry over whatever they'd typed
    resetMessage.textContent = "";
    resetMessage.classList.remove("login-message-success");

    loginForm.classList.add("hidden");
    forgotPasswordLink.parentElement.classList.add("hidden");
    loginMessage.classList.add("hidden");
    resetPasswordBox.classList.remove("hidden");
});

// Back to the normal login form
backToLoginLink.addEventListener("click", function (event) {
    event.preventDefault();

    resetPasswordBox.classList.add("hidden");
    loginForm.classList.remove("hidden");
    forgotPasswordLink.parentElement.classList.remove("hidden");
    loginMessage.classList.remove("hidden");
    resetMessage.textContent = "";
});

// SEND RESET LINK — Firebase emails the user a link; clicking it lets them
// set a brand new password on Firebase's own page. No OTP, no SMTP setup,
// no third-party service needed — this is a built-in Firebase Auth feature.
sendResetBtn.addEventListener("click", async function () {
    const email = resetEmail.value.trim();

    if (!email) {
        resetMessage.classList.remove("login-message-success");
        resetMessage.textContent = "Please enter your email.";
        return;
    }

    resetMessage.classList.remove("login-message-success");
    resetMessage.textContent = "Sending...";
    sendResetBtn.disabled = true;

    try {
        await sendPasswordResetEmail(auth, email);

        resetMessage.classList.add("login-message-success");
        resetMessage.textContent =
            "✅ Reset link sent! Check your inbox (and spam/promotions folder).";

    } catch (error) {
        console.error(error);
        resetMessage.classList.remove("login-message-success");

        // Firebase intentionally doesn't reveal whether an email exists,
        // for privacy — so most errors here are just bad email formatting.
        if (error.code === "auth/invalid-email") {
            resetMessage.textContent = "That doesn't look like a valid email.";
        } else if (error.code === "auth/too-many-requests") {
            resetMessage.textContent = "Too many attempts. Please try again in a while.";
        } else {
            resetMessage.textContent = "Could not send reset email. Try again.";
        }

    } finally {
        sendResetBtn.disabled = false;
    }
});


// AUTH STATE
onAuthStateChanged(auth, async function (user) {

    if (!user) {
        appPage.classList.add("hidden");
        loginPage.classList.remove("hidden");
        profile = null;

        // Always land back on the normal login form, not a leftover reset box
        resetPasswordBox.classList.add("hidden");
        loginForm.classList.remove("hidden");
        forgotPasswordLink.parentElement.classList.remove("hidden");
        loginMessage.classList.remove("hidden");

        return;
    }

    // Look up this user's role + section
    try {
        const profileSnap = await getDoc(doc(db, "users", user.email));

        if (!profileSnap.exists()) {
            alert(
                "Your account is not set up yet. Ask the Admin to assign you a role."
            );
            await signOut(auth);
            return;
        }

        profile = profileSnap.data();

    } catch (error) {
        console.error(error);
        alert("Could not load your account role. Try logging in again.");
        await signOut(auth);
        return;
    }

    loginPage.classList.add("hidden");
    appPage.classList.remove("hidden");

    loggedUser.textContent = user.email;
    roleBadge.textContent = profile.role === "admin" ? "ADMIN" : "CLASS REPRESENTATIVE";

    const isAdmin = profile.role === "admin";

    adminOnlyEls.forEach(function (el) {
        el.classList.toggle("hidden", !isAdmin);
    });

    if (isAdmin) {
        welcomeTitle.textContent = "Admin Dashboard";
        welcomeSubtitle.textContent = "Manage every section, subject and CR from one place.";

        populateSectionDropdowns();

        // No section chosen yet — don't load or show any section's data until Admin picks one
        activeSection = null;
        sectionSelect.value = "";
        workArea.classList.add("hidden");
        noSectionMessage.classList.remove("hidden");
        sectionLabel.textContent = "-- Not selected --";
        headerTotalStudents.textContent = "0";

        await loadCrList();
        await loadSubjects();

    } else {
        welcomeTitle.textContent = `${sectionLabelOf(profile.section)} Attendance Portal`;
        welcomeSubtitle.textContent = "Mark, save and download attendance for your section.";

        activeSection = profile.section;
        workArea.classList.remove("hidden");
        noSectionMessage.classList.add("hidden");
        sectionLabel.textContent = sectionLabelOf(activeSection);

        await loadSubjects();
        await loadStudents();
        await loadRecords();
    }
});


// LOGOUT
logoutBtn.addEventListener("click", async function () {
    await signOut(auth);
});


// SECTION DROPDOWNS (admin)
function populateSectionDropdowns() {
    sectionSelect.innerHTML = `<option value="">-- Select Section --</option>`;
    newCrSection.innerHTML = "";

    SECTIONS.forEach(function (section) {
        const option1 = document.createElement("option");
        option1.value = section.id;
        option1.textContent = section.label;
        sectionSelect.appendChild(option1);

        const option2 = document.createElement("option");
        option2.value = section.id;
        option2.textContent = section.label;
        newCrSection.appendChild(option2);
    });
}

sectionSelect.addEventListener("change", async function () {
    activeSection = sectionSelect.value || null;

    if (!activeSection) {
        workArea.classList.add("hidden");
        noSectionMessage.classList.remove("hidden");
        sectionLabel.textContent = "-- Not selected --";
        headerTotalStudents.textContent = "0";
        totalStudents.textContent = "0";
        presentStudents.textContent = "0";
        absentStudents.textContent = "0";
        studentList.innerHTML = "";
        students = [];
        return;
    }

    workArea.classList.remove("hidden");
    noSectionMessage.classList.add("hidden");
    sectionLabel.textContent = sectionLabelOf(activeSection);

    // Different section = different attendance context, so reset selection
    subjectSelect.value = "";
    attendanceDate.value = new Date().toISOString().split("T")[0];

    await loadStudents();
    await loadSubjects();
    await loadRecords();
});


// STUDENTS — LOAD

async function loadStudents() {
    try {
        const snapshot = await getDocs(
            collection(db, "sections", activeSection, "students")
        );

        students = snapshot.docs
            .map(function (docSnap) {
                return {
                    id: docSnap.id,
                    qid: docSnap.data().qid,
                    name: docSnap.data().name,
                    status: null
                };
            })
            .sort(function (a, b) {
                // Sort by first name, then middle name(s), then last name.
                return compareNames(a.name, b.name);
            });

        displayStudents();

    } catch (error) {
        console.error(error);
        studentList.innerHTML = `<p class="empty-record">Error loading students.</p>`;
    }
}


// STUDENTS — DISPLAY

function displayStudents() {
    const search = searchStudent.value.trim().toLowerCase();
    const isAdmin = profile.role === "admin";

    studentList.innerHTML = "";

    students.forEach(function (student) {

        const matchesName = student.name.toLowerCase().includes(search);
        const matchesQid = student.qid.toLowerCase().includes(search);

        if (!matchesName && !matchesQid) return;

        const row = document.createElement("div");
        row.className = "student";
        row.dataset.id = student.id;

        row.innerHTML = `
            <div class="qid-number">${student.qid}</div>
            <div class="student-name">${student.name}</div>
            <div class="buttons">
                <button class="absent ${student.status === "Absent" ? "active-absent" : ""}"
                    data-id="${student.id}" data-status="Absent">✕ Absent</button>
                <button class="present ${student.status === "Present" ? "active-present" : ""}"
                    data-id="${student.id}" data-status="Present">✓ Present</button>
            </div>
            ${isAdmin ? `<button class="delete-student-btn" data-id="${student.id}" title="Delete student">🗑</button>` : ""}
        `;

        studentList.appendChild(row);
    });

    updateStats();
}

// Handle Present/Absent/Delete clicks via a single delegated listener
// attached once to the list container, instead of re-attaching listeners
// to every button on every render. On Present/Absent, only that row's two
// buttons are updated directly — the whole list is not re-rendered — which
// keeps marking attendance fast even on sections with 60-80+ students.
studentList.addEventListener("click", function (event) {
    const statusButton = event.target.closest("[data-status]");
    if (statusButton) {
        if (!subjectSelect.value) {
            alert("Please select a subject first.");
            return;
        }

        const student = students.find(item => item.id === statusButton.dataset.id);
        if (!student) return;

        student.status = statusButton.dataset.status;

        const row = statusButton.closest(".student");
        const absentBtn = row.querySelector(".absent");
        const presentBtn = row.querySelector(".present");
        absentBtn.classList.toggle("active-absent", student.status === "Absent");
        presentBtn.classList.toggle("active-present", student.status === "Present");

        updateStats();
        return;
    }

    const deleteButton = event.target.closest(".delete-student-btn");
    if (deleteButton) {
        deleteStudent(deleteButton.dataset.id);
    }
});


// STUDENTS — ADD (admin only, enforced by Firestore rules too)

addStudentForm.addEventListener("submit", async function (event) {
    event.preventDefault();

    if (!activeSection) {
        alert("Pick a section first (dropdown above).");
        return;
    }

    const qid = newStudentQid.value.trim();
    const name = newStudentName.value.trim().toUpperCase();

    if (!qid || !name) return;

    try {
        await addDoc(collection(db, "sections", activeSection, "students"), { qid, name });

        newStudentQid.value = "";
        newStudentName.value = "";

        await loadStudents();

    } catch (error) {
        console.error(error);
        alert("Error adding student.");
    }
});


// STUDENTS — DELETE (admin only)

async function deleteStudent(studentId) {
    const student = students.find(item => item.id === studentId);
    const confirmDelete = confirm(`Delete ${student ? student.name : "this student"}?`);
    if (!confirmDelete) return;

    try {
        await deleteDoc(doc(db, "sections", activeSection, "students", studentId));
        await loadStudents();
    } catch (error) {
        console.error(error);
        alert("Error deleting student.");
    }
}


// UPDATE STATS

function updateStats() {
    const total = students.length;
    const present = students.filter(student => student.status === "Present").length;
    const absent = students.filter(student => student.status === "Absent").length;

    totalStudents.textContent = total;
    headerTotalStudents.textContent = total;
    presentStudents.textContent = present;
    absentStudents.textContent = absent;
}


// SEARCH
searchStudent.addEventListener("input", displayStudents);


// MARK ALL PRESENT / ABSENT

markAllPresentBtn.addEventListener("click", function () {
    if (!subjectSelect.value) {
        alert("Please select a subject first.");
        return;
    }
    students.forEach(student => student.status = "Present");
    displayStudents();
});

markAllAbsentBtn.addEventListener("click", function () {
    if (!subjectSelect.value) {
        alert("Please select a subject first.");
        return;
    }
    students.forEach(student => student.status = "Absent");
    displayStudents();
});


// SUBJECTS — LOAD (specific to the currently active section)

async function loadSubjects() {
    if (!activeSection) return;

    try {
        const q = query(collection(db, "sections", activeSection, "subjects"), orderBy("name"));
        const snapshot = await getDocs(q);

        subjects = snapshot.docs.map(docSnap => ({ id: docSnap.id, name: docSnap.data().name }));

        subjectSelect.innerHTML = `<option value="">-- Select Subject --</option>`;
        subjects.forEach(function (subject) {
            const option = document.createElement("option");
            option.value = subject.name;
            option.textContent = subject.name;
            subjectSelect.appendChild(option);
        });

        if (profile.role === "admin") {
            subjectManageList.innerHTML = "";
            subjects.forEach(function (subject) {
                const row = document.createElement("div");
                row.className = "chip-row";
                row.innerHTML = `
                    <span>${subject.name}</span>
                    <button data-id="${subject.id}">Delete</button>
                `;
                subjectManageList.appendChild(row);
            });

            document.querySelectorAll("#subjectManageList button").forEach(function (button) {
                button.addEventListener("click", function () {
                    deleteSubject(button.dataset.id);
                });
            });
        }

    } catch (error) {
        console.error(error);
    }
}


// SUBJECTS — ADD (admin only, added to the currently active section only)

addSubjectForm.addEventListener("submit", async function (event) {
    event.preventDefault();

    if (!activeSection) {
        alert("Pick a section first (dropdown above).");
        return;
    }

    const name = newSubjectName.value.trim();
    if (!name) return;

    try {
        await addDoc(collection(db, "sections", activeSection, "subjects"), { name });
        newSubjectName.value = "";
        await loadSubjects();
    } catch (error) {
        console.error(error);
        alert("Error adding subject.");
    }
});


// SUBJECTS — DELETE (admin only, only from the currently active section)

async function deleteSubject(subjectId) {
    const confirmDelete = confirm("Delete this subject from this section? Existing saved records will keep the old name.");
    if (!confirmDelete) return;

    try {
        await deleteDoc(doc(db, "sections", activeSection, "subjects", subjectId));
        await loadSubjects();
    } catch (error) {
        console.error(error);
        alert("Error deleting subject.");
    }
}


// CR MANAGEMENT — LOAD (admin only)

async function loadCrList() {
    try {
        const snapshot = await getDocs(collection(db, "users"));

        crManageList.innerHTML = "";

        snapshot.docs
            .filter(docSnap => docSnap.data().role === "cr")
            .forEach(function (docSnap) {
                const data = docSnap.data();
                const row = document.createElement("div");
                row.className = "chip-row";
                row.innerHTML = `
                    <span>${docSnap.id} — ${sectionLabelOf(data.section)}</span>
                    <button data-email="${docSnap.id}">Remove</button>
                `;
                crManageList.appendChild(row);
            });

        document.querySelectorAll("#crManageList button").forEach(function (button) {
            button.addEventListener("click", function () {
                deleteCr(button.dataset.email);
            });
        });

    } catch (error) {
        console.error(error);
    }
}


// CR MANAGEMENT — ASSIGN (admin only)

addCrForm.addEventListener("submit", async function (event) {
    event.preventDefault();

    const email = newCrEmail.value.trim().toLowerCase();
    const section = newCrSection.value;

    if (!email || !section) return;

    try {
        await setDoc(doc(db, "users", email), {
            role: "cr",
            section: section,
            createdAt: new Date().toISOString()
        });

        newCrEmail.value = "";
        await loadCrList();

        alert(
            `${email} is assigned as CR for ${sectionLabelOf(section)}.\n` +
            `Make sure their Firebase Auth login uses this exact (lowercase) email.`
        );

    } catch (error) {
        console.error(error);
        alert("Error assigning CR. Make sure the login was created with the exact same lowercase email.");
    }
});


// CR MANAGEMENT — REMOVE (admin only)
// Note: this removes app access only. It does NOT delete their Firebase Auth login —
// do that separately in the Firebase Console if you want to fully revoke access.

async function deleteCr(email) {
    const confirmDelete = confirm(`Remove CR access for ${email}? (Their login will still exist in Firebase Console.)`);
    if (!confirmDelete) return;

    try {
        await deleteDoc(doc(db, "users", email));
        await loadCrList();
    } catch (error) {
        console.error(error);
        alert("Error removing CR.");
    }
}


// GET ATTENDANCE RECORD ID

function getAttendanceId() {
    const date = attendanceDate.value;
    const subject = subjectSelect.value;
    const cleanSubject = subject.replace(/[^a-zA-Z0-9]/g, "_");
    return `${date}_${cleanSubject}`;
}


// LOAD ATTENDANCE (for the selected date + subject)

async function loadAttendance() {
    const date = attendanceDate.value;
    const subject = subjectSelect.value;
    if (!date || !subject) return;

    try {
        const attendanceRef = doc(db, "sections", activeSection, "attendance", getAttendanceId());
        const snapshot = await getDoc(attendanceRef);

        if (snapshot.exists()) {
            const data = snapshot.data();

            students.forEach(function (student) {
                const saved = data.students.find(item => item.qid === student.qid);
                student.status = saved ? saved.status : null;
            });
        } else {
            students.forEach(student => student.status = null);
        }

        displayStudents();

    } catch (error) {
        console.error(error);
        alert("Error loading attendance.");
    }
}

subjectSelect.addEventListener("change", function () {
    if (!attendanceDate.value) {
        alert("Please select date first.");
        subjectSelect.value = "";
        return;
    }
    loadAttendance();
});

attendanceDate.addEventListener("change", function () {
    if (subjectSelect.value) loadAttendance();
});


// SAVE ATTENDANCE

saveBtn.addEventListener("click", async function () {
    const user = auth.currentUser;
    if (!user) return;

    const date = attendanceDate.value;
    const subject = subjectSelect.value;

    if (!date) {
        alert("Please select date first.");
        return;
    }
    if (!subject) {
        alert("Please select subject.");
        return;
    }

    const unmarked = students.filter(student => student.status === null);
    if (unmarked.length > 0) {
        alert(`${unmarked.length} students are not marked.`);
        return;
    }

    try {
        saveBtn.textContent = "Saving...";

        const attendanceRef = doc(db, "sections", activeSection, "attendance", getAttendanceId());

        await setDoc(attendanceRef, {
            university: "Quantum University",
            course: "B.Tech",
            section: sectionLabelOf(activeSection),
            subject: subject,
            date: date,
            students: students,
            markedBy: user.email,
            updatedAt: new Date().toISOString()
        });

        alert("✅ Attendance saved successfully!");
        await loadRecords();

    } catch (error) {
        console.error(error);
        alert("❌ Error saving attendance.");
    } finally {
        saveBtn.textContent = "💾 Save Attendance";
    }
});


// LOAD RECORDS

async function loadRecords() {
    recordsList.innerHTML = `<p class="empty-record">Loading records...</p>`;

    try {
        const attendanceCollection = collection(db, "sections", activeSection, "attendance");
        const q = query(attendanceCollection, orderBy("updatedAt", "desc"));
        const snapshot = await getDocs(q);

        recordsList.innerHTML = "";

        if (snapshot.empty) {
            recordsList.innerHTML = `<p class="empty-record">No previous attendance records found.</p>`;
            return;
        }

        snapshot.forEach(function (documentSnapshot) {
            const data = documentSnapshot.data();
            recordsCache[documentSnapshot.id] = data;

            const present = data.students.filter(student => student.status === "Present").length;
            const absent = data.students.filter(student => student.status === "Absent").length;

            const card = document.createElement("div");
            card.className = "record-card";

            card.innerHTML = `
                <div>
                    <h3>${data.subject}</h3>
                    <p>📅 ${data.date}</p>
                    <p>🟢 Present: ${present} | 🔴 Absent: ${absent}</p>
                </div>
                <div class="record-actions">
                    <button class="view-record-btn" data-id="${documentSnapshot.id}">View</button>
                    <button class="share-record-btn" data-id="${documentSnapshot.id}">Share</button>
                    <button class="delete-record-btn" data-id="${documentSnapshot.id}">Delete</button>
                </div>
            `;

            recordsList.appendChild(card);
        });

        addRecordButtonEvents();

    } catch (error) {
        console.error(error);
        recordsList.innerHTML = `<p class="empty-record">Error loading records.</p>`;
    }
}


// RECORD EVENTS

function addRecordButtonEvents() {
    document.querySelectorAll(".view-record-btn").forEach(function (button) {
        button.addEventListener("click", async function () {
            await viewRecord(button.dataset.id);
        });
    });

    document.querySelectorAll(".share-record-btn").forEach(function (button) {
        button.addEventListener("click", async function () {
            await shareRecord(button.dataset.id);
        });
    });

    document.querySelectorAll(".delete-record-btn").forEach(function (button) {
        button.addEventListener("click", async function () {
            await deleteRecord(button.dataset.id);
        });
    });
}


// SHARE RECORD — builds the XLSX for one saved record and opens the
// device's native share sheet (WhatsApp, Email, etc.) so Admin/CR can send
// it directly, without first downloading and then attaching it manually.

async function shareRecord(recordId) {
    try {
        const data = recordsCache[recordId];

        if (!data) {
            alert("Record data not loaded — try refreshing the records list first.");
            return;
        }

        await ensureXLSXLoaded();

        // Fixed serial = each student's position in the permanent alphabetical
        // roster (data.students, saved in that order at save time) — this
        // stays the same across every attendance record, like a roll number.
        const serialByQid = new Map();
        data.students.forEach(function (student, index) {
            serialByQid.set(student.qid, index + 1);
        });

        const sortedStudents = [...data.students].sort(function (a, b) {
            if (a.status === "Present" && b.status === "Absent") return -1;
            if (a.status === "Absent" && b.status === "Present") return 1;
            return 0;
        });

        const excelData = [
            ["QATTEND ATTENDANCE REPORT"],
            [],
            ["University", data.university || "Quantum University"],
            ["Course", data.course || "B.Tech"],
            ["Section", data.section],
            ["Semester", "5"],
            ["Subject", data.subject],
            ["Date", data.date],
            [],
            ["S.No", "Q.ID", "Name of Student", "Attendance"]
        ];

        const presentSerials = [];

        sortedStudents.forEach(function (student) {
            const serial = serialByQid.get(student.qid);
            excelData.push([serial, student.qid, student.name, student.status === "Present" ? 1 : 0]);
            if (student.status === "Present") presentSerials.push(serial);
        });

        presentSerials.sort((a, b) => a - b);
        excelData.push([]);
        excelData.push(["Present (S.No, for shortcut)", presentSerials.join(" ")]);

        const worksheet = XLSX.utils.aoa_to_sheet(excelData);
        worksheet["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 50 }, { wch: 15 }];

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance");

        const cleanSubject = data.subject.replace(/[^a-zA-Z0-9]/g, "_");
        const fileName = `QAttend_${data.section}_${cleanSubject}_${data.date}.xlsx`;

        const wbArray = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbArray], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });
        const file = new File([blob], fileName, { type: blob.type });

        if (navigator.canShare && navigator.canShare({ files: [file] })) {
            try {
                await navigator.share({
                    files: [file],
                    title: `${data.subject} Attendance`,
                    text: `Attendance for ${data.subject} on ${data.date} — ${data.section}`
                });
                return;
            } catch (shareError) {
                if (shareError.name === "AbortError") return; // user closed the share sheet
                // Some desktop browsers report support (canShare: true) but then
                // refuse the actual share call — fall through to download instead
                // of showing an error, since the file is still perfectly usable.
                console.error(shareError);
            }
        }

        XLSX.writeFile(workbook, fileName);
        alert("Sharing isn't available in this browser — the file was downloaded instead. You can share it manually.");

    } catch (error) {
        console.error(error);
        alert(`Error sharing record: ${error.name || ""} — ${error.message || error}`);
    }
}


// VIEW RECORD

async function viewRecord(recordId) {
    try {
        const recordRef = doc(db, "sections", activeSection, "attendance", recordId);
        const snapshot = await getDoc(recordRef);

        if (!snapshot.exists()) {
            alert("Record not found.");
            return;
        }

        const data = snapshot.data();

        attendanceDate.value = data.date;
        subjectSelect.value = data.subject;

        students.forEach(function (student) {
            const saved = data.students.find(item => item.qid === student.qid);
            student.status = saved ? saved.status : null;
        });

        displayStudents();
        window.scrollTo({ top: 0, behavior: "smooth" });

    } catch (error) {
        console.error(error);
        alert("Error loading record.");
    }
}


// DELETE RECORD

async function deleteRecord(recordId) {
    const confirmDelete = confirm("Are you sure you want to delete this attendance record?");
    if (!confirmDelete) return;

    try {
        await deleteDoc(doc(db, "sections", activeSection, "attendance", recordId));
        alert("Attendance record deleted.");
        await loadRecords();
    } catch (error) {
        console.error(error);
        alert("Error deleting record.");
    }
}


// REFRESH
refreshRecordsBtn.addEventListener("click", loadRecords);


// DOWNLOAD XLSX

downloadBtn.addEventListener("click", async function () {
    const date = attendanceDate.value;
    const subject = subjectSelect.value;

    if (!date || !subject) {
        alert("Select date and subject first.");
        return;
    }

    const unmarked = students.filter(student => student.status === null);
    if (unmarked.length > 0) {
        alert("Please mark all students before downloading.");
        return;
    }

    try {
        await ensureXLSXLoaded();
    } catch (error) {
        alert(error.message);
        return;
    }

    // Fixed serial = each student's position in the current alphabetical
    // roster (students, already alphabetically sorted by loadStudents) —
    // stays the same across records as long as the roster doesn't change.
    const serialByQid = new Map();
    students.forEach(function (student, index) {
        serialByQid.set(student.qid, index + 1);
    });

    const sortedStudents = [...students].sort(function (a, b) {
        if (a.status === "Present" && b.status === "Absent") return -1;
        if (a.status === "Absent" && b.status === "Present") return 1;
        return 0;
    });

    const excelData = [
        ["QATTEND ATTENDANCE REPORT"],
        [],
        ["University", "Quantum University"],
        ["Course", "B.Tech"],
        ["Section", sectionLabelOf(activeSection)],
        ["Semester", "5"],
        ["Subject", subject],
        ["Date", date],
        [],
        ["S.No", "Q.ID", "Name of Student", "Attendance"]
    ];

    const presentSerials = [];

    sortedStudents.forEach(function (student) {
        const serial = serialByQid.get(student.qid);
        excelData.push([serial, student.qid, student.name, student.status === "Present" ? 1 : 0]);
        if (student.status === "Present") presentSerials.push(serial);
    });

    presentSerials.sort((a, b) => a - b);
    excelData.push([]);
    excelData.push(["Present (S.No, for shortcut)", presentSerials.join(" ")]);

    const worksheet = XLSX.utils.aoa_to_sheet(excelData);
    worksheet["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 50 }, { wch: 15 }];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance");

    const cleanSubject = subject.replace(/[^a-zA-Z0-9]/g, "_");
    XLSX.writeFile(workbook, `QAttend_${sectionLabelOf(activeSection)}_${cleanSubject}_${date}.xlsx`);
});