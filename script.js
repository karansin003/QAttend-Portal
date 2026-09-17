import { initializeApp } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js";

import {
    getAuth,
    signInWithEmailAndPassword,
    signOut,
    onAuthStateChanged,
    setPersistence,
    browserLocalPersistence,
    sendPasswordResetEmail,
    createUserWithEmailAndPassword
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
    updateDoc,
    query,
    orderBy,
    limit,
    where,
    writeBatch
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
const authPersistenceReady = setPersistence(auth, browserLocalPersistence).catch(error => {
    console.warn("Could not enable local login persistence:", error);
});

// A secondary Firebase Auth instance is used only when Admin approves a
// new-section request. It creates the CR account without signing the Admin out.
const provisioningApp = initializeApp(firebaseConfig, "qattend-cr-provisioning");
let provisioningAuth = null;

function getProvisioningAuth() {
    if (!provisioningAuth) provisioningAuth = getAuth(provisioningApp);
    return provisioningAuth;
}

// DIRECT EMAIL CONFIGURATION
// Replace these three EmailJS values with your EmailJS account values.
// Emails are sent directly. No Gmail/Apple Mail compose window is opened.
const EMAILJS_CONFIG = {
    publicKey: "SADNPvxpnkAk3gjv_",
    serviceId: "service_ju470aa",
    templateId: "template_9uvbayk",
    adminEmail: "sonusin8672@gmail.com"
};

let emailJsLoadPromise = null;

function ensureEmailJSLoaded() {
    if (window.emailjs) return Promise.resolve(window.emailjs);
    if (!emailJsLoadPromise) {
        emailJsLoadPromise = new Promise((resolve, reject) => {
            const script = document.createElement("script");
            script.src = "https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js";
            script.onload = () => {
                window.emailjs.init({ publicKey: EMAILJS_CONFIG.publicKey });
                resolve(window.emailjs);
            };
            script.onerror = () => {
                emailJsLoadPromise = null;
                reject(new Error("Email service could not be loaded. Check your internet connection."));
            };
            document.head.appendChild(script);
        });
    }
    return emailJsLoadPromise;
}

async function sendDirectEmail({toEmail, subject, message, replyTo=""}) {
    const emailjs = await ensureEmailJSLoaded();
    if (EMAILJS_CONFIG.publicKey.startsWith("YOUR_") || EMAILJS_CONFIG.serviceId.startsWith("YOUR_") || EMAILJS_CONFIG.templateId.startsWith("YOUR_")) {
        throw new Error("EmailJS is not configured yet. Add your EmailJS Public Key, Service ID and Template ID in script.js.");
    }
    return emailjs.send(EMAILJS_CONFIG.serviceId, EMAILJS_CONFIG.templateId, {
        to_email: toEmail,
        subject,
        message,
        from_name: "QAttend",
        reply_to: replyTo || toEmail
    });
}



// LAZY-LOAD SHEETJS

let xlsxLoadPromise = null;

function ensureXLSXLoaded() {

    if (window.XLSX) {
        return Promise.resolve();
    }

    if (!xlsxLoadPromise) {

        xlsxLoadPromise = new Promise(
            function (resolve, reject) {

                const script =
                    document.createElement("script");

                script.src =
                    "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";

                script.onload =
                    resolve;

                script.onerror =
                    function () {

                        xlsxLoadPromise = null;

                        reject(
                            new Error(
                                "Could not load the Excel library. Check your internet connection."
                            )
                        );
                    };

                document.head.appendChild(script);
            }
        );
    }

    return xlsxLoadPromise;
}

// EXCEL TEMPLATE GENERATORS & PARSERS
async function downloadSectionTemplate(e) {
    if (e && typeof e.preventDefault === "function") e.preventDefault();
    try {
        await ensureXLSXLoaded();
        const wb = XLSX.utils.book_new();

        // Sheet 1: Students
        const studentData = [
            ["QID", "Student Name"],
            ["24030101", "Karan Kumar"],
            ["24030102", "Rahul Kumar"],
            ["24030103", "Amit Singh"],
            ["24030104", "Neha Sharma"],
            ["24030105", "Priya Verma"]
        ];
        const wsStudents = XLSX.utils.aoa_to_sheet(studentData);
        XLSX.utils.book_append_sheet(wb, wsStudents, "Students");

        // Sheet 2: Subjects
        const subjectData = [
            ["Subject Name"],
            ["Design and Analysis of Algorithm"],
            ["Database Management System"],
            ["Operating System"],
            ["Computer Networks"],
            ["Machine Learning"],
            ["Artificial Intelligence"]
        ];
        const wsSubjects = XLSX.utils.aoa_to_sheet(subjectData);
        XLSX.utils.book_append_sheet(wb, wsSubjects, "Subjects");

        // Sheet 3: Instructions
        const instructions = [
            ["Sheet / Column", "Rule & Guideline"],
            ["Students: QID", "Fill student university QID (unique, required)"],
            ["Students: Student Name", "Fill student full name (required)"],
            ["Subjects: Subject Name", "Fill each subject on a new row (required)"],
            ["Column Headers", "Do not modify column headers in row 1"],
            ["Duplicates", "Do not add duplicate QIDs"],
            ["File Format", "Save as .xlsx and upload in QAttend"]
        ];
        const wsInstructions = XLSX.utils.aoa_to_sheet(instructions);
        XLSX.utils.book_append_sheet(wb, wsInstructions, "Instructions");

        XLSX.writeFile(wb, "Section_Template.xlsx");
    } catch (err) {
        console.error("Error generating section template:", err);
        window.location.href = "Section_Template.xlsx";
    }
}

async function downloadStudentTemplate(e) {
    if (e && typeof e.preventDefault === "function") e.preventDefault();
    try {
        await ensureXLSXLoaded();
        const wb = XLSX.utils.book_new();
        const studentData = [
            ["QID", "Student Name"],
            ["24030101", "Karan Kumar"],
            ["24030102", "Rahul Kumar"]
        ];
        const wsStudents = XLSX.utils.aoa_to_sheet(studentData);
        XLSX.utils.book_append_sheet(wb, wsStudents, "Students");

        const instructions = [
            ["Column", "Rule"],
            ["QID", "University QID (Required, Unique)"],
            ["Student Name", "Student Full Name (Required)"]
        ];
        const wsInstructions = XLSX.utils.aoa_to_sheet(instructions);
        XLSX.utils.book_append_sheet(wb, wsInstructions, "Instructions");

        XLSX.writeFile(wb, "Student_Template.xlsx");
    } catch (err) {
        console.error("Error generating student template:", err);
    }
}

async function downloadSubjectTemplate(e) {
    if (e && typeof e.preventDefault === "function") e.preventDefault();
    try {
        await ensureXLSXLoaded();
        const wb = XLSX.utils.book_new();
        const subjectData = [
            ["Subject Name"],
            ["Design and Analysis of Algorithm"],
            ["Database Management System"],
            ["Operating System"]
        ];
        const wsSubjects = XLSX.utils.aoa_to_sheet(subjectData);
        XLSX.utils.book_append_sheet(wb, wsSubjects, "Subjects");

        const instructions = [
            ["Column", "Rule"],
            ["Subject Name", "Course Subject Name (Required, Unique)"]
        ];
        const wsInstructions = XLSX.utils.aoa_to_sheet(instructions);
        XLSX.utils.book_append_sheet(wb, wsInstructions, "Instructions");

        XLSX.writeFile(wb, "Subject_Template.xlsx");
    } catch (err) {
        console.error("Error generating subject template:", err);
    }
}

document.querySelectorAll('a[href="Section_Template.xlsx"]').forEach(link => {
    link.addEventListener("click", downloadSectionTemplate);
});

async function parseSectionExcelFile(file) {
    await ensureXLSXLoaded();

    if (!file) {
        throw new Error("Please select an Excel file.");
    }

    const allowed = /\.xlsx$/i.test(file.name);
    if (!allowed) {
        throw new Error("Please upload an .xlsx Excel file.");
    }

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });

    const normalize = value => String(value ?? "").replace(/\u00a0/g, " ").trim();

    const findSheet = name => {
        const target = name.replace(/\s+/g, "").toLowerCase();
        const sheetName = workbook.SheetNames.find(candidate =>
            candidate.replace(/\s+/g, "").toLowerCase() === target
        );
        return sheetName ? workbook.Sheets[sheetName] : null;
    };

    const studentsSheet =
        findSheet("Students") ||
        workbook.Sheets[workbook.SheetNames[0]];

    const subjectsSheet =
        findSheet("Subjects") ||
        (workbook.SheetNames.length > 1 ? workbook.Sheets[workbook.SheetNames[1]] : null);

    const studentRows = studentsSheet
        ? XLSX.utils.sheet_to_json(studentsSheet, { header: 1, defval: "" })
        : [];

    const subjectRows = subjectsSheet
        ? XLSX.utils.sheet_to_json(subjectsSheet, { header: 1, defval: "" })
        : [];

    const studentsOut = [];
    const qids = new Set();
    let duplicateCount = 0;
    let invalidCount = 0;

    studentRows.slice(1).forEach(row => {
        const qid = normalize(row[0]);
        const name = normalize(row[1]);

        if (!qid && !name) return;

        if (!qid || !name) {
            invalidCount++;
            return;
        }

        const qidKey = qid.toLowerCase();
        if (qids.has(qidKey)) {
            duplicateCount++;
            return;
        }

        qids.add(qidKey);
        studentsOut.push({
            qid,
            name: name.toUpperCase()
        });
    });

    const subjectsOut = [];
    const subjectSet = new Set();

    // Preferred format: Subjects sheet, one subject per row.
    subjectRows.slice(1).forEach(row => {
        const name = normalize(row[0]);
        if (!name) return;

        const key = name.toLowerCase();
        if (!subjectSet.has(key)) {
            subjectSet.add(key);
            subjectsOut.push(name);
        }
    });

    // Backward compatibility for the old one-sheet template:
    // Qid | Student Name | Subject1 | Subject2 | Subject3 ...
    if (subjectsOut.length === 0 && workbook.SheetNames.length === 1) {
        const headers = studentRows[0] || [];
        headers.slice(2).forEach(header => {
            const subject = normalize(header).replace(/^subject\s*\d+$/i, "");
            if (!subject) return;
            const key = subject.toLowerCase();
            if (!subjectSet.has(key)) {
                subjectSet.add(key);
                subjectsOut.push(subject);
            }
        });

        studentRows.slice(1).forEach(row => {
            row.slice(2).forEach(value => {
                const subject = normalize(value);
                if (!subject) return;
                const key = subject.toLowerCase();
                if (!subjectSet.has(key)) {
                    subjectSet.add(key);
                    subjectsOut.push(subject);
                }
            });
        });
    }

    if (studentsOut.length === 0 && subjectsOut.length === 0) {
        throw new Error("The Excel file does not contain any valid students or subjects.");
    }

    return {
        students: studentsOut,
        subjects: subjectsOut,
        duplicateCount,
        invalidCount,
        fileName: file.name
    };
}

function renderSectionExcelPreview(prefix, parsedData) {
    const previewEl = document.getElementById(prefix ? `${prefix}ExcelPreview` : "sectionExcelPreview");
    if (!previewEl) return;

    const badge = document.getElementById(prefix ? `${prefix}SuccessBadge` : "sectionExcelSuccessBadge");
    const fileNameEl = document.getElementById(prefix ? `${prefix}FileName` : "sectionExcelFileName");
    if (badge && fileNameEl) {
        fileNameEl.textContent = parsedData.fileName || "section_data.xlsx";
        badge.classList.remove("hidden");
    }

    const studentCountEl = document.getElementById(prefix ? `${prefix}StudentCount` : "excelStudentCount");
    const subjectCountEl = document.getElementById(prefix ? `${prefix}SubjectCount` : "excelSubjectCount");
    const duplicateCountEl = document.getElementById(prefix ? `${prefix}DuplicateCount` : "excelDuplicateCount");
    const invalidCountEl = document.getElementById(prefix ? `${prefix}InvalidCount` : "excelInvalidCount");

    if (studentCountEl) studentCountEl.textContent = parsedData.students.length;
    if (subjectCountEl) subjectCountEl.textContent = parsedData.subjects.length;
    if (duplicateCountEl) duplicateCountEl.textContent = parsedData.duplicateCount;
    if (invalidCountEl) invalidCountEl.textContent = parsedData.invalidCount;

    // Subjects Preview
    const subjectsBox = document.getElementById(prefix ? `${prefix}SubjectsBox` : "sectionExcelSubjectsBox");
    const subjectsList = document.getElementById(prefix ? `${prefix}SubjectsList` : "excelSubjectsList");
    if (subjectsBox && subjectsList) {
        subjectsList.innerHTML = "";
        if (parsedData.subjects.length > 0) {
            parsedData.subjects.forEach(sub => {
                const pill = document.createElement("span");
                pill.className = "subject-pill-tag";
                pill.textContent = sub;
                subjectsList.appendChild(pill);
            });
            subjectsBox.classList.remove("hidden");
        } else {
            subjectsBox.classList.add("hidden");
        }
    }

    // Students Preview (first 10)
    const studentsBox = document.getElementById(prefix ? `${prefix}StudentsBox` : "sectionExcelStudentsBox");
    const tableBody = document.getElementById(prefix ? `${prefix}StudentsTableBody` : "excelStudentsTableBody");
    const moreText = document.getElementById(prefix ? `${prefix}MoreStudentsText` : "excelMoreStudentsText");
    if (studentsBox && tableBody) {
        tableBody.innerHTML = "";
        const previewRows = parsedData.students.slice(0, 10);
        previewRows.forEach((st, idx) => {
            const tr = document.createElement("tr");
            tr.innerHTML = `<td>${idx + 1}</td><td>${escapeHtml(st.qid)}</td><td>${escapeHtml(st.name)}</td>`;
            tableBody.appendChild(tr);
        });

        if (moreText) {
            const remaining = parsedData.students.length - 10;
            if (remaining > 0) {
                moreText.textContent = `+ ${remaining} more students`;
                moreText.classList.remove("hidden");
            } else {
                moreText.textContent = "";
                moreText.classList.add("hidden");
            }
        }
        studentsBox.classList.remove("hidden");
    }

    previewEl.classList.remove("hidden");
}


// FIXED SECTION LIST

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
    {
        id: "CLOUD-TECH-INFOSEC",
        label: "Cloud Technology & Information Security"
    }
];


function sectionLabelOf(id) {

    const match =
        SECTIONS.find(
            section =>
                section.id === id
        );

    return match
        ? match.label
        : id;
}


// HTML SAFETY HELPER
// Escapes dynamic values before inserting them into HTML templates.
function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
window.escapeHtml = escapeHtml;


// NAME COMPARISON

function compareNames(nameA, nameB) {

    const aWords =
        nameA.trim().split(/\s+/);

    const bWords =
        nameB.trim().split(/\s+/);

    const wordCount =
        Math.max(
            aWords.length,
            bWords.length
        );

    for (let i = 0; i < wordCount; i++) {

        const aWord =
            aWords[i] || "";

        const bWord =
            bWords[i] || "";

        const comparison =
            aWord.localeCompare(
                bWord
            );

        if (comparison !== 0) {
            return comparison;
        }
    }

    return 0;
}


// STATE

let profile = null;

let activeSection = null;

let students = [];

let subjects = [];

let recordsCache = {};
let requestFilter = "pending";
let selectedRequestIds = new Set();
let sectionCourseMap = {};


// HTML ELEMENTS

const loginPage =
    document.getElementById(
        "loginPage"
    );

const appPage =
    document.getElementById(
        "appPage"
    );

const pageName = document.body.dataset.page || "login";

if (pageName === "login") {
    appPage?.remove();
} else {
    loginPage?.remove();
    document.getElementById("contactAdminModal")?.remove();
    document.getElementById("addSectionRequestModal")?.remove();
}

if (pageName === "admin") {
    document.querySelectorAll(".cr-only, #requestModal, #adminSidebar").forEach(element => element.remove());
}

if (pageName === "cr") {
    document.querySelectorAll(".admin-only, #adminSidebar, #adminDashboard, #adminCoursesView, #adminSectionsView, #adminPanel, #adminAddSectionModal, #courseModal, #addStudentsModal, #addSubjectsModal, #approveConfirmModal, #approveSuccessModal, #adminRequestTabs, #bulkRequestToolbar, #adminRequestsTableWrapper").forEach(element => element.remove());
}


const loginForm =
    document.getElementById(
        "loginForm"
    );

const emailInput =
    document.getElementById(
        "email"
    );

const passwordInput =
    document.getElementById(
        "password"
    );

const togglePassword =
    document.getElementById(
        "togglePassword"
    );

const loginMessage =
    document.getElementById(
        "loginMessage"
    );


const forgotPasswordLink =
    document.getElementById(
        "forgotPasswordLink"
    );

const resetPasswordBox =
    document.getElementById(
        "resetPasswordBox"
    );

const resetEmail =
    document.getElementById(
        "resetEmail"
    );

const sendResetBtn =
    document.getElementById(
        "sendResetBtn"
    );

const resetMessage =
    document.getElementById(
        "resetMessage"
    );

const backToLoginLink =
    document.getElementById(
        "backToLoginLink"
    );


const logoutBtn =
    document.getElementById(
        "logoutBtn"
    );

const loggedUser =
    document.getElementById(
        "loggedUser"
    );

const roleBadge =
    document.getElementById(
        "roleBadge"
    );


const welcomeTitle =
    document.getElementById(
        "welcomeTitle"
    );

const welcomeSubtitle =
    document.getElementById(
        "welcomeSubtitle"
    );

const sectionLabel =
    document.getElementById(
        "sectionLabel"
    );

const attendanceCourseLabel = document.getElementById("attendanceCourseLabel");
const attendanceCourseSelect = document.getElementById("attendanceCourseSelect");


const sectionSwitchBox =
    document.getElementById(
        "sectionSwitchBox"
    );

const sectionSelect =
    document.getElementById(
        "sectionSelect"
    );

const workArea =
    document.getElementById(
        "workArea"
    );

const noSectionMessage =
    document.getElementById(
        "noSectionMessage"
    );


const adminPanel =
    document.getElementById(
        "adminPanel"
    );

const adminOnlyEls =
    document.querySelectorAll(
        ".admin-only"
    );


const addStudentForm =
    document.getElementById(
        "addStudentForm"
    );

const newStudentQid =
    document.getElementById(
        "newStudentQid"
    );

const newStudentName =
    document.getElementById(
        "newStudentName"
    );


const addSubjectForm =
    document.getElementById(
        "addSubjectForm"
    );

const newSubjectName =
    document.getElementById(
        "newSubjectName"
    );

const subjectManageList =
    document.getElementById(
        "subjectManageList"
    );


const addCrForm =
    document.getElementById(
        "addCrForm"
    );

const newCrEmail =
    document.getElementById(
        "newCrEmail"
    );

const newCrSection =
    document.getElementById(
        "newCrSection"
    );

const crManageList =
    document.getElementById(
        "crManageList"
    );


const studentList =
    document.getElementById(
        "studentList"
    );

const subjectSelect =
    document.getElementById(
        "subjectSelect"
    );

const attendanceDate =
    document.getElementById(
        "attendanceDate"
    );

const searchStudent =
    document.getElementById(
        "searchStudent"
    );


const totalStudents =
    document.getElementById(
        "totalStudents"
    );

const headerTotalStudents =
    document.getElementById(
        "headerTotalStudents"
    );

const presentStudents =
    document.getElementById(
        "presentStudents"
    );

const absentStudents =
    document.getElementById(
        "absentStudents"
    );


const saveBtn =
    document.getElementById(
        "saveBtn"
    );

const downloadBtn =
    document.getElementById(
        "downloadBtn"
    );

const markAllPresentBtn =
    document.getElementById(
        "markAllPresentBtn"
    );

const markAllAbsentBtn =
    document.getElementById(
        "markAllAbsentBtn"
    );


const recordsList =
    document.getElementById(
        "recordsList"
    );

const refreshRecordsBtn =
    document.getElementById(
        "refreshRecordsBtn"
    );


// TODAY DATE

if (attendanceDate) {
    attendanceDate.value =
        new Date()
            .toISOString()
            .split("T")[0];
}


// PASSWORD SHOW / HIDE

togglePassword?.addEventListener(
    "click",
    function () {

        if (
            passwordInput.type ===
            "password"
        ) {

            passwordInput.type =
                "text";

            togglePassword.textContent =
                "🙈";

        } else {

            passwordInput.type =
                "password";

            togglePassword.textContent =
                "👁";
        }
    }
);


// LOGIN

loginForm?.addEventListener(
    "submit",
    async function (event) {

        event.preventDefault();

        const email =
            emailInput.value.trim();

        const password =
            passwordInput.value;

        loginMessage.textContent =
            "Logging in...";

        try {

            await authPersistenceReady;

            await signInWithEmailAndPassword(
                auth,
                email,
                password
            );

            loginMessage.textContent =
                "";

        } catch (error) {

            console.error(error);

            loginMessage.textContent =
                "Invalid email or password.";
        }
    }
);


// LOGIN-PAGE ACCESS MODALS

const contactAdminLink = document.getElementById("contactAdminLink");
const contactAdminModal = document.getElementById("contactAdminModal");
const closeContactAdminModal = document.getElementById("closeContactAdminModal");
const cancelContactAdminBtn = document.getElementById("cancelContactAdminBtn");
const contactAdminForm = document.getElementById("contactAdminForm");
const contactAdminMessage = document.getElementById("contactAdminMessage");
const requestSectionLink = document.getElementById("requestSectionLink");
const addSectionRequestModal = document.getElementById("addSectionRequestModal");
const closeAddSectionRequestModal = document.getElementById("closeAddSectionRequestModal");
const cancelAddSectionRequestBtn = document.getElementById("cancelAddSectionRequestBtn");
const addSectionRequestForm = document.getElementById("addSectionRequestForm");
const sectionExcelFile = document.getElementById("sectionExcelFile");
const sectionExcelFileLabel = document.getElementById("sectionExcelFileLabel");
const sectionExcelPreview = document.getElementById("sectionExcelPreview");
const sectionRequestMessage = document.getElementById("sectionRequestMessage");
const submitAddSectionRequestBtn = document.getElementById("submitAddSectionRequestBtn");

let sectionExcelPayload = null;

function openLoginModal(modal) {
    modal?.classList.remove("hidden");
    document.body.classList.add("modal-open");
}
function closeLoginModal(modal) {
    modal?.classList.add("hidden");
    if (!document.querySelector(".modal-backdrop:not(.hidden)")) document.body.classList.remove("modal-open");
}

contactAdminLink?.addEventListener("click", () => {
    contactAdminMessage.textContent = "";
    openLoginModal(contactAdminModal);
});
closeContactAdminModal?.addEventListener("click", () => closeLoginModal(contactAdminModal));
cancelContactAdminBtn?.addEventListener("click", () => closeLoginModal(contactAdminModal));

contactAdminForm?.addEventListener("submit", async event => {
    event.preventDefault();
    const name = document.getElementById("contactName").value.trim();
    const email = document.getElementById("contactEmail").value.trim();
    const mobile = document.getElementById("contactMobile").value.trim();
    const type = document.getElementById("contactType")?.value.trim() || "Access / Support Request";
    const message = document.getElementById("contactMessage").value.trim();

    if (!/^\d{10}$/.test(mobile)) {
        contactAdminMessage.textContent = "Mobile number must be exactly 10 digits.";
        contactAdminMessage.className = "form-feedback error";
        return;
    }

    contactAdminMessage.textContent = "Sending email to Admin...";
    contactAdminMessage.className = "form-feedback";

    try {
        await addDoc(collection(db, "requests"), {
            type: "contact_admin",
            status: "pending",
            name,
            email,
            mobile,
            message,
            requestedBy: email,
            requestedByRole: "guest",
            createdAt: new Date().toISOString()
        });

        await sendDirectEmail({
            toEmail: EMAILJS_CONFIG.adminEmail,
            subject: `QAttend Access Request - ${name}`,
            message: `New Contact Admin Request\n\nName: ${name}\nEmail: ${email}\nMobile: ${mobile}\nRequest Type: ${type}\n\nMessage:\n${message}`,
            replyTo: email
        });

        contactAdminMessage.textContent = "Email sent successfully to Admin.";
        contactAdminMessage.className = "form-feedback success";
        contactAdminForm.reset();
        setTimeout(() => closeLoginModal(contactAdminModal), 1000);
    } catch (error) {
        console.error("Contact Admin email error:", error);
        contactAdminMessage.textContent = error.message || "Could not send email. Please try again.";
        contactAdminMessage.className = "form-feedback error";
    }
});

requestSectionLink?.addEventListener("click", () => {
    sectionRequestMessage.textContent = "";
    sectionExcelPayload = null;
    if (sectionExcelFile) sectionExcelFile.value = "";
    if (sectionExcelFileLabel) sectionExcelFileLabel.textContent = "Click to upload Excel file";
    sectionExcelPreview?.classList.add("hidden");
    openLoginModal(addSectionRequestModal);
});
closeAddSectionRequestModal?.addEventListener("click", () => closeLoginModal(addSectionRequestModal));
cancelAddSectionRequestBtn?.addEventListener("click", () => closeLoginModal(addSectionRequestModal));

sectionExcelFile?.addEventListener("change", async () => {
    const file = sectionExcelFile.files?.[0];
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
        sectionRequestMessage.textContent = "Please upload an .xlsx or .xls file.";
        sectionRequestMessage.className = "section-request-message error";
        sectionExcelFile.value = "";
        return;
    }
    sectionExcelFileLabel.textContent = file.name;
    sectionRequestMessage.textContent = "Reading Excel file...";
    sectionRequestMessage.className = "section-request-message";
    try {
        await ensureXLSXLoaded();
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, {type:"array"});
        const normalize = v => String(v ?? "").trim();
        const findSheet = name => {
            const target = name.replace(/\s+/g, "").toLowerCase();
            const sheetName = workbook.SheetNames.find(candidate =>
                candidate.replace(/\s+/g, "").toLowerCase() === target
            );
            return sheetName ? workbook.Sheets[sheetName] : null;
        };
        const studentsSheet = findSheet("Students") || workbook.Sheets[workbook.SheetNames[0]];
        const subjectsSheet = findSheet("Subjects") || (workbook.SheetNames[1] ? workbook.Sheets[workbook.SheetNames[1]] : null);
        const studentRows = studentsSheet ? XLSX.utils.sheet_to_json(studentsSheet, {header:1, defval:""}) : [];
        const subjectRows = subjectsSheet ? XLSX.utils.sheet_to_json(subjectsSheet, {header:1, defval:""}) : [];

        const studentsOut = [];
        const qids = new Set();
        let invalid = 0;
        studentRows.slice(1).forEach(row => {
            const qid = normalize(row[0]);
            const name = normalize(row[1]);
            if (!qid && !name) return;
            if (!qid || !name || qids.has(qid.toLowerCase())) { invalid++; return; }
            qids.add(qid.toLowerCase());
            studentsOut.push({qid, name: name.toUpperCase()});
        });

        const subjectsOut = [];
        const subjectSet = new Set();
        subjectRows.slice(1).forEach(row => {
            const name = normalize(row[0]);
            if (!name) return;
            const key = name.toLowerCase();
            if (!subjectSet.has(key)) { subjectSet.add(key); subjectsOut.push(name); }
        });

        sectionExcelPayload = {students:studentsOut, subjects:subjectsOut, duplicateCount:0, invalidCount:invalid};
        document.getElementById("excelStudentCount").textContent = studentsOut.length;
        document.getElementById("excelSubjectCount").textContent = subjectsOut.length;
        document.getElementById("excelDuplicateCount").textContent = 0;
        document.getElementById("excelInvalidCount").textContent = invalid;
        sectionExcelPreview.classList.remove("hidden");
        sectionRequestMessage.textContent = studentsOut.length ? "Excel validated. You can submit the request." : "No valid students found in the Excel file.";
        sectionRequestMessage.className = `section-request-message ${studentsOut.length ? "success" : "error"}`;
    } catch (error) {
        console.error(error);
        sectionExcelPayload = null;
        sectionExcelPreview?.classList.add("hidden");
        sectionRequestMessage.textContent = "Could not read the Excel file. Please use the provided template.";
        sectionRequestMessage.className = "section-request-message error";
    }
});

addSectionRequestForm?.addEventListener("submit", async event => {
    event.preventDefault();
    const get = id => document.getElementById(id)?.value.trim() || "";
    const crQid=get("reqCrQid"), crName=get("reqCrName"), crMobile=get("reqCrMobile"), crEmail=get("reqCrEmail").toLowerCase();
    const mentorName=get("reqMentorName"), mentorMobile=get("reqMentorMobile"), course=get("reqCourse"), sectionName=get("reqSectionName"), semester=get("reqSemester"), year=get("reqYear");
    if (!/^\d{10}$/.test(crMobile) || !/^\d{10}$/.test(mentorMobile)) {
        sectionRequestMessage.textContent = "CR Mobile and Mentor No. must be exactly 10 digits.";
        sectionRequestMessage.className = "section-request-message error";
        return;
    }
    if (!sectionExcelPayload || !sectionExcelPayload.students.length) {
        sectionRequestMessage.textContent = "Please download, fill and upload the Excel template before submitting.";
        sectionRequestMessage.className = "section-request-message error";
        return;
    }
    submitAddSectionRequestBtn.disabled = true;
    sectionRequestMessage.textContent = "Sending request to Admin...";
    sectionRequestMessage.className = "section-request-message";
    const sectionId = `${course}-${sectionName}`.toUpperCase().replace(/[^A-Z0-9]+/g,"-").replace(/^-|-$/g,"");
    const data = {
        type:"add_section", status:"pending", requestedBy:crEmail, requestedByRole:"guest",
        crQid, crName, crMobile, crEmail, mentorName, mentorMobile, course, section:sectionId,
        sectionLabel:sectionName, semester, year, students:sectionExcelPayload.students, subjects:sectionExcelPayload.subjects,
        excelStats:{students:sectionExcelPayload.students.length,subjects:sectionExcelPayload.subjects.length,duplicates:sectionExcelPayload.duplicateCount,invalid:sectionExcelPayload.invalidCount},
        createdAt:new Date().toISOString()
    };
    try {
        await addDoc(collection(db,"requests"),data);
        const subject=`QAttend - New Section Request - ${sectionName}`;
        const body=`New Section Request\n\nCR Q.ID: ${crQid}\nCR Name: ${crName}\nCR Mobile: ${crMobile}\nCR Email: ${crEmail}\n\nMentor Name: ${mentorName}\nMentor No.: ${mentorMobile}\n\nCourse: ${course}\nSection: ${sectionName}\nSemester: ${semester}\nYear: ${year}\n\nStudents: ${data.students.length}\nSubjects: ${data.subjects.join(", ")}\n\nPlease review this request in QAttend Admin Request Center.`;

        sectionRequestMessage.textContent = "Request saved. Sending email to Admin...";
        sectionRequestMessage.className = "section-request-message";

        try {
            await sendDirectEmail({
                toEmail: EMAILJS_CONFIG.adminEmail,
                subject,
                message: body,
                replyTo: crEmail
            });
            sectionRequestMessage.textContent = "Request submitted and email sent to Admin.";
            sectionRequestMessage.className = "section-request-message success";
            setTimeout(() => closeLoginModal(addSectionRequestModal), 1000);
        } catch (emailError) {
            console.error("New Section email error:", emailError);
            sectionRequestMessage.textContent = `Request was saved, but email could not be sent. ${emailError.message || "Please try again."}`;
            sectionRequestMessage.className = "section-request-message error";
        }
    } catch (error) {
        console.error(error);
        sectionRequestMessage.textContent = "Could not submit request. Check Firestore Rules and try again.";
        sectionRequestMessage.className = "section-request-message error";
    } finally { submitAddSectionRequestBtn.disabled = false; }
});

[contactAdminModal, addSectionRequestModal].forEach(modal => modal?.addEventListener("click", e => { if(e.target===modal) closeLoginModal(modal); }));

// FORGOT PASSWORD

forgotPasswordLink?.addEventListener(
    "click",
    function (event) {

        event.preventDefault();

        resetEmail.value =
            emailInput.value.trim();

        resetMessage.textContent =
            "";

        resetMessage.classList.remove(
            "login-message-success"
        );

        loginForm.classList.add(
            "hidden"
        );

        forgotPasswordLink.parentElement.classList.add(
            "hidden"
        );

        loginMessage.classList.add(
            "hidden"
        );

        resetPasswordBox.classList.remove(
            "hidden"
        );
    }
);


// BACK TO LOGIN

backToLoginLink?.addEventListener(
    "click",
    function (event) {

        event.preventDefault();

        resetPasswordBox.classList.add(
            "hidden"
        );

        loginForm.classList.remove(
            "hidden"
        );

        forgotPasswordLink.parentElement.classList.remove(
            "hidden"
        );

        loginMessage.classList.remove(
            "hidden"
        );

        resetMessage.textContent =
            "";
    }
);


// SEND RESET LINK

sendResetBtn?.addEventListener(
    "click",
    async function () {

        const email =
            resetEmail.value.trim();


        if (!email) {

            resetMessage.classList.remove(
                "login-message-success"
            );

            resetMessage.textContent =
                "Please enter your email.";

            return;
        }


        resetMessage.classList.remove(
            "login-message-success"
        );


        resetMessage.textContent =
            "Sending...";


        sendResetBtn.disabled =
            true;


        try {

            await sendPasswordResetEmail(
                auth,
                email
            );


            resetMessage.classList.add(
                "login-message-success"
            );


            resetMessage.textContent =
                "✅ Reset link sent! Check your inbox (and spam/promotions folder).";


        } catch (error) {

            console.error(error);


            resetMessage.classList.remove(
                "login-message-success"
            );


            if (
                error.code ===
                "auth/invalid-email"
            ) {

                resetMessage.textContent =
                    "That doesn't look like a valid email.";


            } else if (
                error.code ===
                "auth/too-many-requests"
            ) {

                resetMessage.textContent =
                    "Too many attempts. Please try again in a while.";


            } else {

                resetMessage.textContent =
                    "Could not send reset email. Try again.";
            }


        } finally {

            sendResetBtn.disabled =
                false;
        }
    }
);


// AUTH STATE

onAuthStateChanged(
    auth,
    async function (user) {

        if (!user) {

            const earlyPage =
                (window.location.pathname.split("/").pop() || "login.html").toLowerCase();

            if (earlyPage === "admin.html" || earlyPage === "cr.html") {
                window.location.replace("login.html");
                return;
            }

            appPage?.classList.add(
                "hidden"
            );

            loginPage?.classList.remove(
                "hidden"
            );

            profile = null;

            document.body.classList.remove(
                "admin-view",
                "cr-view"
            );


            resetPasswordBox?.classList.add(
                "hidden"
            );

            loginForm?.classList.remove(
                "hidden"
            );

            forgotPasswordLink?.parentElement?.classList.remove(
                "hidden"
            );

            loginMessage?.classList.remove(
                "hidden"
            );

            return;
        }


        try {

            const earlyPage =
                (window.location.pathname.split("/").pop() || "login.html").toLowerCase();
            let cachedProfile = null;
            try {
                const cached = JSON.parse(localStorage.getItem("qattend-profile") || "null");
                if (cached?.email === String(user.email || "").toLowerCase() && cached.profile?.role) {
                    cachedProfile = cached.profile;
                }
            } catch (_) {}

            if (cachedProfile && (earlyPage === "login.html" || earlyPage === "index.html" || earlyPage === "")) {
                window.location.replace(cachedProfile.role === "admin" ? "admin.html" : "cr.html");
                return;
            }

            const profileSnap =
                await getDoc(
                    doc(
                        db,
                        "users",
                        String(user.email || "").toLowerCase()
                    )
                );


            if (
                !profileSnap.exists()
            ) {

                alert(
                    "Your account is not set up yet. Ask the Admin to assign you a role."
                );

                await signOut(
                    auth
                );

                return;
            }


            profile =
                profileSnap.data();

            try {
                localStorage.setItem("qattend-profile", JSON.stringify({
                    email: String(user.email || "").toLowerCase(),
                    profile,
                    cachedAt: Date.now()
                }));
            } catch (_) {}

            // Separate page routing: login.html -> admin.html / cr.html.
            // admin.html and cr.html also reject the wrong role.
            const currentPage =
                (window.location.pathname.split("/").pop() || "login.html").toLowerCase();

            if (currentPage === "login.html" || currentPage === "index.html" || currentPage === "") {
                window.location.replace(profile.role === "admin" ? "admin.html" : "cr.html");
                return;
            }

            if (currentPage === "admin.html" && profile.role !== "admin") {
                window.location.replace("cr.html");
                return;
            }

            if (currentPage === "cr.html" && profile.role === "admin") {
                window.location.replace("admin.html");
                return;
            }

            // Record the successful login. Logging failures must never block the app.
            void writeActivityLog("LOGIN", "User signed in to QAttend.");


        } catch (error) {

            console.error(error);

            alert(
                "Could not load your account role. Try logging in again."
            );

            await signOut(
                auth
            );

            return;
        }


        loginPage?.classList.add(
            "hidden"
        );

        appPage?.classList.remove(
            "hidden"
        );


        if (loggedUser) loggedUser.textContent = user.email;


        if (roleBadge) {
            roleBadge.textContent =
                profile.role === "admin"
                    ? "ADMIN"
                    : "CLASS REPRESENTATIVE";
        }

        const drawerRequestsTitle =
            document.getElementById("drawerRequestsTitle");
        const drawerRequestsText =
            document.getElementById("drawerRequestsText");

        if (drawerRequestsTitle) {
            drawerRequestsTitle.textContent =
                profile.role === "admin" ? "Requests" : "Send Request";
        }

        if (drawerRequestsText) {
            drawerRequestsText.textContent =
                profile.role === "admin"
                    ? "View all CR student and subject requests"
                    : "Send student or subject change request to Admin";
        }


        const isAdmin =
            profile.role === "admin";


        // Role-specific layout classes.
        // Admin and CR have different responsive layouts.

        document.body.classList.remove(
            "admin-view",
            "cr-view"
        );

        document.body.classList.add(
            isAdmin
                ? "admin-view"
                : "cr-view"
        );


        adminOnlyEls.forEach(
            function (el) {

                el.classList.toggle(
                    "hidden",
                    !isAdmin
                );
            }
        );


        if (isAdmin) {

            // Admin lands on the Admin Dashboard after login.
            // Mark Attendance is opened explicitly from the hamburger menu.
            showAdminDashboard();

            activeSection = null;
            managementSection = null;
            if (sectionSelect) sectionSelect.value = "";
            if (manageSectionSelect) manageSectionSelect.value = "";

            void (async () => {
                try {
                    await ensureLegacyBtechCourse();
                    await populateSectionDropdowns();
                    await populateAttendanceCourseDropdown();
                } catch (migrationError) {
                    console.warn("Could not prepare admin course data:", migrationError);
                }
            })();

        } else {

            showAttendanceView();

            if (welcomeTitle) welcomeTitle.textContent =
                `${sectionLabelOf(
                    profile.section
                )} Attendance Portal`;


            if (welcomeSubtitle) welcomeSubtitle.textContent =
                "Mark, save and download attendance for your section.";


            activeSection =
                profile.section;


            workArea.classList.remove(
                "hidden"
            );


            noSectionMessage.classList.add(
                "hidden"
            );


            if (sectionLabel) sectionLabel.textContent =
                sectionLabelOf(
                    activeSection
                );


            await Promise.all([loadSubjects(), loadStudents()]);
        }
    }
);


// SECTION DROPDOWNS
// Sections are owned by Firestore. Never assume that a section belongs to
// B.Tech just because it appears in the legacy fixed list. A course/section
// created from a request must immediately become selectable here.
function normalizeAcademicKey(value) {
    return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

function resolveCourseIdForSection(sectionData, sectionId, courseDocs) {
    const rawValues = [
        sectionData?.course,
        sectionData?.courseId,
        sectionData?.courseName,
        sectionData?.courseCode
    ].map(normalizeAcademicKey).filter(Boolean);

    const match = courseDocs.find(courseDoc => {
        const course = courseDoc.data() || {};
        const candidates = [
            courseDoc.id,
            course.name,
            course.code
        ].map(normalizeAcademicKey).filter(Boolean);
        return rawValues.some(value => candidates.includes(value));
    });

    if (match) return match.id;

    // Keep legacy B.Tech data working if its section documents only contain
    // the old BTECH identifier and no matching course document is available.
    if (rawValues.includes("btech") || rawValues.includes("b.tech")) return "BTECH";
    return String(sectionData?.course || sectionData?.courseId || "").trim();
}

async function getAcademicData() {
    const [coursesSnap, sectionsSnap] = await Promise.all([
        getDocs(collection(db, "courses")),
        getDocs(collection(db, "sections"))
    ]);
    return {
        courses: coursesSnap.docs,
        sections: sectionsSnap.docs
    };
}

async function populateSectionDropdowns() {
    let courses = [];
    let sectionDocs = [];
    try {
        const data = await getAcademicData();
        courses = data.courses;
        sectionDocs = data.sections;
    } catch (error) {
        console.warn("Could not load dynamic sections:", error);
    }

    sectionCourseMap = {};
    const allSections = sectionDocs.map(sectionDoc => {
        const data = sectionDoc.data() || {};
        const canonicalCourseId = resolveCourseIdForSection(data, sectionDoc.id, courses);
        sectionCourseMap[sectionDoc.id] = canonicalCourseId;
        return {
            id: sectionDoc.id,
            label: data.label || data.section || sectionDoc.id,
            courseId: canonicalCourseId
        };
    });

    // Only include legacy fixed sections when they actually exist in Firestore.
    // This prevents the UI from displaying fake sections for a newly created course.
    const existingIds = new Set(allSections.map(section => section.id));
    SECTIONS.forEach(legacySection => {
        if (existingIds.has(legacySection.id)) return;
        // No Firestore document = no section option.
    });

    if (sectionSelect) {
        sectionSelect.innerHTML = `<option value="">-- Select Section --</option>`;
        sectionSelect.disabled = true;
    }
    if (newCrSection) newCrSection.innerHTML = `<option value="">-- Select Section --</option>`;

    const manageSection = document.getElementById("manageSectionSelect");
    if (manageSection) {
        manageSection.innerHTML = `<option value="">-- Select Section --</option>`;
        manageSection.disabled = !manageCourseSelect?.value;
    }

    allSections
        .sort((a, b) => String(a.label).localeCompare(String(b.label), undefined, { numeric: true, sensitivity: "base" }))
        .forEach(section => {
            if (sectionSelect) {
                const option = document.createElement("option");
                option.value = section.id;
                option.textContent = section.label;
                sectionSelect.appendChild(option);
            }
            if (newCrSection) {
                const option = document.createElement("option");
                option.value = section.id;
                option.textContent = section.label;
                newCrSection.appendChild(option);
            }
            if (manageSection) {
                const option = document.createElement("option");
                option.value = section.id;
                option.textContent = section.label;
                manageSection.appendChild(option);
            }
        });

    if (manageCourseSelect?.value) filterManageSections(manageCourseSelect.value);
}

async function populateAttendanceCourseDropdown() {
    if (!attendanceCourseSelect) return;
    const selected = attendanceCourseSelect.value;
    attendanceCourseSelect.innerHTML = `<option value="">-- Select Course --</option>`;

    try {
        const snap = await getDocs(collection(db, "courses"));
        const seen = new Set();
        snap.docs
            .sort((a, b) => String(a.data()?.name || a.id).localeCompare(String(b.data()?.name || b.id)))
            .forEach(courseDoc => {
                const course = courseDoc.data() || {};
                const name = String(course.name || course.code || courseDoc.id).trim();
                const key = normalizeAcademicKey(courseDoc.id) || normalizeAcademicKey(name);
                if (seen.has(key)) return;
                seen.add(key);
                const option = document.createElement("option");
                option.value = courseDoc.id;
                option.textContent = name;
                attendanceCourseSelect.appendChild(option);
            });
    } catch (error) {
        console.warn("Could not load attendance courses:", error);
    }

    if (selected) attendanceCourseSelect.value = selected;
}

async function ensureLegacyBtechCourse() {
    // Compatibility only. Older B.Tech installations used a hard-coded
    // BTECH course/section list. Never rewrite existing sections here.
    if (profile?.role !== "admin") return;

    try {
        const coursesSnap = await getDocs(collection(db, "courses"));
        const hasBtech = coursesSnap.docs.some(courseDoc => {
            const course = courseDoc.data() || {};
            const name = normalizeAcademicKey(course.name);
            const code = normalizeAcademicKey(course.code);
            return name === "b.tech" || name === "b.tech ai & ml" || code === "btech" || code === "btech-ai-ml";
        });

        if (!hasBtech) {
            // Do not manufacture a course when the database already uses the
            // request-driven model. B.Tech is created only by its own data flow.
            return;
        }
    } catch (error) {
        console.warn("Could not check legacy B.Tech course:", error);
    }
}

        welcome?.classList.remove("hidden");
    }

    control?.classList.remove("hidden");
    syncAttendanceGate();
}


async function loadAdminDashboardData(){
    if(profile?.role !== "admin") return;

    try{
        const [coursesSnap, sectionsSnap, requestsSnap] = await Promise.all([
            getDocs(collection(db, "courses")),
            getDocs(collection(db, "sections")),
            getDocs(collection(db, "requests"))
        ]);
        const usersSnap = await getDocs(collection(db,"users"));
        const crCount = usersSnap.docs.filter(d => d.data()?.role === "cr").length;

        document.getElementById("dashTotalCourses")?.replaceChildren(String(coursesSnap.size));
        document.getElementById("dashTotalSections")?.replaceChildren(String(sectionsSnap.size));
        document.getElementById("dashPendingRequests")?.replaceChildren(String(requestsSnap.docs.filter(d => d.data()?.status === "pending").length));

        const allSectionIds = new Set(SECTIONS.map(section => section.id));
        try {
            const dynamicSections = await getDocs(collection(db,"sections"));
            dynamicSections.docs.forEach(d => allSectionIds.add(d.id));
        } catch (e) {
            console.warn("Could not read dynamic sections for dashboard:", e);
        }
        const sectionTasks = [...allSectionIds].map(async sectionId => {
            const [studentSnap, subjectSnap] = await Promise.all([
                getDocs(collection(db,"sections",sectionId,"students")),
                getDocs(collection(db,"sections",sectionId,"subjects"))
            ]);
            return {students: studentSnap.size, subjects: subjectSnap.size};
        });

        const sectionCounts = await Promise.all(sectionTasks);
        const totalStudentsCount = sectionCounts.reduce((sum,x) => sum + x.students, 0);
        const totalSubjectsCount = sectionCounts.reduce((sum,x) => sum + x.subjects, 0);

        if(document.getElementById("dashTotalStudents")){
            document.getElementById("dashTotalStudents").textContent = totalStudentsCount;
        }
        if(document.getElementById("dashTotalCRs")){
            document.getElementById("dashTotalCRs").textContent = crCount;
        }
        if(document.getElementById("dashTotalSubjects")){
            document.getElementById("dashTotalSubjects").textContent = totalSubjectsCount;
        }

        await Promise.allSettled([
            loadAdminActivityLogs(),
            loadAdminRequestPreview()
        ]);
    }catch(e){
        console.error("Dashboard data failed:", e);
    }
}

function showAdminDashboard(){
    document.body.classList.remove("records-only-view");
    hideMainViews();
    adminDashboard?.classList.remove("hidden");
    void loadAdminDashboardData();
}

function showRequestsPanel(){
    document.body.classList.remove("records-only-view");
    hideMainViews();
    requestsPanel?.classList.remove("hidden");
    void loadRequestCenter();
}

function showManageData(){
    document.body.classList.remove("records-only-view");
    hideMainViews();
    adminPanel?.classList.remove("hidden");

    if(profile?.role !== "admin") return;

    void populateCourseOptions();
    void populateSectionDropdowns();

    if(manageSectionSelect){
        manageSectionSelect.value = managementSection || "";
    }

    if(managementSection){
        void refreshManageData();
    }else{
        subjectManageList.innerHTML = "";
        crManageList.innerHTML = "";
        if(manageSectionMessage) manageSectionMessage.textContent = "Please select a section to manage its data.";
    }
}

function filterManageSections(courseId) {
    if (!manageSectionSelect) return;
    manageSectionSelect.disabled = !courseId;
    const selectedSection = manageSectionSelect.value;

    [...manageSectionSelect.options].forEach(option => {
        if (!option.value) {
            option.hidden = false;
            return;
        }
        const mappedCourse = normalizeAcademicKey(sectionCourseMap[option.value]);
        const selectedCourse = normalizeAcademicKey(courseId);
        option.hidden = !selectedCourse || mappedCourse !== selectedCourse;
    });

    if (selectedSection && !manageSectionSelect.selectedOptions[0]?.hidden) return;
    managementSection = null;
    manageSectionSelect.value = "";
    subjectManageList.innerHTML = "";
    crManageList.innerHTML = "";
    if (manageSectionMessage) manageSectionMessage.textContent = courseId
        ? "Please select a section to manage its data."
        : "Please select a course first.";
}

function showPreviousRecords(){
    if(profile?.role !== "cr") return;
    document.body.classList.add("records-only-view");
    requestsPanel?.classList.add("hidden");
    adminDashboard?.classList.add("hidden");
    document.querySelector(".welcome-section")?.classList.add("hidden");
    document.querySelector(".control-section")?.classList.add("hidden");
    document.getElementById("noSectionMessage")?.classList.add("hidden");
    document.getElementById("attendanceGateMessage")?.classList.add("hidden");
    workArea?.classList.remove("hidden");
    const recordsSection = document.querySelector(".records-section");
    if(recordsSection){
        void loadRecords();
        setTimeout(() => recordsSection.scrollIntoView({behavior:"smooth", block:"start"}), 40);
    }
}

function openSupport(){
    supportModal?.classList.remove("hidden");
}

async function logoutFromApp(){
    try {
        await writeActivityLog("LOGOUT", "User signed out of QAttend.");
    } catch (e) {
        console.warn("Logout log failed:", e);
    }
    await signOut(auth);
    window.location.replace("login.html");
}



manageSectionSelect?.addEventListener("change", async () => {
    if(profile?.role !== "admin") return;
    managementSection = manageSectionSelect.value || null;

    if(!managementSection){
        subjectManageList.innerHTML = "";
        crManageList.innerHTML = "";
        if(manageSectionMessage) manageSectionMessage.textContent = "Please select a section to manage its data.";
        if(newCrSection) newCrSection.disabled = false;
        return;
    }

    await refreshManageData();
});

manageCourseSelect?.addEventListener("change", async () => {
    managementSection = null;
    if (manageSectionSelect) manageSectionSelect.value = "";
    filterManageSections(manageCourseSelect.value || "");
});

document.querySelectorAll(".drawer-item").forEach(btn => {
    btn.addEventListener("click", async () => {
        const feature = btn.dataset.feature;
        closeFeatureDrawerFn();

        if(feature === "dashboard" && profile?.role === "admin") showAdminDashboard();
        if(feature === "courses" && profile?.role === "admin") {
            hideMainViews();
            document.getElementById("adminCoursesView")?.classList.remove("hidden");
            await loadCourses();
        }
        if(feature === "attendance") showAttendanceView();
        if(feature === "requests") showRequestsPanel();
        if(feature === "manage" && profile?.role === "admin") showManageData();
        if(feature === "records" && profile?.role === "cr") showPreviousRecords();
        if(feature === "send-request" && profile?.role === "cr") openRequestModal();
        if(feature === "support") openSupport();
        if(feature === "logout") await logoutFromApp();
    });
});

document.getElementById("closeSupportModal")?.addEventListener("click",()=>supportModal?.classList.add("hidden"));
supportModal?.addEventListener("click",e=>{
    if(e.target===supportModal) supportModal.classList.add("hidden");
});

function friendlyLogAction(action){
    return {
        LOGIN:"Logged in",
        LOGOUT:"Logged out",
        ADD_STUDENT:"Student added",
        EDIT_STUDENT:"Student edited",
        DELETE_STUDENT:"Student deleted",
        ADD_SUBJECT:"Subject added",
        DELETE_SUBJECT:"Subject deleted",
        ASSIGN_CR:"CR assigned",
        REMOVE_CR:"CR removed",
        SAVE_ATTENDANCE:"Attendance saved",
        DELETE_ATTENDANCE:"Attendance record deleted",
        REQUEST_SUBMITTED:"Request submitted",
        REQUEST_APPROVED:"Request approved",
        REQUEST_REJECTED:"Request rejected",
        ADD_SECTION:"Section added"
    }[action] || action;
}

function logBadgeClass(action){
    if(action.includes("DELETE") || action.includes("REJECT")) return "danger";
    if(action.includes("REQUEST") || action.includes("ASSIGN")) return "warning";
    if(action === "LOGIN" || action.includes("SAVE") || action.includes("ADD") || action.includes("APPROVED")) return "success";
    return "";
}

async function writeActivityLog(action, details = "", extra = {}){
    const user = auth.currentUser;
    if(!user?.email) return false;
    try{
        await addDoc(collection(db,"activityLogs"),{
            action,
            details,
            userEmail:String(user.email).trim(),
            userRole:profile?.role || "unknown",
            createdAt:new Date().toISOString(),
            ...extra
        });
        return true;
    }catch(e){
        console.warn("Activity log failed:",e);
        return false;
    }
}

async function loadAdminActivityLogs(){
    if(profile?.role !== "admin" || !activityLogList) return;
    activityLogList.innerHTML = `<p class="empty-record">Loading activity...</p>`;
    try{
        // Client-side sort avoids a Firestore composite/index dependency.
        const snap = await getDocs(collection(db,"activityLogs"));
        const cutoff = Date.now() - (2 * 60 * 60 * 1000);
        const rows = snap.docs
            .map(d => ({id:d.id,...d.data()}))
            .filter(row => {
                const createdAt = new Date(row.createdAt || 0).getTime();
                return Number.isFinite(createdAt) && createdAt >= cutoff;
            })
            .sort((a,b)=>String(b.createdAt||"").localeCompare(String(a.createdAt||"")))
            .slice(0,100);

        activityLogList.innerHTML = "";
        if(!rows.length){
            activityLogList.innerHTML = `<p class="empty-record">No activity logs yet.</p>`;
            return;
        }

        rows.forEach(x=>{
            const item=document.createElement("div");
            item.className="activity-item";
            const date=x.createdAt?new Date(x.createdAt).toLocaleString():"";
            item.innerHTML=`<span class="activity-icon">◷</span><div><b>${escapeHtml(friendlyLogAction(x.action))}<span class="log-badge ${logBadgeClass(x.action)}">${escapeHtml(x.userRole||"")}</span></b><small>${escapeHtml(x.userEmail||"")} · ${escapeHtml(date)}</small><small>${escapeHtml(x.details||"")}</small></div>`;
            activityLogList.appendChild(item);
        });
    }catch(e){
        console.error(e);
        activityLogList.innerHTML = `<p class="empty-record">Could not load activity logs. Make sure the latest Firestore Rules are published.</p>`;
    }
}

async function loadAdminRequestPreview(){
    const previewTarget = dashboardRequestList || document.getElementById("dashboardRequestTableBody");
    if(profile?.role !== "admin" || !previewTarget) return;
    try{
        const snap=await getDocs(collection(db,"requests"));
        const rows=snap.docs
            .map(d=>({id:d.id,...d.data()}))
            .filter(x=>x.status==="pending")
            .sort((a,b)=>String(b.createdAt||"").localeCompare(String(a.createdAt||"")))
            .slice(0,5);

        previewTarget.innerHTML="";
        if(!rows.length){
            previewTarget.innerHTML=`<tr><td colspan="7" class="empty-record">No pending requests.</td></tr>`;
            return;
        }

        rows.forEach(r=>{
            const el=document.createElement(previewTarget.tagName === "TBODY" ? "tr" : "div");
            el.className=previewTarget.tagName === "TBODY" ? "" : "activity-item";
            const title = r.type?.includes("student")
                ? (r.type==="add_student"?"Add Student":"Delete Student")
                : r.type?.includes("subject")
                    ? (r.type==="add_subject"?"Add Subject":"Delete Subject")
                    : r.type === "add_section" ? "Add New Section" : "Request";
            el.innerHTML=previewTarget.tagName === "TBODY"
                ? `<td>${escapeHtml(r.crName || r.requestedBy || "")}</td><td>${escapeHtml(r.sectionLabel || sectionLabelOf(r.section || ""))}</td><td>${escapeHtml(r.course || "")}</td><td>${r.students?.length || "-"}</td><td>${r.subjects?.length || "-"}</td><td>${escapeHtml(r.createdAt ? new Date(r.createdAt).toLocaleDateString() : "")}</td><td>${escapeHtml(String(r.status || "pending"))}</td>`
                : `<span class="activity-icon">↗</span><div><b>${escapeHtml(title||"Request")}</b><small>${escapeHtml(r.requestedBy||"")} · ${escapeHtml(r.sectionLabel||sectionLabelOf(r.section||""))}</small><small>${escapeHtml(r.reason||"")}</small></div>`;
            previewTarget.appendChild(el);
        });
    }catch(e){
        console.error(e);
        previewTarget.innerHTML=previewTarget.tagName === "TBODY"
            ? `<tr><td colspan="7" class="empty-record">Could not load requests.</td></tr>`
            : `<p class="empty-record">Could not load requests. Make sure the latest Firestore Rules are published.</p>`;
    }
}

document.getElementById("refreshLogsBtn")?.addEventListener("click",loadAdminActivityLogs);
document.getElementById("openLogsBtn")?.addEventListener("click",loadAdminActivityLogs);
document.getElementById("openRequestsFromDashboard")?.addEventListener("click",showRequestsPanel);
document.getElementById("backFromRequestsBtn")?.addEventListener("click",()=> profile?.role === "admin" ? showAdminDashboard() : showAttendanceView());

function showRequestDetails(request) {
    const modal = document.getElementById("viewRequestDetailModal");
    const content = document.getElementById("viewRequestDetailContent");
    if (!modal || !content) return;
    const values = request.type === "contact_admin"
        ? [
            ["Type", "Contact Admin"], ["Name", request.name], ["Email", request.email],
            ["Mobile", request.mobile], ["Message", request.message], ["Status", request.status]
        ]
        : [
            ["Type", request.type === "add_section" ? "New Section" : friendlyRequestType(request.type)],
            ["Requested By", request.crName || request.requestedBy], ["Course", request.course],
            ["Section", request.sectionLabel || request.section], ["Reason", request.reason],
            ["Students", request.students?.length || request.name || request.qid || "-"],
            ["Subjects", request.subjects?.length || request.subject || "-"], ["Status", request.status]
        ];
    document.getElementById("viewRequestDetailTitle").textContent = "Request Details";
    content.innerHTML = values.map(([label, value]) => `<div class="summary-row"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value || "-")}</strong></div>`).join("");
    if (request.type === "add_section") {
        const downloadButton = document.createElement("button");
        downloadButton.type = "button";
        downloadButton.className = "excel-download-btn";
        downloadButton.textContent = "Download Excel Data";
        downloadButton.addEventListener("click", () => downloadRequestExcel(request));
        content.appendChild(downloadButton);
    }
    modal.classList.remove("hidden");
    document.body.classList.add("modal-open");
}

async function downloadRequestExcel(request) {
    await ensureXLSXLoaded();
    const workbook = XLSX.utils.book_new();
    const students = [["QID", "Student Name"], ...(request.students || []).map(student => [student.qid || "", student.name || ""] )];
    const subjects = [["Subject Name"], ...(request.subjects || []).map(subject => [subject])];
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(students), "Students");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(subjects), "Subjects");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
        ["Course", request.course || ""], ["Section", request.sectionLabel || request.section || ""],
        ["Semester", request.semester || ""], ["Year", request.year || ""]
    ]), "Instructions");
    XLSX.writeFile(workbook, `Section_Request_${String(request.sectionLabel || request.section || "data").replace(/[^a-z0-9]+/gi, "_")}.xlsx`);
}

function closeRequestDetails() {
    document.getElementById("viewRequestDetailModal")?.classList.add("hidden");
    document.body.classList.remove("modal-open");
}

async function deleteRequest(request) {
    if (profile?.role !== "admin" || !request?.id) return;
    if (!confirm("Delete this request permanently?")) return;
    try {
        await deleteDoc(doc(db, "requests", request.id));
        selectedRequestIds.delete(request.id);
        await loadRequestCenter();
        await loadCourseSectionRequests();
        await loadAdminDashboardData();
    } catch (error) {
        console.error("Could not delete request:", error);
        alert("Could not delete request. Check Firestore rules and try again.");
    }
}

document.getElementById("closeViewRequestDetailModal")?.addEventListener("click", closeRequestDetails);
document.getElementById("closeViewRequestDetailBtn")?.addEventListener("click", closeRequestDetails);

async function loadRequestCenter(){
    const adminTableBody = document.getElementById("adminRequestsTableBody");
    const target = profile?.role === "admin" ? adminTableBody : requestList;
    if(!target || !profile) return;

    if(profile.role === "admin") {
        target.innerHTML = `<tr><td colspan="9" class="empty-record">Loading requests...</td></tr>`;
    } else {
        target.innerHTML=`<p class="empty-record">Loading requests...</p>`;
    }

    try{
        let snap;
        if(profile.role === "cr"){
            const email = String(auth.currentUser?.email || "").trim();
            snap = await getDocs(
                query(
                    collection(db,"requests"),
                    where("requestedBy","==",email)
                )
            );
        }else{
            snap = await getDocs(collection(db,"requests"));
        }
        const rows=snap.docs.map(d=>({id:d.id,...d.data()}));
        window.__qattendRequestCache = Object.fromEntries(rows.map(r=>[r.id,r]));

        rows.sort((a,b)=>String(b.createdAt||"").localeCompare(String(a.createdAt||"")));

        const filtered=rows.filter(r=>{
            if(profile.role !== "admin") return true;
            if(requestFilter==="pending") return r.status==="pending";
            if(requestFilter==="student") return String(r.type||"").includes("student");
            if(requestFilter==="subject") return String(r.type||"").includes("subject");
            if(requestFilter==="section") return r.type === "add_section";
            return true;
        });

        const visiblePendingIds = new Set(filtered.filter(r=>profile.role === "admin" && r.status === "pending").map(r=>r.id));
        selectedRequestIds = new Set([...selectedRequestIds].filter(id=>visiblePendingIds.has(id)));
        updateBulkRequestToolbar(filtered);
        target.innerHTML="";

        if(profile.role === "admin") {
            const pendingCount = rows.filter(row => row.status === "pending").length;
            const pendingBadge = document.getElementById("tabCountPending");
            if (pendingBadge) pendingBadge.textContent = pendingCount;

            if (!filtered.length) {
                target.innerHTML = `<tr><td colspan="9" class="empty-record">No requests found.</td></tr>`;
            } else {
                filtered.forEach(request => {
                    const row = document.createElement("tr");
                    const requestTitle = request.type === "add_section" ? "New Section" : request.type === "contact_admin" ? "Contact Admin" : friendlyRequestType(request.type);
                    const requester = request.crName || request.name || request.requestedBy || request.email || "Guest";
                    const course = request.course || "-";
                    const section = request.sectionLabel || request.section || "-";
                    const detail = request.type === "contact_admin"
                        ? `${request.email || ""} · ${request.mobile || ""}`
                        : request.type.includes("student")
                            ? `${request.qid || ""} ${request.name || ""}`
                            : request.subject || `${request.students?.length || 0} students, ${request.subjects?.length || 0} subjects`;
                    const checkbox = request.status === "pending" ? `<input type="checkbox" class="request-checkbox" data-request-id="${escapeHtml(request.id)}" ${selectedRequestIds.has(request.id) ? "checked" : ""} aria-label="Select request">` : "";
                    row.innerHTML = `<td>${checkbox}</td><td><strong>${escapeHtml(requester)}</strong><small class="request-table-detail">${escapeHtml(requestTitle)} · ${escapeHtml(detail)}</small></td><td>${escapeHtml(course)}</td><td>${escapeHtml(section)}</td><td>${request.type === "add_section" ? request.students?.length || 0 : "-"}</td><td>${request.type === "add_section" ? request.subjects?.length || 0 : "-"}</td><td>${escapeHtml(request.createdAt ? new Date(request.createdAt).toLocaleDateString() : "")}</td><td><span class="status-pill status-${escapeHtml(request.status || "pending")}">${escapeHtml(String(request.status || "pending").toUpperCase())}</span></td><td><button type="button" class="text-btn view-request-btn">View</button>${request.status === "pending" ? ` <button type="button" class="approve-btn request-table-action">Approve</button> <button type="button" class="reject-btn request-table-action">Reject</button>` : ""} <button type="button" class="text-btn danger-text delete-request-btn">Delete</button></td>`;

                    row.querySelector(".request-checkbox")?.addEventListener("change", event => {
                        if (event.target.checked) selectedRequestIds.add(request.id);
                        else selectedRequestIds.delete(request.id);
                        updateBulkRequestToolbar(filtered);
                    });
                    row.querySelector(".view-request-btn")?.addEventListener("click", () => showRequestDetails(request));
                    if (request.type === "add_section") {
                        const excelButton = document.createElement("button");
                        excelButton.type = "button";
                        excelButton.className = "text-btn request-excel-btn";
                        excelButton.textContent = "View Excel";
                        excelButton.addEventListener("click", () => downloadRequestExcel(request));
                        row.querySelector("td:last-child")?.prepend(excelButton);
                    }
                    row.querySelector(".approve-btn")?.addEventListener("click", () => processChangeRequest(request, true));
                    row.querySelector(".reject-btn")?.addEventListener("click", () => processChangeRequest(request, false));
                    row.querySelector(".delete-request-btn")?.addEventListener("click", () => deleteRequest(request));
                    target.appendChild(row);
                });
            }
            return;
        }

        if(profile.role === "cr"){
            const intro=document.createElement("div");
            intro.className="request-card request-intro-card";
            intro.innerHTML=`<div class="request-top"><div class="request-main"><b>Send a change request to Admin</b><small>Your request will be reviewed by an Admin. You do not receive Admin access.</small></div><span class="status-pill status-pending">REQUEST</span></div>`;
            target.appendChild(intro);
        }

        if(!filtered.length){
            const empty=document.createElement("p");
            empty.className="empty-record";
            empty.textContent=profile.role === "cr" ? "No requests submitted yet." : "No requests found.";
            target.appendChild(empty);
        }else{
            filtered.forEach(r=>{
                const card=document.createElement("div");
                card.className="request-card";
                const title = r.type === "add_student" ? "Add Student" : r.type === "delete_student" ? "Delete Student" : r.type === "add_subject" ? "Add Subject" : "Delete Subject";
                const subjectOrStudent = r.type.includes("student") ? `${r.qid||""}${r.name?` · ${r.name}`:""}` : (r.subject||"");
                const selectedCheckbox = (profile.role === "admin" && r.status === "pending") ? `<label class="request-select-box" title="Select this request"><input type="checkbox" class="request-checkbox" data-request-id="${escapeHtml(r.id)}" ${selectedRequestIds.has(r.id)?"checked":""} aria-label="Select request"></label>` : "";
                card.innerHTML=`<div class="request-card-row">${selectedCheckbox}<div class="request-top"><div class="request-main"><b>${escapeHtml(title)}${subjectOrStudent?` · ${escapeHtml(subjectOrStudent)}`:""}</b><small>${profile.role === "admin" ? `CR: ${escapeHtml(r.requestedBy||"")} · ` : ""}${escapeHtml(r.sectionLabel||sectionLabelOf(r.section||""))}</small><small>${escapeHtml(r.reason||"")}</small><small>${escapeHtml(r.createdAt?new Date(r.createdAt).toLocaleString():"")}</small></div><span class="status-pill status-${escapeHtml(r.status||"pending")}">${escapeHtml(String(r.status||"pending").toUpperCase())}</span></div></div>`;

                if(profile.role === "admin" && r.status === "pending"){
                    const checkbox=card.querySelector(".request-checkbox");
                    checkbox?.addEventListener("change",()=>{
                        if(checkbox.checked) selectedRequestIds.add(r.id);
                        else selectedRequestIds.delete(r.id);
                        updateBulkRequestToolbar(filtered);
                    });

                    const actions=document.createElement("div");
                    actions.className="request-actions";
                    const approve=document.createElement("button");
                    approve.className="approve-btn";
                    approve.textContent="Approve";
                    const reject=document.createElement("button");
                    reject.className="reject-btn";
                    reject.textContent="Reject";
                    approve.onclick=()=>processChangeRequest(r,true);
                    reject.onclick=()=>processChangeRequest(r,false);
                    actions.append(approve,reject);
                    card.appendChild(actions);
                }

                target.appendChild(card);
            });
        }
    }catch(e){
        console.error(e);
        target.innerHTML=profile.role === "admin"
            ? `<tr><td colspan="9" class="empty-record">Could not load requests. Publish Firestore rules and try again.</td></tr>`
            : `<p class="empty-record">Could not load requests. Make sure the latest Firestore Rules are published.</p>`;
    }

    if(profile.role === "admin") await loadAdminRequestPreview();
}

function updateBulkRequestToolbar(filteredRows=[]){
    if(!bulkRequestToolbar || profile?.role !== "admin") return;
    bulkRequestToolbar.classList.remove("hidden");
    const pendingRows = filteredRows.filter(r=>r.status === "pending");
    const pendingIds = pendingRows.map(r=>r.id);
    const selectedCount = pendingIds.filter(id=>selectedRequestIds.has(id)).length;
    if(selectedRequestsCount) selectedRequestsCount.textContent = `${selectedCount} selected`;
    const hasSelection = selectedCount > 0;
    if(bulkApproveRequestsBtn) bulkApproveRequestsBtn.disabled = !hasSelection;
    if(bulkRejectRequestsBtn) bulkRejectRequestsBtn.disabled = !hasSelection;
    if(selectAllRequests){
        selectAllRequests.disabled = pendingIds.length === 0;
        selectAllRequests.checked = pendingIds.length > 0 && selectedCount === pendingIds.length;
        selectAllRequests.indeterminate = selectedCount > 0 && selectedCount < pendingIds.length;
    }
}

selectAllRequests?.addEventListener("change",()=>{
    const checkboxes=[...document.querySelectorAll(".request-checkbox")];
    if(selectAllRequests.checked){
        checkboxes.forEach(cb=>{ selectedRequestIds.add(cb.dataset.requestId); cb.checked=true; });
    }else{
        checkboxes.forEach(cb=>{ selectedRequestIds.delete(cb.dataset.requestId); cb.checked=false; });
    }
    updateBulkRequestToolbar(checkboxes.map(cb=>({id:cb.dataset.requestId,status:"pending"})));
});

async function processBulkRequests(approve){
    if(profile?.role !== "admin") return;
    const ids=[...selectedRequestIds];
    if(!ids.length) return;
    const actionText=approve ? "accept" : "reject";
    if(!confirm(`Are you sure you want to ${actionText} ${ids.length} selected request${ids.length>1?"s":""}?`)) return;

    const approveBtnState=bulkApproveRequestsBtn?.disabled;
    const rejectBtnState=bulkRejectRequestsBtn?.disabled;
    if(bulkApproveRequestsBtn) bulkApproveRequestsBtn.disabled=true;
    if(bulkRejectRequestsBtn) bulkRejectRequestsBtn.disabled=true;

    let success=0, failed=0;
    try{
        for(const id of ids){
            const request = (window.__qattendRequestCache || {})[id];
            if(!request) { failed++; continue; }
            try{
                await processChangeRequest(request,approve,{silent:true});
                success++;
                selectedRequestIds.delete(id);
            }catch(error){
                console.error(`Bulk ${actionText} failed for ${id}:`,error);
                failed++;
            }
        }
        await loadRequestCenter();
        await loadAdminDashboardData();
        await populateSectionDropdowns();
        alert(`${approve?"✅ Accepted":"✅ Rejected"} ${success} request${success===1?"":"s"}${failed?` · ${failed} failed`:""}.`);
    }finally{
        if(bulkApproveRequestsBtn && approveBtnState===false) bulkApproveRequestsBtn.disabled=false;
        if(bulkRejectRequestsBtn && rejectBtnState===false) bulkRejectRequestsBtn.disabled=false;
    }
}

bulkApproveRequestsBtn?.addEventListener("click",()=>processBulkRequests(true));
bulkRejectRequestsBtn?.addEventListener("click",()=>processBulkRequests(false));

document.querySelectorAll(".request-tab").forEach(tab=>tab.addEventListener("click",()=>{
    document.querySelectorAll(".request-tab").forEach(x=>x.classList.remove("active"));
    tab.classList.add("active");
    requestFilter=tab.dataset.requestFilter||"all";
    void loadRequestCenter();
}));

function populateRequestStudentOptions(){
    const select=document.getElementById("requestStudentSelect");
    if(!select) return;
    select.innerHTML=`<option value="">-- Select Student --</option>`;
    students.forEach(s=>{
        const option=document.createElement("option");
        option.value=s.qid;
        option.textContent=`${s.name} — ${s.qid}`;
        select.appendChild(option);
    });
}

function populateRequestSubjectOptions(){
    const select=document.getElementById("requestSubjectSelect");
    if(!select) return;
    select.innerHTML=`<option value="">-- Select Subject --</option>`;
    subjects.forEach(s=>{
        const option=document.createElement("option");
        option.value=s.name;
        option.textContent=s.name;
        select.appendChild(option);
    });
}

function setInputRequired(id, required){
    const el=document.getElementById(id);
    if(el) el.required = required;
}

function openRequestModal(requestType=""){
    if(profile?.role !== "cr") return;
    populateRequestStudentOptions();
    populateRequestSubjectOptions();
    requestModal?.classList.remove("hidden");
    document.body.classList.add("modal-open");

    const choicePanel=document.getElementById("requestChoicePanel");
    const form=document.getElementById("requestForm");
    if(choicePanel) choicePanel.classList.toggle("hidden",!!requestType);
    if(form) form.classList.toggle("hidden",!requestType);

    const type=document.getElementById("requestType");
    if(type) type.value=requestType || "add_student";
    if(requestType) syncRequestForm();
}

function closeRequestModal(){
    requestModal?.classList.add("hidden");
    document.body.classList.remove("modal-open");
    document.getElementById("requestChoicePanel")?.classList.remove("hidden");
    document.getElementById("requestForm")?.classList.add("hidden");
}

document.querySelectorAll("[data-request-type]").forEach(button=>{
    button.addEventListener("click",()=>openRequestModal(button.dataset.requestType||"add_student"));
});

document.getElementById("requestBackBtn")?.addEventListener("click",()=>openRequestModal());

function syncRequestForm(){
    const type=document.getElementById("requestType")?.value || "add_student";
    const student=type.includes("student");
    const addStudent=type === "add_student";
    const deleteStudent=type === "delete_student";
    const addSubject=type === "add_subject";
    const deleteSubject=type === "delete_subject";

    const titleMap={add_student:"Add Student Request",delete_student:"Delete Student Request",add_subject:"Add Subject Request",delete_subject:"Delete Subject Request"};
    const formTitle=document.getElementById("requestFormTitle");
    if(formTitle) formTitle.textContent=titleMap[type] || "Send Request";

    document.getElementById("requestStudentFields")?.classList.toggle("hidden",!student);
    document.getElementById("requestSubjectFields")?.classList.toggle("hidden",student);

    document.getElementById("requestQidGroup")?.classList.toggle("hidden",!addStudent);
    document.getElementById("requestStudentNameGroup")?.classList.toggle("hidden",!addStudent);
    document.getElementById("requestStudentSelectGroup")?.classList.toggle("hidden",!deleteStudent);

    document.getElementById("requestSubjectNameGroup")?.classList.toggle("hidden",!addSubject);
    document.getElementById("requestSubjectSelectGroup")?.classList.toggle("hidden",!deleteSubject);

    setInputRequired("requestQid",addStudent);
    setInputRequired("requestStudentName",addStudent);
    setInputRequired("requestStudentSelect",deleteStudent);
    setInputRequired("requestSubjectName",addSubject);
    setInputRequired("requestSubjectSelect",deleteSubject);
}

document.getElementById("requestType")?.addEventListener("change",syncRequestForm);
document.getElementById("crSendRequestBtn")?.addEventListener("click",()=>openRequestModal("add_student"));
document.getElementById("closeRequestModal")?.addEventListener("click",closeRequestModal);
document.getElementById("cancelRequestBtn")?.addEventListener("click",closeRequestModal);
requestModal?.addEventListener("click",e=>{
    if(e.target===requestModal) closeRequestModal();
});

document.getElementById("requestForm")?.addEventListener("submit",async e=>{
    e.preventDefault();
    if(profile?.role!=="cr" || !activeSection || !auth.currentUser) return;

    const type=document.getElementById("requestType").value;
    const reason=document.getElementById("requestReason").value.trim();
    if(!reason) return alert("Please enter a reason.");

    const data={
        type,
        reason,
        section:activeSection,
        sectionLabel:sectionLabelOf(activeSection),
        requestedBy:String(auth.currentUser.email||"").trim(),
        status:"pending",
        requestedByRole:"cr",
        createdAt:new Date().toISOString()
    };

    if(type === "add_student"){
        data.qid=document.getElementById("requestQid").value.trim();
        data.name=document.getElementById("requestStudentName").value.trim().toUpperCase();
        if(!data.qid || !data.name) return alert("Enter Q.ID and student name.");
    }else if(type === "delete_student"){
        data.qid=document.getElementById("requestStudentSelect").value;
        const selected=students.find(s=>s.qid===data.qid);
        data.name=selected?.name || "";
        if(!data.qid) return alert("Please select a student.");
    }else if(type === "add_subject"){
        data.subject=document.getElementById("requestSubjectName").value.trim();
        if(!data.subject) return alert("Enter subject name.");
    }else if(type === "delete_subject"){
        data.subject=document.getElementById("requestSubjectSelect").value;
        if(!data.subject) return alert("Please select a subject.");
    }

    const submitButton=e.target.querySelector('button[type="submit"]');
    if(submitButton) submitButton.disabled=true;

    try{
        await addDoc(collection(db,"requests"),data);

        // Logging is intentionally non-blocking. A temporary log permission/network
        // issue must never make an already-saved request look like it failed.
        void writeActivityLog(
            "REQUEST_SUBMITTED",
            `${friendlyRequestType(type)} · ${data.qid||data.subject||""}`,
            {requestType:type,section:activeSection}
        );

        closeRequestModal();
        e.target.reset();
        syncRequestForm();
        await loadRequestCenter();
        alert("✅ Request sent to Admin.");
    }catch(err){
        console.error(err);
        alert(`Could not send request: ${err.message || err}. Publish the latest Firestore Rules first.`);
    }finally{
        if(submitButton) submitButton.disabled=false;
    }
});

function friendlyRequestType(type){
    return ({
        add_student:"Add Student",
        delete_student:"Delete Student",
        add_subject:"Add Subject",
        delete_subject:"Delete Subject",
        add_section:"Add New Section",
        contact_admin:"Contact Admin"
    })[type] || type;
}

async function findStudentByQid(section,qid){
    const snap=await getDocs(collection(db,"sections",section,"students"));
    return snap.docs.find(d=>String(d.data()?.qid||"").toLowerCase()===String(qid||"").toLowerCase());
}

async function findSubjectByName(section,name){
    const snap=await getDocs(collection(db,"sections",section,"subjects"));
    return snap.docs.find(d=>String(d.data()?.name||"").toLowerCase()===String(name||"").toLowerCase());
}

async function findExistingSectionForCourse(courseId, courseName, sectionLabel) {
    const snap = await getDocs(collection(db, "sections"));
    const wantedCourse = normalizeAcademicKey(courseId);
    const wantedName = normalizeAcademicKey(courseName);
    const wantedLabel = normalizeAcademicKey(sectionLabel);

    return snap.docs.find(sectionDoc => {
        const section = sectionDoc.data() || {};
        const label = normalizeAcademicKey(section.label || section.section || sectionDoc.id);
        if (label !== wantedLabel) return false;

        const sectionCourseKeys = [
            section.course,
            section.courseId,
            section.courseName,
            section.courseCode
        ].map(normalizeAcademicKey).filter(Boolean);

        return sectionCourseKeys.includes(wantedCourse) ||
            (wantedName && sectionCourseKeys.includes(wantedName));
    }) || null;
}

async function upsertRequestedStudents(sectionId, students) {
    const existingSnap = await getDocs(collection(db, "sections", sectionId, "students"));
    const byQid = new Map();
    existingSnap.docs.forEach(d => {
        const qid = normalizeAcademicKey(d.data()?.qid);
        if (qid) byQid.set(qid, d);
    });

    const batch = writeBatch(db);
    let writes = 0;
    for (const student of (students || [])) {
        const qid = String(student?.qid || "").trim();
        const name = String(student?.name || "").trim().toUpperCase();
        if (!qid || !name) continue;

        const existing = byQid.get(normalizeAcademicKey(qid));
        if (existing) {
            batch.set(existing.ref, { qid, name, updatedAt: new Date().toISOString() }, { merge: true });
        } else {
            batch.set(doc(collection(db, "sections", sectionId, "students")), { qid, name });
        }
        writes++;
    }
    if (writes) await batch.commit();
}

async function upsertRequestedSubjects(sectionId, subjects) {
    const existingSnap = await getDocs(collection(db, "sections", sectionId, "subjects"));
    const byName = new Map();
    existingSnap.docs.forEach(d => {
        const name = normalizeAcademicKey(d.data()?.name);
        if (name) byName.set(name, d);
    });

    const batch = writeBatch(db);
    let writes = 0;
    for (const rawSubject of (subjects || [])) {
        const name = String(rawSubject || "").trim();
        if (!name) continue;

        const existing = byName.get(normalizeAcademicKey(name));
        if (existing) {
            batch.set(existing.ref, { name, updatedAt: new Date().toISOString() }, { merge: true });
        } else {
            batch.set(doc(collection(db, "sections", sectionId, "subjects")), { name });
        }
        writes++;
    }
    if (writes) await batch.commit();
}

async function processChangeRequest(request,approve,options={}){
    const silent=options.silent===true;
    if(profile?.role!=="admin") return false;
    try{
        if(approve){
            if(request.type === "add_student"){
                if(!request.qid || !request.name) throw new Error("Student details are incomplete.");
                if(await findStudentByQid(request.section,request.qid)) throw new Error("A student with this Q.ID already exists.");
                await addDoc(collection(db,"sections",request.section,"students"),{qid:request.qid,name:String(request.name).toUpperCase()});
            }
            if(request.type === "delete_student"){const d=await findStudentByQid(request.section,request.qid);if(!d)throw new Error("Student not found.");await deleteDoc(d.ref);}
            if(request.type === "add_subject"){if(await findSubjectByName(request.section,request.subject))throw new Error("This subject already exists in the section.");await addDoc(collection(db,"sections",request.section,"subjects"),{name:request.subject});}
            if(request.type === "delete_subject"){const d=await findSubjectByName(request.section,request.subject);if(!d)throw new Error("Subject not found.");await deleteDoc(d.ref);}

            if(request.type === "add_section"){
                if(!request.sectionLabel && !request.section) throw new Error("Section name is missing.");

                const requestedCourse = String(request.course || "").trim();
                if(!requestedCourse) throw new Error("Course is missing.");

                // 1. Reuse the course created earlier if the request refers to
                // the same course name or code. Otherwise create it once.
                const courseSnap = await getDocs(collection(db, "courses"));
                const requestedCourseKey = normalizeAcademicKey(requestedCourse);
                let existingCourse = courseSnap.docs.find(courseDoc => {
                    const course = courseDoc.data() || {};
                    return [courseDoc.id, course.name, course.code]
                        .map(normalizeAcademicKey)
                        .filter(Boolean)
                        .includes(requestedCourseKey);
                });

                const courseId = existingCourse?.id ||
                    requestedCourse.toUpperCase().replace(/[^A-Z0-9]+/g, "-").replace(/^-|-$/g, "") ||
                    `COURSE-${Date.now()}`;

                if (!existingCourse) {
                    await setDoc(doc(db, "courses", courseId), {
                        name: requestedCourse,
                        code: courseId,
                        department: request.department || "",
                        duration: request.duration || "",
                        createdAt: new Date().toISOString(),
                        updatedAt: new Date().toISOString()
                    }, { merge: true });
                } else {
                    // Update the existing course without destroying fields that
                    // may already have been configured by Admin.
                    const existingData = existingCourse.data() || {};
                    await setDoc(doc(db, "courses", existingCourse.id), {
                        name: existingData.name || requestedCourse,
                        code: existingData.code || courseId,
                        ...(request.department ? { department: request.department } : {}),
                        ...(request.duration ? { duration: request.duration } : {}),
                        updatedAt: new Date().toISOString()
                    }, { merge: true });
                }

                const requestedLabel = String(request.sectionLabel || request.section || "").trim();
                const generatedSectionId = `${courseId}-${requestedLabel}`
                    .toUpperCase()
                    .replace(/[^A-Z0-9]+/g, "-")
                    .replace(/^-|-$/g, "");

                // 2. If this course already has this section, update that
                // document. If it does not, create the new section document.
                let existingSection = await findExistingSectionForCourse(courseId, requestedCourse, requestedLabel);
                if (!existingSection && request.section) {
                    const directRef = doc(db, "sections", request.section);
                    const directSnap = await getDoc(directRef);
                    if (directSnap.exists()) {
                        const directData = directSnap.data() || {};
                        const directCourse = normalizeAcademicKey(directData.course || directData.courseId || directData.courseName);
                        if (directCourse === normalizeAcademicKey(courseId) || directCourse === requestedCourseKey) {
                            existingSection = directSnap;
                        }
                    }
                }

                const sectionId = existingSection?.id || generatedSectionId || `SECTION-${Date.now()}`;
                const existingSectionData = existingSection?.data?.() || {};

                await setDoc(doc(db,"sections",sectionId),{
                    label: requestedLabel || existingSectionData.label || sectionId,
                    section: requestedLabel || existingSectionData.section || sectionId,
                    course: courseId,
                    courseName: requestedCourse,
                    semester: request.semester || existingSectionData.semester || "",
                    year: request.year || existingSectionData.year || "",
                    crQid: request.crQid || existingSectionData.crQid || "",
                    crName: request.crName || existingSectionData.crName || "",
                    crMobile: request.crMobile || existingSectionData.crMobile || "",
                    crEmail: request.crEmail || existingSectionData.crEmail || "",
                    mentorName: request.mentorName || existingSectionData.mentorName || "",
                    mentorMobile: request.mentorMobile || existingSectionData.mentorMobile || "",
                    updatedAt: new Date().toISOString(),
                    ...(existingSection ? {} : { createdAt: new Date().toISOString() })
                },{merge:true});

                // 3. Update existing students/subjects and add only the missing
                // ones. Approving the same section request twice can therefore
                // never create duplicate QIDs or duplicate subjects.
                await upsertRequestedStudents(sectionId, request.students || []);
                await upsertRequestedSubjects(sectionId, request.subjects || []);

                // 4. Provision/update the CR account for the final section ID.
                if (!request.crEmail) throw new Error("CR email is missing.");
                let tempPassword=request.generatedPassword||`QA@${Math.random().toString(36).slice(2,8)}${Math.floor(10+Math.random()*90)}`;
                try{await createUserWithEmailAndPassword(getProvisioningAuth(),request.crEmail,tempPassword);}
                catch(authError){if(authError.code!=="auth/email-already-in-use")throw authError;}
                finally{try{await signOut(getProvisioningAuth());}catch(_) {}}
                await setDoc(doc(db,"users",request.crEmail),{
                    role:"cr",email:request.crEmail,qid:request.crQid,name:request.crName,mobile:request.crMobile,
                    section:sectionId,sectionLabel:requestedLabel || sectionId,
                    course:requestedCourse,semester:request.semester||"",year:request.year||"",
                    mentorName:request.mentorName||"",mentorMobile:request.mentorMobile||"",updatedAt:new Date().toISOString()
                },{merge:true});

                if(!silent) {
                    const mailSubject=`QAttend - CR Account Assigned - ${requestedLabel||sectionId}`;
                    const mailBody=`Hello ${request.crName},\n\nYou have been assigned as the Class Representative (CR) for ${requestedLabel||sectionId}.\n\nLogin Email: ${request.crEmail}\nTemporary Password: ${tempPassword}\nCourse: ${requestedCourse}\nSemester: ${request.semester||""}\nYear: ${request.year||""}\n\nPlease change your password after your first login.`;
                    try {
                        await sendDirectEmail({
                            toEmail: request.crEmail,
                            subject: mailSubject,
                            message: mailBody,
                            replyTo: EMAILJS_CONFIG.adminEmail
                        });
                    } catch(emailError) {
                        console.error("CR invitation email error:", emailError);
                        alert(`CR account was created/updated, but the invitation email could not be sent. ${emailError.message || "Please try again."}`);
                    }
                }
            }
        }

        await updateDoc(doc(db,"requests",request.id),{status:approve?"approved":"rejected",reviewedBy:String(auth.currentUser?.email||"").toLowerCase(),reviewedAt:new Date().toISOString()});
        await writeActivityLog(approve?"REQUEST_APPROVED":"REQUEST_REJECTED",`${friendlyRequestType(request.type)} · ${request.qid||request.subject||request.sectionLabel||""}`,{requestId:request.id,section:request.section});
        if(approve){
            if(request.type==="add_student")await writeActivityLog("ADD_STUDENT",`${request.qid||""} · ${request.name||""}`,{section:request.section});
            if(request.type==="delete_student")await writeActivityLog("DELETE_STUDENT",`${request.qid||""} · ${request.name||""}`,{section:request.section});
            if(request.type==="add_subject")await writeActivityLog("ADD_SUBJECT",request.subject||"",{section:request.section});
            if(request.type==="delete_subject")await writeActivityLog("DELETE_SUBJECT",request.subject||"",{section:request.section});
            if(request.type==="add_section")await writeActivityLog("ADD_SECTION",`${request.sectionLabel||request.section} · ${request.course||""}`,{section:request.section});
            if(activeSection===request.section){
                if(request.type.includes("student")||request.type==="add_section")await loadStudents();
                if(request.type.includes("subject")||request.type==="add_section")await loadSubjects();
            }
        }
        if(!silent){await loadRequestCenter();await loadAdminDashboardData();await populateSectionDropdowns();await populateAttendanceCourseDropdown();alert(approve?"✅ Request approved and change applied.":"Request rejected. No data was changed.");}
        return true;
    }catch(e){console.error(e);if(!silent)alert(`Could not ${approve?"approve":"reject"} request: ${e.message||e}`);throw e;}
}


// Admin opens Mark Attendance by default. Dashboard is explicit from the hamburger menu.
// CR opens its assigned attendance section directly and cannot switch sections.

const courseModal = document.getElementById("courseModal");
const courseForm = document.getElementById("courseForm");
const courseTableBody = document.getElementById("coursesTableBody");
const sectionTableBody = document.getElementById("sectionsTableBody");
const courseFilter = document.getElementById("manageSectionsCourseFilter");
const adminSectionForm = document.getElementById("adminAddSectionForm");
const adminSectionFile = document.getElementById("adminSecExcelFile");
let adminSectionExcelPayload = null;

function openModalById(id) {
    const modal = document.getElementById(id);
    modal?.classList.remove("hidden");
    document.body.classList.add("modal-open");
}

function closeModalById(id) {
    document.getElementById(id)?.classList.add("hidden");
    if (!document.querySelector(".modal-backdrop:not(.hidden)")) document.body.classList.remove("modal-open");
}

async function loadCourses() {
    if (profile?.role !== "admin" || !courseTableBody) return;
    courseTableBody.innerHTML = `<tr><td colspan="6" class="empty-record">Loading courses...</td></tr>`;
    try {
        const snap = await getDocs(collection(db, "courses"));
        courseTableBody.innerHTML = "";
        if (!snap.size) {
            courseTableBody.innerHTML = `<tr><td colspan="6" class="empty-record">No courses found.</td></tr>`;
        }
        snap.docs.forEach((courseDoc, index) => {
        const course = courseDoc.data();
        const row = document.createElement("tr");
        const isBca = `${course.name || ""} ${course.code || ""} ${courseDoc.id}`.toLowerCase().includes("bca");
        row.innerHTML = `<td>${index + 1}</td><td>${escapeHtml(course.code || courseDoc.id)}</td><td>${escapeHtml(course.name || "")}</td><td>${escapeHtml(course.department || "")}</td><td>${escapeHtml(course.duration || "")}</td><td><button type="button" class="text-btn" data-edit-course="${escapeHtml(courseDoc.id)}">Edit</button> <button type="button" class="text-btn danger-text" data-delete-course="${escapeHtml(courseDoc.id)}">Delete</button>${isBca ? ` <button type="button" class="text-btn danger-text" data-clear-course-students="${escapeHtml(courseDoc.id)}">Remove Students</button>` : ""}</td>`;
        row.querySelector("[data-edit-course]")?.addEventListener("click", () => {
            document.getElementById("courseEditId").value = courseDoc.id;
            document.getElementById("courseNameInput").value = course.name || "";
            document.getElementById("courseCodeInput").value = course.code || courseDoc.id;
            document.getElementById("courseDeptInput").value = course.department || "";
            document.getElementById("courseDurationInput").value = course.duration || "";
            document.getElementById("courseModalTitle").textContent = "Edit Course";
            document.getElementById("saveCourseBtn").textContent = "Save Changes";
            openModalById("courseModal");
        });
        row.querySelector("[data-delete-course]")?.addEventListener("click", async () => {
            if (!confirm(`Delete ${course.name || courseDoc.id}?`)) return;
            try {
                await deleteCourseWithSections(courseDoc.id);
                await loadCourses();
                await populateCourseOptions();
                await populateAttendanceCourseDropdown();
            } catch (error) {
                console.error("Could not delete course:", error);
                alert("Could not delete the course and its sections.");
            }
        });
        row.querySelector("[data-clear-course-students]")?.addEventListener("click", async () => {
            if (!confirm("Remove all students from every BCA section? Subjects, sections, and the course will remain.")) return;
            try {
                await removeCourseStudents(courseDoc.id);
                await loadAdminDashboardData();
                alert("BCA students removed successfully.");
            } catch (error) {
                console.error("Could not remove BCA students:", error);
                alert("Could not remove BCA students.");
            }
        });
        courseTableBody.appendChild(row);
        });
        await loadCourseSectionRequests();
    } catch (error) {
        console.error("Could not load courses:", error);
        courseTableBody.innerHTML = `<tr><td colspan="6" class="empty-record">Could not load courses. Check Firestore rules and try again.</td></tr>`;
    }
}

async function removeCourseStudents(courseId) {
    const courseSnap = await getDoc(doc(db, "courses", courseId));
    const course = courseSnap.exists() ? courseSnap.data() : {};
    const courseKeys = new Set([
        courseId.toLowerCase(),
        String(course.name || "").toLowerCase(),
        String(course.code || "").toLowerCase()
    ].filter(Boolean));
    const sectionsSnap = await getDocs(collection(db, "sections"));
    const matchingSections = sectionsSnap.docs.filter(sectionDoc => {
        const section = sectionDoc.data() || {};
        const sectionText = `${sectionDoc.id} ${section.label || ""} ${section.section || ""}`.toLowerCase();
        return courseKeys.has(String(section.course || "").toLowerCase())
            || courseKeys.has(String(section.courseName || "").toLowerCase())
            || (courseKeys.has("bca") && sectionText.includes("bca"));
    });
    for (const sectionDoc of matchingSections) {
        const studentsSnap = await getDocs(collection(db, "sections", sectionDoc.id, "students"));
        for (const studentDoc of studentsSnap.docs) await deleteDoc(studentDoc.ref);
    }
}

async function deleteCourseWithSections(courseId) {
    const sectionsSnap = await getDocs(collection(db, "sections"));
    const matchingSections = sectionsSnap.docs.filter(sectionDoc => sectionDoc.data()?.course === courseId);
    for (const sectionDoc of matchingSections) {
        for (const subcollection of ["students", "subjects", "attendance"]) {
            const subSnap = await getDocs(collection(db, "sections", sectionDoc.id, subcollection));
            for (const item of subSnap.docs) await deleteDoc(item.ref);
        }
        const usersSnap = await getDocs(collection(db, "users"));
        for (const userDoc of usersSnap.docs) {
            if (userDoc.data()?.section === sectionDoc.id) await deleteDoc(userDoc.ref);
        }
        await deleteDoc(sectionDoc.ref);
    }
    await deleteDoc(doc(db, "courses", courseId));
}

async function loadCourseSectionRequests() {
    const container = document.getElementById("courseRequestList");
    if (!container || profile?.role !== "admin") return;
    container.innerHTML = `<p class="empty-record">Loading section requests...</p>`;
    try {
        const snap = await getDocs(collection(db, "requests"));
        const requests = snap.docs
            .map(requestDoc => ({ id: requestDoc.id, ...requestDoc.data() }))
            .filter(request => request.type === "add_section" && request.status === "pending")
            .sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));

        container.innerHTML = "";
        if (!requests.length) {
            container.innerHTML = `<p class="empty-record">No pending section requests.</p>`;
            return;
        }

        requests.forEach(request => {
            const card = document.createElement("div");
            card.className = "request-card";
            card.innerHTML = `<div class="request-top"><div class="request-main"><b>${escapeHtml(request.course || "New Course")} · ${escapeHtml(request.sectionLabel || request.section || "New Section")}</b><small>CR: ${escapeHtml(request.crName || request.crEmail || "")} · ${escapeHtml(request.crMobile || "")}</small><small>${request.students?.length || 0} students · ${request.subjects?.length || 0} subjects</small></div><span class="status-pill status-pending">PENDING</span></div>`;
            const actions = document.createElement("div");
            actions.className = "request-actions";
            const viewButton = document.createElement("button");
            viewButton.className = "text-btn";
            viewButton.textContent = "View";
            viewButton.addEventListener("click", () => showRequestDetails(request));
            const approveButton = document.createElement("button");
            approveButton.className = "approve-btn";
            approveButton.textContent = "Approve & Add";
            approveButton.addEventListener("click", async () => {
                await processChangeRequest(request, true);
                await loadCourses();
            });
            const deleteButton = document.createElement("button");
            deleteButton.className = "text-btn danger-text";
            deleteButton.textContent = "Delete";
            deleteButton.addEventListener("click", () => deleteRequest(request));
            const excelButton = document.createElement("button");
            excelButton.className = "text-btn";
            excelButton.textContent = "View Excel";
            excelButton.addEventListener("click", () => downloadRequestExcel(request));
            actions.append(viewButton, excelButton, approveButton, deleteButton);
            card.appendChild(actions);
            container.appendChild(card);
        });
    } catch (error) {
        console.error("Could not load section requests for courses:", error);
        container.innerHTML = `<p class="empty-record">Could not load section requests.</p>`;
    }
}

async function populateCourseOptions() {
    let courses = [];
    try {
        const snap = await getDocs(collection(db, "courses"));
        courses = snap.docs;
    } catch (error) {
        console.warn("Could not load course options:", error);
    }

    [courseFilter, document.getElementById("adminSecCourse"), manageCourseSelect].forEach(select => {
        if (!select) return;
        const selected = select.value;
        select.innerHTML = `<option value="">-- Select Course --</option>`;
        courses
            .sort((a, b) => String(a.data()?.name || a.id).localeCompare(String(b.data()?.name || b.id), undefined, { sensitivity: "base" }))
            .forEach(courseDoc => {
                const course = courseDoc.data() || {};
                const option = document.createElement("option");
                option.value = courseDoc.id;
                option.textContent = `${course.name || courseDoc.id}${course.code ? ` (${course.code})` : ""}`;
                select.appendChild(option);
            });
        if (selected) select.value = selected;
    });
    if (manageCourseSelect?.value) filterManageSections(manageCourseSelect.value);
}

async function loadSectionsForCourse(courseId = courseFilter?.value) {
    if (!sectionTableBody) return;
    sectionTableBody.innerHTML = "";
    if (!courseId) {
        sectionTableBody.innerHTML = `<tr><td colspan="7" class="empty-record">Select a course to view sections.</td></tr>`;
        return;
    }

    try {
        const { courses, sections } = await getAcademicData();
        const selectedCourse = courses.find(d => String(d.id) === String(courseId));
        const selectedData = selectedCourse?.data() || {};
        const selectedKeys = new Set([
            normalizeAcademicKey(courseId),
            normalizeAcademicKey(selectedData.name),
            normalizeAcademicKey(selectedData.code)
        ].filter(Boolean));

        const rows = sections.filter(sectionDoc => {
            const section = sectionDoc.data() || {};
            const resolvedCourse = resolveCourseIdForSection(section, sectionDoc.id, courses);
            if (normalizeAcademicKey(resolvedCourse) === normalizeAcademicKey(courseId)) return true;
            const sectionKeys = [section.course, section.courseId, section.courseName, section.courseCode]
                .map(normalizeAcademicKey).filter(Boolean);
            return sectionKeys.some(key => selectedKeys.has(key));
        });

        if (!rows.length) {
            sectionTableBody.innerHTML = `<tr><td colspan="7" class="empty-record">No sections found for this course.</td></tr>`;
            return;
        }

        for (const [index, sectionDoc] of rows.entries()) {
            const section = sectionDoc.data() || {};
            const [studentsSnap, subjectsSnap] = await Promise.all([
                getDocs(collection(db, "sections", sectionDoc.id, "students")),
                getDocs(collection(db, "sections", sectionDoc.id, "subjects"))
            ]);
            const row = document.createElement("tr");
            row.innerHTML = `<td>${index + 1}</td><td>${escapeHtml(section.label || section.section || sectionDoc.id)}</td><td>${escapeHtml(section.semester || "")}</td><td>${escapeHtml(section.year || "")}</td><td>${studentsSnap.size}</td><td>${subjectsSnap.size}</td><td><button type="button" class="text-btn" data-open-section="${escapeHtml(sectionDoc.id)}">Open</button> <button type="button" class="text-btn danger-text" data-delete-section="${escapeHtml(sectionDoc.id)}">Delete</button></td>`;
            row.querySelector("[data-open-section]")?.addEventListener("click", async () => {
                managementSection = sectionDoc.id;
                await populateSectionDropdowns();
                if (manageCourseSelect) manageCourseSelect.value = courseId;
                filterManageSections(courseId);
                if (manageSectionSelect) manageSectionSelect.value = sectionDoc.id;
                await refreshManageData();
            });
            row.querySelector("[data-delete-section]")?.addEventListener("click", async () => {
                if (!confirm(`Delete ${section.label || sectionDoc.id}?`)) return;
                await deleteDoc(sectionDoc.ref);
                if (managementSection === sectionDoc.id) managementSection = null;
                await populateSectionDropdowns();
                await loadSectionsForCourse(courseId);
            });
            sectionTableBody.appendChild(row);
        }
    } catch (error) {
        console.error("Could not load sections for course:", error);
        sectionTableBody.innerHTML = `<tr><td colspan="7" class="empty-record">Could not load sections. Check Firestore rules and try again.</td></tr>`;
    }
}

document.querySelectorAll(".sidebar-btn").forEach(button => button.addEventListener("click", async () => {
    const feature = button.dataset.feature;
    if (feature === "dashboard") showAdminDashboard();
    if (feature === "courses") { hideMainViews(); document.getElementById("adminCoursesView")?.classList.remove("hidden"); await loadCourses(); }
    if (feature === "sections") { hideMainViews(); document.getElementById("adminSectionsView")?.classList.remove("hidden"); await populateCourseOptions(); await loadSectionsForCourse(); }
    if (feature === "manage") showManageData();
    if (feature === "manage-subjects") showManageData();
    if (feature === "requests") showRequestsPanel();
    if (feature === "attendance") showAttendanceView();
    if (feature === "logs") { showAdminDashboard(); await loadAdminActivityLogs(); }
    if (feature === "logout") await logoutFromApp();
}));

document.getElementById("openAddCourseBtn")?.addEventListener("click", () => {
    courseForm?.reset();
    document.getElementById("courseEditId").value = "";
    document.getElementById("courseModalTitle").textContent = "Add New Course";
    document.getElementById("saveCourseBtn").textContent = "Create Course";
    openModalById("courseModal");
});
document.getElementById("refreshCourseRequestsBtn")?.addEventListener("click", loadCourseSectionRequests);
document.getElementById("closeCourseModal")?.addEventListener("click", () => closeModalById("courseModal"));
document.getElementById("cancelCourseBtn")?.addEventListener("click", () => closeModalById("courseModal"));
courseForm?.addEventListener("submit", async event => {
    event.preventDefault();
    const message = document.getElementById("courseFormMessage");
    const saveButton = document.getElementById("saveCourseBtn");
    const editId = document.getElementById("courseEditId")?.value.trim();
    const name = document.getElementById("courseNameInput")?.value.trim();
    const code = document.getElementById("courseCodeInput")?.value.trim().toUpperCase();
    const department = document.getElementById("courseDeptInput")?.value.trim();
    const duration = document.getElementById("courseDurationInput")?.value.trim();
    const id = editId || code?.replace(/[^A-Z0-9]+/g, "-");

    if (!name || !code || !department || !duration || !id) {
        if (message) message.textContent = "Please fill all course fields.";
        return;
    }

    if (saveButton) saveButton.disabled = true;
    if (message) message.textContent = "Saving course...";
    try {
        await setDoc(doc(db, "courses", id), { name, code, department, duration, updatedAt: new Date().toISOString() }, { merge: true });
        closeModalById("courseModal");
        await loadCourses();
        await populateCourseOptions();
        await populateAttendanceCourseDropdown();
    } catch (error) {
        console.error("Could not save course:", error);
        if (message) message.textContent = "Could not save course. Check Firestore rules and try again.";
    } finally {
        if (saveButton) saveButton.disabled = false;
    }
});

courseFilter?.addEventListener("change", () => loadSectionsForCourse(courseFilter.value));
document.getElementById("openAddSectionBtn")?.addEventListener("click", async () => {
    await populateCourseOptions();
    adminSectionExcelPayload = null;
    adminSectionForm?.reset();
    document.getElementById("adminSecExcelPreview")?.classList.add("hidden");
    openModalById("adminAddSectionModal");
});
document.getElementById("closeAdminAddSectionModal")?.addEventListener("click", () => closeModalById("adminAddSectionModal"));
document.getElementById("cancelAdminAddSectionBtn")?.addEventListener("click", () => closeModalById("adminAddSectionModal"));
adminSectionFile?.addEventListener("change", async () => {
    const file = adminSectionFile.files?.[0];
    if (!file) return;
    adminSectionExcelPayload = await parseSectionExcelFile(file);
    renderSectionExcelPreview("adminSec", adminSectionExcelPayload);
    document.getElementById("adminSecMessage").textContent = "Excel preview ready.";
});
adminSectionForm?.addEventListener("submit", async event => {
    event.preventDefault();
    const courseId = document.getElementById("adminSecCourse").value;
    const sectionLabelValue = document.getElementById("adminSecName").value.trim();
    if (!courseId || !sectionLabelValue || !adminSectionExcelPayload) return;
    const courseSnap = await getDocs(collection(db, "courses"));
    const selectedCourseDoc = courseSnap.docs.find(d => String(d.id) === String(courseId));
    const selectedCourse = selectedCourseDoc?.data() || {};
    const existingSection = await findExistingSectionForCourse(courseId, selectedCourse.name || courseId, sectionLabelValue);
    const sectionId = existingSection?.id || `${courseId}-${sectionLabelValue}`.toUpperCase().replace(/[^A-Z0-9]+/g, "-").replace(/^-|-$/g, "");
    await setDoc(doc(db, "sections", sectionId), {
        course: courseId,
        courseName: selectedCourse.name || courseId,
        label: sectionLabelValue,
        section: sectionLabelValue,
        semester: document.getElementById("adminSecSemester").value.trim(),
        year: document.getElementById("adminSecYear").value.trim(),
        mentorName: document.getElementById("adminSecMentorName").value.trim(),
        mentorMobile: document.getElementById("adminSecMentorMobile").value.trim(),
        updatedAt: new Date().toISOString(),
        ...(existingSection ? {} : { createdAt: new Date().toISOString() })
    }, { merge: true });
    await upsertRequestedStudents(sectionId, adminSectionExcelPayload.students);
    await upsertRequestedSubjects(sectionId, adminSectionExcelPayload.subjects);
    closeModalById("adminAddSectionModal");
    await populateSectionDropdowns();
    await loadSectionsForCourse(courseId);
});

