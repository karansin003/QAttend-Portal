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
    initializeFirestore,
    persistentLocalCache,
    persistentMultipleTabManager,
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

// Offline-first Firestore:
// - IndexedDB keeps actively used Firestore data available after reload/offline.
// - Local writes are queued by Firestore and synchronized automatically when the
//   network returns.
// - Multi-tab persistence lets admin/CR tabs share the same local cache.
const db = initializeFirestore(app, {
    localCache: persistentLocalCache({
        tabManager: persistentMultipleTabManager()
    })
});

const authPersistenceReady = setPersistence(auth, browserLocalPersistence).catch(error => {
    console.warn("Could not enable local login persistence:", error);
});

// ------------------------------------------------------------
// OFFLINE-FIRST APP STATUS
// ------------------------------------------------------------
let qattendNetworkStatusEl = null;

function ensureNetworkStatusIndicator() {
    if (qattendNetworkStatusEl) return qattendNetworkStatusEl;

    qattendNetworkStatusEl = document.createElement("div");
    qattendNetworkStatusEl.id = "qattendNetworkStatus";
    qattendNetworkStatusEl.className = "qattend-network-status";
    qattendNetworkStatusEl.setAttribute("role", "status");
    qattendNetworkStatusEl.setAttribute("aria-live", "polite");

    // Keep the existing UI untouched: the indicator is injected only when needed.
    document.body.appendChild(qattendNetworkStatusEl);
    return qattendNetworkStatusEl;
}

function updateNetworkStatusIndicator() {
    const el = ensureNetworkStatusIndicator();
    const online = navigator.onLine;

    el.classList.toggle("offline", !online);
    el.classList.toggle("online", online);

    if (online) {
        el.textContent = "● Online";
        el.title = "QAttend is online. Firestore will synchronize pending offline changes.";
    } else {
        el.textContent = "● Offline — changes will sync when online";
        el.title = "QAttend is using cached data. Firestore writes will synchronize when the network returns.";
    }
}

window.addEventListener("online", () => {
    updateNetworkStatusIndicator();
    // Give the browser a moment to restore connectivity before refreshing UI data.
    setTimeout(() => {
        window.dispatchEvent(new CustomEvent("qattend:network-restored"));
    }, 300);
});

window.addEventListener("offline", updateNetworkStatusIndicator);

function registerQAttendServiceWorker() {
    if (!("serviceWorker" in navigator)) return;

    navigator.serviceWorker.register("./service-worker.js", { scope: "./" })
        .then(registration => {
            console.info("QAttend service worker registered:", registration.scope);

            // Pick up an updated worker without requiring a manual refresh.
            registration.addEventListener("updatefound", () => {
                const worker = registration.installing;
                if (!worker) return;

                worker.addEventListener("statechange", () => {
                    if (worker.state === "installed" && navigator.serviceWorker.controller) {
                        worker.postMessage({ type: "SKIP_WAITING" });
                    }
                });
            });
        })
        .catch(error => {
            // Offline app functionality must never be blocked if SW registration fails.
            console.warn("QAttend service worker registration failed:", error);
        });
}

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
        updateNetworkStatusIndicator();
        registerQAttendServiceWorker();
    }, { once: true });
} else {
    updateNetworkStatusIndicator();
    registerQAttendServiceWorker();
}



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
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const normalize = v => String(v ?? "").trim();
    const findSheet = name => {
        const target = name.replace(/\s+/g, "").toLowerCase();
        const sheetName = workbook.SheetNames.find(candidate => candidate.replace(/\s+/g, "").toLowerCase() === target);
        return sheetName ? workbook.Sheets[sheetName] : null;
    };

    const studentsSheet = findSheet("Students") || workbook.Sheets[workbook.SheetNames[0]];
    const subjectsSheet = findSheet("Subjects") || (workbook.SheetNames.length > 1 ? workbook.Sheets[workbook.SheetNames[1]] : null);

    const studentRows = studentsSheet ? XLSX.utils.sheet_to_json(studentsSheet, { header: 1, defval: "" }) : [];
    const subjectRows = subjectsSheet ? XLSX.utils.sheet_to_json(subjectsSheet, { header: 1, defval: "" }) : [];

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
        if (qids.has(qid.toLowerCase())) {
            duplicateCount++;
            return;
        }
        qids.add(qid.toLowerCase());
        studentsOut.push({ qid, name: name.toUpperCase() });
    });

    const subjectsOut = [];
    const subjectSet = new Set();
    subjectRows.slice(1).forEach(row => {
        const name = normalize(row[0]);
        if (!name) return;
        const key = name.toLowerCase();
        if (!subjectSet.has(key)) {
            subjectSet.add(key);
            subjectsOut.push(name);
        }
    });

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
            emailInput.value.trim().toLowerCase();

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
            if (loginSubmitBtn) loginSubmitBtn.disabled = false;
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

            const userEmailKey = String(user.email || "").toLowerCase();
            let profileSnap = null;

            try {
                profileSnap = await getDoc(
                    doc(db, "users", userEmailKey)
                );
            } catch (profileError) {
                // When offline, use the last verified profile cached after a
                // successful login. Do NOT treat an arbitrary localStorage
                // value as authentication; Firebase Auth is still the source
                // of the authenticated user.
                if (cachedProfile) {
                    console.warn("Using cached QAttend profile while offline:", profileError);
                    profile = cachedProfile;
                } else {
                    throw profileError;
                }
            }

            if (!profile) {
                if (!profileSnap?.exists()) {
                    alert(
                        "Your account is not set up yet. Ask the Admin to assign you a role."
                    );

                    await signOut(auth);
                    return;
                }

                profile = profileSnap.data();
            }

            try {
                localStorage.setItem("qattend-profile", JSON.stringify({
                    email: userEmailKey,
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

            // A previously authenticated user can continue using the app
            // offline when Firebase Auth restored the session but Firestore
            // cannot reach the backend and no cached profile was available.
            // In that case, fail safely instead of logging them out.
            if (!navigator.onLine) {
                let fallbackProfile = null;
                try {
                    const cached = JSON.parse(localStorage.getItem("qattend-profile") || "null");
                    if (
                        cached?.email === String(user.email || "").toLowerCase() &&
                        cached.profile?.role
                    ) {
                        fallbackProfile = cached.profile;
                    }
                } catch (_) {}

                if (fallbackProfile) {
                    profile = fallbackProfile;
                    console.warn("Continuing with cached QAttend profile while offline.");
                } else {
                    alert("You are offline and this account has not been cached on this device yet. Connect to the internet once, then reopen QAttend.");
                    return;
                }
            } else {
                alert("Could not load your account role. Try logging in again.");
                await signOut(auth);
                return;
            }
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

async function populateSectionDropdowns() {
    const allSections = [...SECTIONS];
    sectionCourseMap = Object.fromEntries(SECTIONS.map(section => [section.id, "BTECH"]));
    try {
        const snap = await getDocs(collection(db,"sections"));
        snap.docs.forEach(d=>{
            const x=d.data()||{};
            if(!d.id) return;
            sectionCourseMap[d.id] = x.course || x.courseId || sectionCourseMap[d.id] || "BTECH";
            if(allSections.some(s=>s.id===d.id)) return;
            const item={id:d.id,label:x.label||x.section||d.id};
            allSections.push(item);
            SECTIONS.push(item);
        });
    }catch(e){console.warn("Could not load dynamic sections:",e);}

    if (sectionSelect) {
        sectionSelect.innerHTML=`<option value="">-- Select Section --</option>`;
        sectionSelect.disabled = true;
    }
    if(newCrSection)newCrSection.innerHTML="";
    const manageSection=document.getElementById("manageSectionSelect");
    if(manageSection) {
        manageSection.innerHTML=`<option value="">-- Select Section --</option>`;
        manageSection.disabled = !manageCourseSelect?.value;
    }

    allSections.forEach(section=>{
        if (sectionSelect) { const option1=document.createElement("option");option1.value=section.id;option1.textContent=section.label;sectionSelect.appendChild(option1); }
        if(newCrSection){const option2=document.createElement("option");option2.value=section.id;option2.textContent=section.label;newCrSection.appendChild(option2);}
        if(manageSection){const option3=document.createElement("option");option3.value=section.id;option3.textContent=section.label;manageSection.appendChild(option3);}
    });
    if (manageCourseSelect?.value) filterManageSections(manageCourseSelect.value);
}

async function populateAttendanceCourseDropdown() {
    if (!attendanceCourseSelect) return;
    attendanceCourseSelect.innerHTML = `<option value="">-- Select Course --</option>`;
    const visibleCourses = new Set(["b.tech"]);
    const btechOption = document.createElement("option");
    btechOption.value = "BTECH";
    btechOption.textContent = "B.Tech";
    attendanceCourseSelect.appendChild(btechOption);

    try {
        const snap = await getDocs(collection(db, "courses"));
        snap.docs.forEach(courseDoc => {
            const course = courseDoc.data() || {};
            const courseName = String(course.name || course.code || courseDoc.id).trim();
            const courseKey = courseName.toLowerCase() === "b.tech ai & ml" ? "b.tech" : courseName.toLowerCase();
            if (visibleCourses.has(courseKey)) return;
            visibleCourses.add(courseKey);
            const option = document.createElement("option");
            option.value = courseDoc.id;
            option.textContent = courseKey === "b.tech" ? "B.Tech" : courseName;
            attendanceCourseSelect.appendChild(option);
        });
    } catch (error) {
        console.warn("Could not load attendance courses:", error);
    }
}

async function ensureLegacyBtechCourse() {
    if (profile?.role !== "admin") return;
    const coursesSnap = await getDocs(collection(db, "courses"));
    const existingCourse = coursesSnap.docs.find(courseDoc => {
        const course = courseDoc.data() || {};
        const name = String(course.name || "").trim().toLowerCase();
        const code = String(course.code || "").trim().toLowerCase();
        return name === "b.tech" || name === "b.tech ai & ml" || code === "btech" || code === "btech-ai-ml";
    });
    const courseId = existingCourse?.id || "BTECH";
    const courseRef = doc(db, "courses", courseId);
    await setDoc(courseRef, {
        name: "B.Tech",
        code: "BTECH",
        department: existingCourse?.data()?.department || "Computer Science",
        duration: existingCourse?.data()?.duration || "4 Years",
        updatedAt: new Date().toISOString()
    }, { merge: true });

    const sectionWrites = SECTIONS.map(section => setDoc(
        doc(db, "sections", section.id),
        {
            label: section.label,
            course: courseId,
            courseName: "B.Tech",
            updatedAt: new Date().toISOString()
        },
        { merge: true }
    ));
    await Promise.all(sectionWrites);

    const sectionsSnap = await getDocs(collection(db, "sections"));
    await Promise.all(sectionsSnap.docs.map(sectionDoc => {
        return setDoc(sectionDoc.ref, {
            course: "BTECH",
            courseName: "B.Tech",
            updatedAt: new Date().toISOString()
        }, { merge: true });
    }));
}

function filterAttendanceSections(courseId) {
    if (!sectionSelect) return;
    sectionSelect.disabled = !courseId;
    const selectedSection = sectionSelect.value;
    [...sectionSelect.options].forEach(option => {
        if (!option.value) {
            option.hidden = false;
            return;
        }
        const mappedCourse = String(sectionCourseMap[option.value] || "").toUpperCase();
        option.hidden = !courseId || Boolean(mappedCourse && mappedCourse !== String(courseId).toUpperCase());
    });
    if (selectedSection && !sectionSelect.selectedOptions[0]?.hidden) return;
    sectionSelect.value = "";
    activeSection = null;
    attendanceCourseLabel.textContent = courseId
        ? attendanceCourseSelect.selectedOptions[0]?.textContent || "--"
        : "--";
    workArea?.classList.add("hidden");
    noSectionMessage?.classList.remove("hidden");
}

attendanceCourseSelect?.addEventListener("change", () => {
    const courseId = attendanceCourseSelect.value;
    filterAttendanceSections(courseId);
});


// SECTION CHANGE

sectionSelect?.addEventListener(
    "change",
    async function () {

        activeSection =
            sectionSelect.value ||
            null;


        if (!activeSection) {

            workArea.classList.add(
                "hidden"
            );

            noSectionMessage.classList.remove(
                "hidden"
            );

            sectionLabel.textContent =
                "-- Not selected --";
            if (attendanceCourseLabel) attendanceCourseLabel.textContent = "--";

            headerTotalStudents.textContent =
                "0";

            totalStudents.textContent =
                "0";

            presentStudents.textContent =
                "0";

            absentStudents.textContent =
                "0";

            studentList.innerHTML =
                "";

            students = [];

            return;
        }


        workArea.classList.remove(
            "hidden"
        );


        noSectionMessage.classList.add(
            "hidden"
        );


        sectionLabel.textContent =
            sectionLabelOf(
                activeSection
            );

        if (attendanceCourseLabel && attendanceCourseSelect) {
            const courseId = sectionCourseMap[activeSection];
            if (courseId) attendanceCourseSelect.value = courseId;
            attendanceCourseLabel.textContent = attendanceCourseSelect.selectedOptions[0]?.textContent || sectionCourseMap[activeSection] || "--";
        }


        subjectSelect.value =
            "";


        attendanceDate.value =
            new Date()
                .toISOString()
                .split("T")[0];


        await loadStudents();

        await loadSubjects();

        await loadRecords();
        syncAttendanceGate();
    }
);


// STUDENTS - LOAD

async function loadStudents() {

    try {

        const snapshot =
            await getDocs(
                collection(
                    db,
                    "sections",
                    activeSection,
                    "students"
                )
            );


        students =
            snapshot.docs

                .map(
                    function (docSnap) {

                        return {

                            id:
                                docSnap.id,

                            qid:
                                docSnap.data()
                                    .qid,

                            name:
                                docSnap.data()
                                    .name,

                            status:
                                null
                        };
                    }
                )

                .sort(
                    function (a, b) {

                        return compareNames(
                            a.name,
                            b.name
                        );
                    }
                );


        displayStudents();


    } catch (error) {

        console.error(error);


        studentList.innerHTML =
            `<p class="empty-record">
                Error loading students.
            </p>`;
    }
}


// STUDENTS - DISPLAY

function displayStudents() {

    const search =
        searchStudent.value
            .trim()
            .toLowerCase();


    const isAdmin =
        profile.role ===
        "admin";


    studentList.innerHTML =
        "";


    students.forEach(
        function (student) {

            const matchesName =
                student.name
                    .toLowerCase()
                    .includes(search);


            const matchesQid =
                student.qid
                    .toLowerCase()
                    .includes(search);


            if (
                !matchesName &&
                !matchesQid
            ) {
                return;
            }


            const row =
                document.createElement(
                    "div"
                );


            row.className =
                "student";


            row.dataset.id =
                student.id;


            row.innerHTML = `

                <div class="qid-number">
                    ${student.qid}
                </div>


                <div class="student-name">
                    ${student.name}
                </div>


                <div class="buttons">

                    <button
                        type="button"
                        class="absent ${
                            student.status === "Absent"
                                ? "active-absent"
                                : ""
                        }"
                        data-id="${student.id}"
                        data-status="Absent">

                        ✕ Absent

                    </button>


                    <button
                        type="button"
                        class="present ${
                            student.status === "Present"
                                ? "active-present"
                                : ""
                        }"
                        data-id="${student.id}"
                        data-status="Present">

                        ✓ Present

                    </button>

                </div>


                ${
                    isAdmin
                        ? `

                    <div class="student-menu-wrapper">

                        <button
                            type="button"
                            class="student-menu-btn"
                            data-menu-id="${student.id}"
                            aria-label="Student options"
                            title="Student options">

                            ⋮

                        </button>


                        <div
                            class="student-menu"
                            data-menu="${student.id}">

                            <button
                                type="button"
                                class="edit-student-btn"
                                data-id="${student.id}">

                                ✏️ Edit

                            </button>


                            <button
                                type="button"
                                class="delete-student-menu-btn"
                                data-id="${student.id}">

                                🗑️ Delete

                            </button>

                        </div>

                    </div>

                `
                        : ""
                }

            `;


            studentList.appendChild(
                row
            );
        }
    );


    updateStats();
}


// STUDENT CLICK HANDLER

studentList?.addEventListener(
    "click",
    function (event) {


        // THREE-DOT MENU

        const menuButton =
            event.target.closest(
                ".student-menu-btn"
            );


        if (menuButton) {

            if (
                !profile ||
                profile.role !==
                    "admin"
            ) {

                return;
            }


            const studentId =
                menuButton.dataset.menuId;


            const menu =
                document.querySelector(
                    `.student-menu[data-menu="${studentId}"]`
                );


            document
                .querySelectorAll(
                    ".student-menu.show"
                )
                .forEach(
                    function (item) {

                        if (
                            item !==
                            menu
                        ) {

                            item.classList.remove(
                                "show"
                            );
                        }
                    }
                );


            if (menu) {

                menu.classList.toggle(
                    "show"
                );
            }


            return;
        }


        // PRESENT / ABSENT

        const statusButton =
            event.target.closest(
                "[data-status]"
            );


        if (statusButton) {

            if (
                !subjectSelect.value
            ) {

                alert(
                    "Please select a subject first."
                );

                return;
            }


            const student =
                students.find(
                    item =>
                        item.id ===
                        statusButton.dataset.id
                );


            if (!student) {
                return;
            }


            student.status =
                statusButton.dataset.status;


            const row =
                statusButton.closest(
                    ".student"
                );


            const absentBtn =
                row.querySelector(
                    ".absent"
                );


            const presentBtn =
                row.querySelector(
                    ".present"
                );


            absentBtn.classList.toggle(
                "active-absent",
                student.status ===
                    "Absent"
            );


            presentBtn.classList.toggle(
                "active-present",
                student.status ===
                    "Present"
            );


            updateStats();
            saveAttendanceDraft();


            return;
        }


        // EDIT

        const editButton =
            event.target.closest(
                ".edit-student-btn"
            );


        if (editButton) {

            if (
                !profile ||
                profile.role !==
                    "admin"
            ) {

                return;
            }


            document
                .querySelectorAll(
                    ".student-menu.show"
                )
                .forEach(
                    function (menu) {

                        menu.classList.remove(
                            "show"
                        );
                    }
                );


            editStudent(
                editButton.dataset.id
            );


            return;
        }


        // DELETE

        const deleteButton =
            event.target.closest(
                ".delete-student-menu-btn"
            );


        if (deleteButton) {

            if (
                !profile ||
                profile.role !==
                    "admin"
            ) {

                return;
            }


            document
                .querySelectorAll(
                    ".student-menu.show"
                )
                .forEach(
                    function (menu) {

                        menu.classList.remove(
                            "show"
                        );
                    }
                );


            deleteStudent(
                deleteButton.dataset.id
            );
        }
    }
);


// CLOSE THREE-DOT MENU

document.addEventListener(
    "click",
    function (event) {

        if (
            !event.target.closest(
                ".student-menu-wrapper"
            )
        ) {

            document
                .querySelectorAll(
                    ".student-menu.show"
                )
                .forEach(
                    function (menu) {

                        menu.classList.remove(
                            "show"
                        );
                    }
                );
        }
    }
);


// ADD STUDENT

addStudentForm?.addEventListener(
    "submit",
    async function (event) {

        event.preventDefault();


        const adminPanelOpen = !document.getElementById("adminPanel")?.classList.contains("hidden");
        const targetSection = adminPanelOpen
            ? document.getElementById("manageSectionSelect")?.value || null
            : activeSection;

        if (!targetSection) {
            alert("Please select a section first.");
            return;
        }


        const qid =
            newStudentQid.value.trim();


        const name =
            newStudentName.value
                .trim()
                .toUpperCase();


        if (
            !qid ||
            !name
        ) {
            return;
        }


        try {

            await addDoc(

                collection(
                    db,
                    "sections",
                    targetSection,
                    "students"
                ),

                {
                    qid,
                    name
                }
            );


            newStudentQid.value =
                "";


            newStudentName.value =
                "";


            if(adminPanelOpen){
                managementSection = targetSection;
                await refreshManageData();
            }else{
                await loadStudents();
            }
            await writeActivityLog("ADD_STUDENT",`${qid} · ${name}`,{section:targetSection});

        } catch (error) {

            console.error(error);


            alert(
                "Error adding student."
            );
        }
    }
);


// EDIT STUDENT

async function editStudent(
    studentId
) {

    if (
        !profile ||
        profile.role !==
            "admin"
    ) {

        return;
    }


    const student =
        students.find(
            item =>
                item.id ===
                studentId
        );


    if (!student) {

        alert(
            "Student not found."
        );

        return;
    }


    const newQid =
        prompt(
            "Enter Q.ID:",
            student.qid
        );


    if (
        newQid === null
    ) {

        return;
    }


    const newName =
        prompt(
            "Enter student name:",
            student.name
        );


    if (
        newName === null
    ) {

        return;
    }


    const qid =
        newQid.trim();


    const name =
        newName
            .trim()
            .toUpperCase();


    if (
        !qid ||
        !name
    ) {

        alert(
            "Q.ID and student name cannot be empty."
        );

        return;
    }


    const duplicate =
        students.some(
            function (item) {

                return (
                    item.id !==
                        studentId &&
                    item.qid.toLowerCase() ===
                        qid.toLowerCase()
                );
            }
        );


    if (duplicate) {

        alert(
            "A student with this Q.ID already exists in this section."
        );

        return;
    }


    try {

        await updateDoc(

            doc(
                db,
                "sections",
                activeSection,
                "students",
                studentId
            ),

            {
                qid:
                    qid,

                name:
                    name
            }
        );


        void writeActivityLog("EDIT_STUDENT",`${qid} · ${name}`,{section:activeSection});
        await loadStudents();


    } catch (error) {

        console.error(error);


        alert(
            "Error updating student."
        );
    }
}


// DELETE STUDENT

async function deleteStudent(
    studentId
) {

    if (
        !profile ||
        profile.role !==
            "admin"
    ) {

        return;
    }


    const student =
        students.find(
            item =>
                item.id ===
                studentId
        );


    const confirmDelete =
        confirm(
            `Delete ${
                student
                    ? student.name
                    : "this student"
            }?`
        );


    if (!confirmDelete) {
        return;
    }


    try {

        await deleteDoc(

            doc(
                db,
                "sections",
                activeSection,
                "students",
                studentId
            )
        );


        void writeActivityLog("DELETE_STUDENT",`${student.qid} · ${student.name}`,{section:activeSection});
        await loadStudents();


    } catch (error) {

        console.error(error);


        alert(
            "Error deleting student."
        );
    }
}


// UPDATE STATS

function updateStats() {

    const total =
        students.length;


    const present =
        students.filter(
            student =>
                student.status ===
                "Present"
        ).length;


    const absent =
        students.filter(
            student =>
                student.status ===
                "Absent"
        ).length;


    totalStudents.textContent =
        total;


    headerTotalStudents.textContent =
        total;


    presentStudents.textContent =
        present;


    absentStudents.textContent =
        absent;
}


// SEARCH

searchStudent?.addEventListener(
    "input",
    displayStudents
);


// MARK ALL PRESENT

markAllPresentBtn?.addEventListener(
    "click",
    function () {

        if (
            !subjectSelect.value
        ) {

            alert(
                "Please select a subject first."
            );

            return;
        }


        students.forEach(
            student =>
                student.status =
                    "Present"
        );


        displayStudents();
        saveAttendanceDraft();
    }
);


// MARK ALL ABSENT

markAllAbsentBtn?.addEventListener(
    "click",
    function () {

        if (
            !subjectSelect.value
        ) {

            alert(
                "Please select a subject first."
            );

            return;
        }


        students.forEach(
            student =>
                student.status =
                    "Absent"
        );


        displayStudents();
        saveAttendanceDraft();
    }
);


// LOAD SUBJECTS

async function loadSubjects() {

    if (!activeSection) {
        return;
    }


    try {

        const q =
            query(

                collection(
                    db,
                    "sections",
                    activeSection,
                    "subjects"
                ),

                orderBy(
                    "name"
                )
            );


        const snapshot =
            await getDocs(q);


        subjects =
            snapshot.docs.map(
                docSnap => ({

                    id:
                        docSnap.id,

                    name:
                        docSnap.data()
                            .name
                })
            );


        subjectSelect.innerHTML =
            `<option value="">
                -- Select Subject --
            </option>`;


        subjects.forEach(
            function (subject) {

                const option =
                    document.createElement(
                        "option"
                    );


                option.value =
                    subject.name;


                option.textContent =
                    subject.name;


                subjectSelect.appendChild(
                    option
                );
            }
        );


        if (
            profile.role ===
            "admin"
        ) {

            subjectManageList.innerHTML =
                "";


            subjects.forEach(
                function (subject) {

                    const row =
                        document.createElement(
                            "div"
                        );


                    row.className =
                        "chip-row";


                    row.innerHTML = `

                        <span>
                            ${subject.name}
                        </span>


                        <button
                            type="button"
                            data-id="${subject.id}">

                            Delete

                        </button>

                    `;


                    subjectManageList.appendChild(
                        row
                    );
                }
            );


            document
                .querySelectorAll(
                    "#subjectManageList button"
                )
                .forEach(
                    function (button) {

                        button.addEventListener(
                            "click",
                            function () {

                                deleteSubject(
                                    button.dataset.id
                                );
                            }
                        );
                    }
                );
        }


    } catch (error) {

        console.error(error);
    }
}


// ADD SUBJECT

addSubjectForm?.addEventListener(
    "submit",
    async function (event) {

        event.preventDefault();


        const adminPanelOpen = !document.getElementById("adminPanel")?.classList.contains("hidden");
        const targetSection = adminPanelOpen
            ? document.getElementById("manageSectionSelect")?.value || null
            : activeSection;

        if (!targetSection) {
            alert("Please select a section first.");
            return;
        }


        const name =
            newSubjectName.value.trim();


        if (!name) {
            return;
        }


        try {

            await addDoc(

                collection(
                    db,
                    "sections",
                    targetSection,
                    "subjects"
                ),

                {
                    name
                }
            );


            newSubjectName.value =
                "";


            if(adminPanelOpen){
                managementSection = targetSection;
                await refreshManageData();
            }else{
                await loadSubjects();
            }
            await writeActivityLog("ADD_SUBJECT",name,{section:targetSection});

        } catch (error) {

            console.error(error);


            alert(
                "Error adding subject."
            );
        }
    }
);


// DELETE SUBJECT

async function deleteSubject(
    subjectId
) {

    const confirmDelete =
        confirm(
            "Delete this subject from this section? Existing saved records will keep the old name."
        );


    if (
        !confirmDelete
    ) {
        return;
    }


    try {

        const targetSection = managementSection || activeSection;
        const deletedSubject = subjects.find(item => item.id === subjectId);

        if(!targetSection){
            alert("Please select a section first.");
            return;
        }

        await deleteDoc(
            doc(
                db,
                "sections",
                targetSection,
                "subjects",
                subjectId
            )
        );

        void writeActivityLog("DELETE_SUBJECT",deletedSubject?.name || "",{section:targetSection});
        await loadSubjects();


    } catch (error) {

        console.error(error);


        alert(
            "Error deleting subject."
        );
    }
}


// LOAD ALL CR PROFILES (used as a fallback / global list)
async function loadCrList() {
    if (!profile || profile.role !== "admin") return;

    const container = document.getElementById("crManageList");
    if (!container) return;

    try {
        const snapshot = await getDocs(collection(db, "users"));
        const crs = snapshot.docs
            .filter(d => d.data()?.role === "cr")
            .map(d => ({
                email: String(d.id || "").toLowerCase(),
                section: String(d.data()?.section || "")
            }))
            .sort((a,b) => a.email.localeCompare(b.email));

        container.innerHTML = "";
        if (!crs.length) {
            container.innerHTML = `<p class="empty-record">No CRs assigned yet.</p>`;
            return;
        }

        crs.forEach(cr => {
            const row = document.createElement("div");
            row.className = "chip-row";
            row.innerHTML = `
                <span>${escapeHtml(cr.email)} — ${escapeHtml(sectionLabelOf(cr.section))}</span>
                <button type="button" data-email="${escapeHtml(cr.email)}">Remove</button>
            `;
            container.appendChild(row);
        });
        bindCrRemoveButtons(container);
    } catch (error) {
        console.error("Could not load CR list:", error);
        container.innerHTML = `<p class="empty-record">Could not load CRs.</p>`;
    }
}

function bindCrRemoveButtons(container) {
    container.querySelectorAll("button[data-email]").forEach(button => {
        button.addEventListener("click", () => deleteCr(button.dataset.email));
    });
}

async function loadCrListForManagementSection() {
    const container = document.getElementById("crManageList");
    if (!container || profile?.role !== "admin") return;

    const section = managementSection || document.getElementById("manageSectionSelect")?.value || "";
    if (!section) {
        container.innerHTML = "";
        return;
    }

    container.innerHTML = `<p class="empty-record">Loading CRs...</p>`;

    try {
        const snapshot = await getDocs(collection(db, "users"));
        const crs = snapshot.docs
            .filter(d => {
                const data = d.data() || {};
                return data.role === "cr" && String(data.section || "") === String(section);
            })
            .sort((a,b) => String(a.id || "").localeCompare(String(b.id || "")));

        container.innerHTML = "";
        if (!crs.length) {
            container.innerHTML = `<p class="empty-record">No CR is assigned to ${escapeHtml(sectionLabelOf(section))}.</p>`;
            return;
        }

        crs.forEach(d => {
            const data = d.data() || {};
            const row = document.createElement("div");
            row.className = "chip-row";
            row.innerHTML = `
                <span>${escapeHtml(d.id)} — ${escapeHtml(sectionLabelOf(data.section))}</span>
                <button type="button" data-email="${escapeHtml(d.id)}">Remove</button>
            `;
            container.appendChild(row);
        });
        bindCrRemoveButtons(container);
    } catch (error) {
        console.error("Could not load CRs for management section:", error);
        container.innerHTML = `<p class="empty-record">Could not load CRs for this section.</p>`;
    }
}

async function loadManagementSubjects(section) {
    const container = document.getElementById("subjectManageList");
    if (!container || !section) return;

    try {
        const snapshot = await getDocs(
            query(collection(db, "sections", section, "subjects"), orderBy("name"))
        );

        container.innerHTML = "";
        if (!snapshot.docs.length) {
            container.innerHTML = `<p class="empty-record">No subjects in this section.</p>`;
            return;
        }

        snapshot.docs.forEach(d => {
            const row = document.createElement("div");
            row.className = "chip-row";
            row.innerHTML = `
                <span>${escapeHtml(d.data()?.name || "")}</span>
                <button type="button" data-subject-id="${escapeHtml(d.id)}">Delete</button>
            `;
            container.appendChild(row);
        });

        container.querySelectorAll("button[data-subject-id]").forEach(button => {
            button.addEventListener("click", () => deleteSubject(button.dataset.subjectId));
        });
    } catch (error) {
        console.error("Could not load management subjects:", error);
        container.innerHTML = `<p class="empty-record">Could not load subjects.</p>`;
    }
}

async function refreshManageData() {
    const section = managementSection || document.getElementById("manageSectionSelect")?.value || "";
    const message = document.getElementById("manageSectionMessage");
    const subjectContainer = document.getElementById("subjectManageList");
    const crContainer = document.getElementById("crManageList");

    if (!section) {
        if (message) message.textContent = "Please select a section to manage its data.";
        if (subjectContainer) subjectContainer.innerHTML = "";
        if (crContainer) crContainer.innerHTML = "";
        return;
    }

    managementSection = section;
    if (message) message.textContent = `Managing ${sectionLabelOf(section)}.`;
    if (subjectContainer) subjectContainer.innerHTML = `<p class="empty-record">Loading subjects...</p>`;

    await Promise.all([
        loadManagementSubjects(section),
        loadCrListForManagementSection()
    ]);
}


// ASSIGN CR

addCrForm?.addEventListener(
    "submit",
    async function (event) {

        event.preventDefault();


        if (
            !profile ||
            profile.role !==
                "admin"
        ) {
            return;
        }


        const email =
            newCrEmail.value
                .trim()
                .toLowerCase();

        const section = managementSection ||
            document.getElementById("manageSectionSelect")?.value || null;

        if (!section) {
            alert("Please select a section first.");
            return;
        }

        if (
            !email
        ) {
            return;
        }


        try {
            const sectionSnap = await getDoc(doc(db, "sections", section));
            const sectionData = sectionSnap.exists() ? sectionSnap.data() : {};
            const tempPassword = `QA@${Math.random().toString(36).slice(2, 8)}${Math.floor(10 + Math.random() * 90)}`;
            try {
                await createUserWithEmailAndPassword(getProvisioningAuth(), email, tempPassword);
            } catch (authError) {
                if (authError.code !== "auth/email-already-in-use") throw authError;
            } finally {
                try { await signOut(getProvisioningAuth()); } catch (_) {}
            }

            await setDoc(doc(db, "users", email), {
                role: "cr",
                email,
                qid: sectionData.crQid || "",
                name: sectionData.crName || "",
                mobile: sectionData.crMobile || "",
                section,
                sectionLabel: sectionData.label || section,
                course: sectionData.course || "",
                semester: sectionData.semester || "",
                year: sectionData.year || "",
                mentorName: sectionData.mentorName || "",
                mentorMobile: sectionData.mentorMobile || "",
                createdAt: new Date().toISOString()
            }, { merge: true });

            try {
                await sendDirectEmail({
                    toEmail: email,
                    subject: `QAttend - CR Account Assigned - ${sectionData.label || section}`,
                    message: `You have been assigned as a Class Representative.\n\nLogin Email: ${email}\nTemporary Password: ${tempPassword}\nCourse: ${sectionData.course || ""}\nSection: ${sectionData.label || section}\nSemester: ${sectionData.semester || ""}\nYear: ${sectionData.year || ""}`,
                    replyTo: EMAILJS_CONFIG.adminEmail
                });
            } catch (emailError) {
                console.warn("CR invitation email failed:", emailError);
            }


            newCrEmail.value =
                "";


            managementSection = section;
            await loadCrListForManagementSection();
            await writeActivityLog(
                "ASSIGN_CR",
                `${email} · ${sectionLabelOf(section)}`,
                {section}
            );


            alert(`${email} is assigned as CR for ${sectionLabelOf(section)}.\n\nTemporary password: ${tempPassword}`);


        } catch (error) {

            console.error(error);


            alert(
                "Error assigning CR. Make sure the login was created with the exact same lowercase email."
            );
        }
    }
);


// REMOVE CR

async function deleteCr(
    email
) {

    if (
        !profile ||
        profile.role !==
            "admin"
    ) {
        return;
    }


    const confirmDelete =
        confirm(
            `Remove CR access for ${email}? (Their login will still exist in Firebase Console.)`
        );


    if (
        !confirmDelete
    ) {
        return;
    }


    try {

        await deleteDoc(

            doc(
                db,
                "users",
                email
            )
        );


        await writeActivityLog(
            "REMOVE_CR",
            email,
            {section: managementSection || ""}
        );
        if (managementSection) {
            await loadCrListForManagementSection();
        } else {
            await loadCrList();
        }


    } catch (error) {

        console.error(error);


        alert(
            "Error removing CR."
        );
    }
}


// GET ATTENDANCE RECORD ID

function getAttendanceId() {

    const date =
        attendanceDate.value;


    const subject =
        subjectSelect.value;


    const cleanSubject =
        subject.replace(
            /[^a-zA-Z0-9]/g,
            "_"
        );


    return `${date}_${cleanSubject}`;
}

function getAttendanceDraftKey() {
    const email = String(auth.currentUser?.email || "guest").toLowerCase();
    return `qattend-draft:${email}:${activeSection || ""}:${getAttendanceId()}`;
}

function saveAttendanceDraft() {
    if (!activeSection || !attendanceDate?.value || !subjectSelect?.value || !students.length) return;
    try {
        localStorage.setItem(getAttendanceDraftKey(), JSON.stringify({
            date: attendanceDate.value,
            subject: subjectSelect.value,
            students: students.map(student => ({ qid: student.qid, status: student.status }))
        }));
    } catch (error) {
        console.warn("Could not save attendance draft:", error);
    }
}

function restoreAttendanceDraft() {
    if (!activeSection || !attendanceDate?.value || !subjectSelect?.value) return false;
    try {
        const draft = JSON.parse(localStorage.getItem(getAttendanceDraftKey()) || "null");
        if (!draft?.students) return false;
        const statusByQid = new Map(draft.students.map(student => [student.qid, student.status]));
        students.forEach(student => {
            if (statusByQid.has(student.qid)) student.status = statusByQid.get(student.qid);
        });
        return true;
    } catch (error) {
        console.warn("Could not restore attendance draft:", error);
        return false;
    }
}

function clearAttendanceDraft() {
    try { localStorage.removeItem(getAttendanceDraftKey()); } catch (_) {}
}


// LOAD ATTENDANCE

async function loadAttendance() {

    const date =
        attendanceDate.value;


    const subject =
        subjectSelect.value;


    if (
        !date ||
        !subject
    ) {
        return;
    }


    try {

        const attendanceRef =
            doc(

                db,

                "sections",

                activeSection,

                "attendance",

                getAttendanceId()

            );


        const snapshot =
            await getDoc(
                attendanceRef
            );


        if (
            snapshot.exists()
        ) {

            const data =
                snapshot.data();


            students.forEach(
                function (student) {

                    const saved =
                        data.students.find(
                            item =>
                                item.qid ===
                                student.qid
                        );


                    student.status =
                        saved
                            ? saved.status
                            : null;
                }
            );


        } else {

            students.forEach(
                student =>
                    student.status =
                        null
            );
        }

        restoreAttendanceDraft();


        displayStudents();


    } catch (error) {

        console.error(error);


        alert(
            "Error loading attendance."
        );
    }
}


// SUBJECT CHANGE

subjectSelect?.addEventListener(
    "change",
    function () {

        if (
            !attendanceDate.value
        ) {

            alert(
                "Please select date first."
            );


            subjectSelect.value =
                "";


            return;
        }


        loadAttendance();
        syncAttendanceGate();
    }
);


// DATE CHANGE

attendanceDate?.addEventListener(
    "change",
    function () {

        if (
            subjectSelect.value
        ) {

            loadAttendance();
        }
        syncAttendanceGate();
    }
);


// SAVE ATTENDANCE

saveBtn?.addEventListener(
    "click",
    async function () {

        const user =
            auth.currentUser;


        if (!user) {
            return;
        }


        const date =
            attendanceDate.value;


        const subject =
            subjectSelect.value;


        if (!date) {

            alert(
                "Please select date first."
            );

            return;
        }


        if (!subject) {

            alert(
                "Please select subject."
            );

            return;
        }


        const unmarked =
            students.filter(
                student =>
                    student.status ===
                    null
            );


        if (
            unmarked.length > 0
        ) {

            alert(
                `${unmarked.length} students are not marked.`
            );

            return;
        }


        try {

            saveBtn.textContent =
                "Saving...";


            const attendanceRef =
                doc(

                    db,

                    "sections",

                    activeSection,

                    "attendance",

                    getAttendanceId()

                );


            await setDoc(

                attendanceRef,

                {
                    university:
                        "Quantum University",

                    course:
                        "B.Tech",

                    section:
                        sectionLabelOf(
                            activeSection
                        ),

                    subject:
                        subject,

                    date:
                        date,

                    students:
                        students,

                    markedBy:
                        user.email,

                    updatedAt:
                        new Date()
                            .toISOString()
                }
            );

            clearAttendanceDraft();


            void writeActivityLog("SAVE_ATTENDANCE",`${sectionLabelOf(activeSection)} · ${subject} · ${date}`);

            alert(
                "✅ Attendance saved successfully!"
            );


            await loadRecords();


        } catch (error) {

            console.error(error);


            alert(
                "❌ Error saving attendance."
            );


        } finally {

            saveBtn.textContent =
                "💾 Save Attendance";
        }
    }
);


// LOAD RECORDS

async function loadRecords() {

    if (!activeSection) {
        return;
    }


    recordsList.innerHTML =
        `<p class="empty-record">
            Loading records...
        </p>`;


    try {

        const attendanceCollection =
            collection(

                db,

                "sections",

                activeSection,

                "attendance"

            );


        const q =
            query(

                attendanceCollection,

                orderBy(
                    "updatedAt",
                    "desc"
                )

            );


        const snapshot =
            await getDocs(q);


        recordsList.innerHTML =
            "";


        if (
            snapshot.empty
        ) {

            recordsList.innerHTML =
                `<p class="empty-record">
                    No previous attendance records found.
                </p>`;


            return;
        }


        snapshot.forEach(
            function (
                documentSnapshot
            ) {

                const data =
                    documentSnapshot.data();


                recordsCache[
                    documentSnapshot.id
                ] = data;


                const present =
                    data.students.filter(
                        student =>
                            student.status ===
                            "Present"
                    ).length;


                const absent =
                    data.students.filter(
                        student =>
                            student.status ===
                            "Absent"
                    ).length;


                const card =
                    document.createElement(
                        "div"
                    );


                card.className =
                    "record-card";


                card.innerHTML = `

                    <div>

                        <h3>
                            ${data.subject}
                        </h3>

                        <p>
                            📅 ${data.date}
                        </p>

                        <p>
                            🟢 Present:
                            ${present}
                            |
                            🔴 Absent:
                            ${absent}
                        </p>

                    </div>


                    <div class="record-actions">

                        <button
                            class="view-record-btn"
                            data-id="${documentSnapshot.id}">

                            View

                        </button>


                        <button
                            class="share-record-btn"
                            data-id="${documentSnapshot.id}">

                            Share

                        </button>


                        <button
                            class="delete-record-btn"
                            data-id="${documentSnapshot.id}">

                            Delete

                        </button>

                    </div>

                `;


                recordsList.appendChild(
                    card
                );
            }
        );


        addRecordButtonEvents();


    } catch (error) {

        console.error(error);


        recordsList.innerHTML =
            `<p class="empty-record">
                Error loading records.
            </p>`;
    }
}


// RECORD EVENTS

function addRecordButtonEvents() {

    document
        .querySelectorAll(
            ".view-record-btn"
        )
        .forEach(
            function (button) {

                button.addEventListener(
                    "click",
                    async function () {

                        await viewRecord(
                            button.dataset.id
                        );
                    }
                );
            }
        );


    document
        .querySelectorAll(
            ".share-record-btn"
        )
        .forEach(
            function (button) {

                button.addEventListener(
                    "click",
                    async function () {

                        await shareRecord(
                            button.dataset.id
                        );
                    }
                );
            }
        );


    document
        .querySelectorAll(
            ".delete-record-btn"
        )
        .forEach(
            function (button) {

                button.addEventListener(
                    "click",
                    async function () {

                        await deleteRecord(
                            button.dataset.id
                        );
                    }
                );
            }
        );
}


// SHARE RECORD

async function shareRecord(
    recordId
) {

    try {

        const data =
            recordsCache[
                recordId
            ];


        if (!data) {

            alert(
                "Record data not loaded — try refreshing the records list first."
            );


            return;
        }


        await ensureXLSXLoaded();


        const serialByStudent = new Map();


        data.students.forEach(
            function (
                student,
                index
            ) {

                serialByStudent.set(student, index + 1);
            }
        );


        const sortedStudents =
            [...data.students].sort(
                function (a, b) {

                    if (
                        a.status === "Present" &&
                        b.status === "Absent"
                    ) {

                        return -1;
                    }


                    if (
                        a.status === "Absent" &&
                        b.status === "Present"
                    ) {

                        return 1;
                    }


                    return 0;
                }
            );


        const excelData = [

            [
                "QATTEND ATTENDANCE REPORT"
            ],

            [],

            [
                "University",
                data.university ||
                    "Quantum University"
            ],

            [
                "Course",
                data.course ||
                    "B.Tech"
            ],

            [
                "Section",
                data.section
            ],

            [
                "Semester",
                "5"
            ],

            [
                "Subject",
                data.subject
            ],

            [
                "Date",
                data.date
            ],

            [],

            [
                "S.No",
                "Q.ID",
                "Name of Student",
                "Attendance"
            ]
        ];


        const presentSerials = [];


        sortedStudents.forEach(
            function (student) {

                const serial =
                    serialByStudent.get(student);


                excelData.push([

                    serial,

                    student.qid,

                    student.name,

                    student.status ===
                        "Present"
                        ? 1
                        : 0

                ]);


                if (
                    student.status ===
                    "Present"
                ) {

                    presentSerials.push(
                        serial
                    );
                }
            }
        );


        presentSerials.sort(
            (a, b) =>
                a - b
        );


        excelData.push([]);


        excelData.push([

            "Present (S.No, for shortcut)",

            presentSerials.join(" ")

        ]);


        const worksheet =
            XLSX.utils.aoa_to_sheet(
                excelData
            );


        worksheet["!cols"] = [

            { wch: 10 },

            { wch: 18 },

            { wch: 50 },

            { wch: 15 }

        ];

        worksheet["!autofilter"] = { ref: `A10:D${10 + data.students.length}` };


        const workbook =
            XLSX.utils.book_new();


        XLSX.utils.book_append_sheet(

            workbook,

            worksheet,

            "Attendance"

        );


        const cleanSubject =
            data.subject.replace(
                /[^a-zA-Z0-9]/g,
                "_"
            );


        const fileName =
            `QAttend_${data.section}_${cleanSubject}_${data.date}.xlsx`;


        const wbArray =
            XLSX.write(

                workbook,

                {
                    bookType:
                        "xlsx",

                    type:
                        "array"
                }

            );


        const blob =
            new Blob(

                [wbArray],

                {
                    type:
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }

            );


        const file =
            new File(

                [blob],

                fileName,

                {
                    type:
                        blob.type
                }

            );


        if (
            navigator.canShare &&
            navigator.canShare({
                files: [file]
            })
        ) {

            try {

                await navigator.share({

                    files: [file],

                    title:
                        `${data.subject} Attendance`,

                    text:
                        `Attendance for ${data.subject} on ${data.date} — ${data.section}`

                });


                return;


            } catch (
                shareError
            ) {

                if (
                    shareError.name ===
                    "AbortError"
                ) {

                    return;
                }


                console.error(
                    shareError
                );
            }
        }


        XLSX.writeFile(
            workbook,
            fileName
        );


        alert(
            "Sharing isn't available in this browser — the file was downloaded instead. You can share it manually."
        );


    } catch (error) {

        console.error(error);


        alert(
            `Error sharing record: ${
                error.name || ""
            } — ${
                error.message || error
            }`
        );
    }
}


// VIEW RECORD

async function viewRecord(
    recordId
) {

    try {

        const recordRef =
            doc(

                db,

                "sections",

                activeSection,

                "attendance",

                recordId

            );


        const snapshot =
            await getDoc(
                recordRef
            );


        if (
            !snapshot.exists()
        ) {

            alert(
                "Record not found."
            );


            return;
        }


        const data =
            snapshot.data();


        attendanceDate.value =
            data.date;


        subjectSelect.value =
            data.subject;


        students.forEach(
            function (student) {

                const saved =
                    data.students.find(
                        item =>
                            item.qid ===
                            student.qid
                    );


                student.status =
                    saved
                        ? saved.status
                        : null;
            }
        );


        // Make sure the attendance UI is explicitly opened when an admin
        // clicks View from Previous Attendance. The normal attendance gate
        // can otherwise remain hidden from the previous screen state.
        workArea?.classList.remove("hidden");
        noSectionMessage?.classList.add("hidden");
        document.getElementById("attendanceGateMessage")?.classList.add("hidden");
        document.querySelector(".stats")?.classList.remove("hidden");
        document.querySelector(".top-actions")?.classList.remove("hidden");
        document.querySelector(".attendance-section")?.classList.remove("hidden");
        document.querySelector(".action-buttons")?.classList.remove("hidden");

        displayStudents();

        // Re-sync the gate after restoring the record so the normal view
        // state is consistent with the selected section/date/subject.
        syncAttendanceGate();

        window.scrollTo({

            top: 0,

            behavior:
                "smooth"

        });


    } catch (error) {

        console.error(error);


        alert(
            "Error loading record."
        );
    }
}


// DELETE RECORD

// NOTE:
// This is intentionally NOT admin-only.
// CRs can delete attendance records from their assigned section,
// provided Firestore rules allow that write.

async function deleteRecord(
    recordId
) {

    const confirmDelete =
        confirm(
            "Are you sure you want to delete this attendance record?"
        );


    if (
        !confirmDelete
    ) {
        return;
    }


    try {

        await deleteDoc(

            doc(

                db,

                "sections",

                activeSection,

                "attendance",

                recordId

            )
        );


        alert(
            "Attendance record deleted."
        );


        void writeActivityLog("DELETE_ATTENDANCE",recordId,{section:activeSection});
        await loadRecords();


    } catch (error) {

        console.error(error);


        alert(
            "Error deleting record."
        );
    }
}


// REFRESH

refreshRecordsBtn?.addEventListener(
    "click",
    loadRecords
);


// DOWNLOAD XLSX

downloadBtn?.addEventListener(
    "click",
    async function () {

        const date =
            attendanceDate.value;


        const subject =
            subjectSelect.value;


        if (
            !date ||
            !subject
        ) {

            alert(
                "Select date and subject first."
            );


            return;
        }


        const unmarked =
            students.filter(
                student =>
                    student.status ===
                    null
            );


        if (
            unmarked.length > 0
        ) {

            alert(
                "Please mark all students before downloading."
            );


            return;
        }


        try {

            await ensureXLSXLoaded();


        } catch (error) {

            alert(
                error.message
            );


            return;
        }


        const serialByStudent = new Map();


        students.forEach(
            function (
                student,
                index
            ) {

                serialByStudent.set(student, index + 1);
            }
        );


        const sortedStudents =
            [...students].sort(
                function (a, b) {

                    if (
                        a.status === "Present" &&
                        b.status === "Absent"
                    ) {

                        return -1;
                    }


                    if (
                        a.status === "Absent" &&
                        b.status === "Present"
                    ) {

                        return 1;
                    }


                    return 0;
                }
            );


        const excelData = [

            [
                "QATTEND ATTENDANCE REPORT"
            ],

            [],

            [
                "University",
                "Quantum University"
            ],

            [
                "Course",
                "B.Tech"
            ],

            [
                "Section",
                sectionLabelOf(
                    activeSection
                )
            ],

            [
                "Semester",
                "5"
            ],

            [
                "Subject",
                subject
            ],

            [
                "Date",
                date
            ],

            [],

            [
                "S.No",
                "Q.ID",
                "Name of Student",
                "Attendance"
            ]
        ];


        const presentSerials = [];


        sortedStudents.forEach(
            function (student) {

                const serial =
                    serialByStudent.get(student);


                excelData.push([

                    serial,

                    student.qid,

                    student.name,

                    student.status ===
                        "Present"
                        ? 1
                        : 0

                ]);


                if (
                    student.status ===
                    "Present"
                ) {

                    presentSerials.push(
                        serial
                    );
                }
            }
        );


        presentSerials.sort(
            (a, b) =>
                a - b
        );


        excelData.push([]);


        excelData.push([

            "Present (S.No, for shortcut)",

            presentSerials.join(" ")

        ]);


        const worksheet =
            XLSX.utils.aoa_to_sheet(
                excelData
            );


        worksheet["!cols"] = [

            { wch: 10 },

            { wch: 18 },

            { wch: 50 },

            { wch: 15 }

        ];

        worksheet["!autofilter"] = { ref: `A10:D${10 + students.length}` };


        const workbook =
            XLSX.utils.book_new();


        XLSX.utils.book_append_sheet(

            workbook,

            worksheet,

            "Attendance"

        );


        const cleanSubject =
            subject.replace(
                /[^a-zA-Z0-9]/g,
                "_"
            );


        XLSX.writeFile(

            workbook,

            `QAttend_${sectionLabelOf(
                activeSection
            )}_${cleanSubject}_${date}.xlsx`

        );
    }
);


/* =========================================================
   MORE DRAWER / REQUEST CENTER / ADMIN ACTIVITY LOGS
   ========================================================= */

const moreMenuBtn = document.getElementById("moreMenuBtn");
const featureDrawer = document.getElementById("featureDrawer");
const featureBackdrop = document.getElementById("featureBackdrop");
const closeFeatureDrawer = document.getElementById("closeFeatureDrawer");
const requestsPanel = document.getElementById("requestsPanel");
const adminDashboard = document.getElementById("adminDashboard");
const requestList = document.getElementById("requestList");
const bulkRequestToolbar = document.getElementById("bulkRequestToolbar");
const selectAllRequests = document.getElementById("selectAllRequests");
const selectedRequestsCount = document.getElementById("selectedRequestsCount");
const bulkApproveRequestsBtn = document.getElementById("bulkApproveRequestsBtn");
const bulkRejectRequestsBtn = document.getElementById("bulkRejectRequestsBtn");
const activityLogList = document.getElementById("activityLogList");
const dashboardRequestList = document.getElementById("dashboardRequestList");
const requestModal = document.getElementById("requestModal");
const supportModal = document.getElementById("supportModal");
const manageSectionSelect = document.getElementById("manageSectionSelect");
const manageSectionMessage = document.getElementById("manageSectionMessage");
const manageCourseSelect = document.getElementById("manageCourseSelect");

let managementSection = null;

function openFeatureDrawer(){
    featureDrawer?.classList.add("open");
    featureBackdrop?.classList.remove("hidden");
    featureDrawer?.setAttribute("aria-hidden","false");
    moreMenuBtn?.setAttribute("aria-expanded","true");
}

function closeFeatureDrawerFn(){
    featureDrawer?.classList.remove("open");
    featureBackdrop?.classList.add("hidden");
    featureDrawer?.setAttribute("aria-hidden","true");
    moreMenuBtn?.setAttribute("aria-expanded","false");
}

moreMenuBtn?.addEventListener("click", e => {
    e.stopPropagation();
    openFeatureDrawer();
});

closeFeatureDrawer?.addEventListener("click", closeFeatureDrawerFn);
featureBackdrop?.addEventListener("click", closeFeatureDrawerFn);

document.addEventListener("keydown", e => {
    if(e.key === "Escape"){
        closeFeatureDrawerFn();
        requestModal?.classList.add("hidden");
        supportModal?.classList.add("hidden");
    }
});

function hideMainViews(){
    document.querySelectorAll(".welcome-section,.control-section,#adminPanel,#adminCoursesView,#adminSectionsView,#noSectionMessage,#workArea").forEach(el => {
        el?.classList.add("hidden");
    });
    requestsPanel?.classList.add("hidden");
    adminDashboard?.classList.add("hidden");
}

function syncAttendanceGate(){
    const dateReady = Boolean(attendanceDate.value);
    const subjectReady = Boolean(subjectSelect.value);
    const sectionReady = Boolean(activeSection);
    const ready = sectionReady && dateReady && subjectReady;

    const gate = document.getElementById("attendanceGateMessage");
    const recordsSection = document.querySelector(".records-section");
    const attendanceSection = document.querySelector(".attendance-section");
    const statsSection = document.querySelector(".stats");
    const topActions = document.querySelector(".top-actions");
    const saveActions = document.querySelector(".action-buttons");

    if(gate){
        gate.classList.toggle("hidden", ready);
        if(!sectionReady){
            gate.textContent = "Select a section to begin attendance.";
        }else if(!subjectReady){
            gate.textContent = "Select a subject to open the student attendance list.";
        }else if(!dateReady){
            gate.textContent = "Select an attendance date to open the student list.";
        }
    }

    if(profile?.role === "admin"){
        if(workArea){
            workArea.classList.toggle("hidden", !sectionReady);
        }
        noSectionMessage?.classList.toggle("hidden", sectionReady);

        [statsSection, topActions, attendanceSection, saveActions].forEach(el=>{
            el?.classList.toggle("hidden", !ready);
        });

        recordsSection?.classList.toggle("hidden", !sectionReady);
        return;
    }

    noSectionMessage?.classList.add("hidden");
    workArea?.classList.toggle("hidden", !ready);
}


function showAttendanceView(){
    document.body.classList.remove("records-only-view");
    hideMainViews();

    const welcome = document.querySelector(".welcome-section");
    const control = document.querySelector(".control-section");

    adminDashboard?.classList.add("hidden");
    requestsPanel?.classList.add("hidden");
    adminPanel?.classList.add("hidden");

    if(profile?.role === "admin"){
        welcome?.classList.add("hidden");
        activeSection = sectionSelect?.value || null;
    }else{
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
        option.hidden = !courseId || String(sectionCourseMap[option.value] || "").toUpperCase() !== String(courseId).toUpperCase();
    });
    if (selectedSection && !manageSectionSelect.selectedOptions[0]?.hidden) return;
    managementSection = null;
    manageSectionSelect.value = "";
    subjectManageList.innerHTML = "";
    crManageList.innerHTML = "";
    if (manageSectionMessage) manageSectionMessage.textContent = courseId ? "Please select a section to manage its data." : "Please select a course first.";
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
                const sectionId=request.section;
                if(!sectionId) throw new Error("Section ID is missing.");
                const requestedCourse = String(request.course || "").trim();
                const courseSnap = await getDocs(collection(db, "courses"));
                const existingCourse = courseSnap.docs.find(courseDoc => {
                    const course = courseDoc.data() || {};
                    return String(course.name || "").trim().toLowerCase() === requestedCourse.toLowerCase()
                        || String(course.code || "").trim().toLowerCase() === requestedCourse.toLowerCase();
                });
                const courseId = existingCourse?.id || requestedCourse.toUpperCase().replace(/[^A-Z0-9]+/g, "-").replace(/^-|-$/g, "") || `COURSE-${Date.now()}`;
                if (!existingCourse) {
                    await setDoc(doc(db, "courses", courseId), {
                        name: requestedCourse || courseId,
                        code: courseId,
                        department: request.department || "",
                        duration: request.duration || "",
                        createdAt: new Date().toISOString()
                    }, { merge: true });
                }
                await setDoc(doc(db,"sections",sectionId),{
                    label:request.sectionLabel||sectionId, course:courseId, courseName:requestedCourse, semester:request.semester||"", year:request.year||"",
                    crQid:request.crQid||"", crName:request.crName||"", crMobile:request.crMobile||"", crEmail:request.crEmail||"",
                    mentorName:request.mentorName||"", mentorMobile:request.mentorMobile||"", createdAt:new Date().toISOString()
                },{merge:true});
                const batch=writeBatch(db);
                (request.students||[]).forEach(st=>{const ref=doc(collection(db,"sections",sectionId,"students"));batch.set(ref,{qid:String(st.qid||""),name:String(st.name||"").toUpperCase()});});
                (request.subjects||[]).forEach(name=>{const ref=doc(collection(db,"sections",sectionId,"subjects"));batch.set(ref,{name:String(name)});});
                await batch.commit();

                let tempPassword=request.generatedPassword||`QA@${Math.random().toString(36).slice(2,8)}${Math.floor(10+Math.random()*90)}`;
                try{await createUserWithEmailAndPassword(getProvisioningAuth(),request.crEmail,tempPassword);}
                catch(authError){if(authError.code!=="auth/email-already-in-use")throw authError;}
                finally{try{await signOut(getProvisioningAuth());}catch(_) {}}
                await setDoc(doc(db,"users",request.crEmail),{
                    role:"cr",email:request.crEmail,qid:request.crQid,name:request.crName,mobile:request.crMobile,section:sectionId,sectionLabel:request.sectionLabel||sectionId,
                    course:request.course||"",semester:request.semester||"",year:request.year||"",mentorName:request.mentorName||"",mentorMobile:request.mentorMobile||"",createdAt:new Date().toISOString()
                },{merge:true});
                // Send the CR invitation directly by email. Bulk/silent processing keeps the previous behavior and skips the invitation email.
                if(!silent) {
                    const mailSubject=`QAttend - CR Account Assigned - ${request.sectionLabel||sectionId}`;
                    const mailBody=`Hello ${request.crName},\n\nYou have been assigned as the Class Representative (CR) for ${request.sectionLabel||sectionId}.\n\nLogin Email: ${request.crEmail}\nTemporary Password: ${tempPassword}\nCourse: ${request.course||""}\nSemester: ${request.semester||""}\nYear: ${request.year||""}\n\nPlease change your password after your first login.`;
                    try {
                        await sendDirectEmail({
                            toEmail: request.crEmail,
                            subject: mailSubject,
                            message: mailBody,
                            replyTo: EMAILJS_CONFIG.adminEmail
                        });
                    } catch(emailError) {
                        console.error("CR invitation email error:", emailError);
                        alert(`CR account was created, but the invitation email could not be sent. ${emailError.message || "Please try again."}`);
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
        const btechOption = document.createElement("option");
        btechOption.value = "BTECH";
        btechOption.textContent = "B.Tech";
        select.appendChild(btechOption);
        courses.forEach(courseDoc => {
            const course = courseDoc.data();
            const courseName = String(course.name || courseDoc.id);
            if (courseName.trim().toLowerCase() === "b.tech" || courseName.trim().toLowerCase() === "b.tech ai & ml") return;
            const option = document.createElement("option");
            option.value = courseDoc.id;
            option.textContent = `${course.name || courseDoc.id} (${course.code || courseDoc.id})`;
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
    const snap = await getDocs(collection(db, "sections"));
    const rows = snap.docs.filter(sectionDoc => sectionDoc.data()?.course === courseId);
    if (!rows.length) sectionTableBody.innerHTML = `<tr><td colspan="7" class="empty-record">No sections found for this course.</td></tr>`;
    for (const [index, sectionDoc] of rows.entries()) {
        const section = sectionDoc.data();
        const [studentsSnap, subjectsSnap] = await Promise.all([
            getDocs(collection(db, "sections", sectionDoc.id, "students")),
            getDocs(collection(db, "sections", sectionDoc.id, "subjects"))
        ]);
        const row = document.createElement("tr");
        row.innerHTML = `<td>${index + 1}</td><td>${escapeHtml(section.label || section.section || sectionDoc.id)}</td><td>${escapeHtml(section.semester || "")}</td><td>${escapeHtml(section.year || "")}</td><td>${studentsSnap.size}</td><td>${subjectsSnap.size}</td><td><button type="button" class="text-btn" data-open-section="${escapeHtml(sectionDoc.id)}">Open</button> <button type="button" class="text-btn danger-text" data-delete-section="${escapeHtml(sectionDoc.id)}">Delete</button></td>`;
        row.querySelector("[data-open-section]")?.addEventListener("click", () => {
            managementSection = sectionDoc.id;
            showManageData();
            if (manageSectionSelect) manageSectionSelect.value = sectionDoc.id;
        });
        row.querySelector("[data-delete-section]")?.addEventListener("click", async () => {
            if (!confirm(`Delete ${section.label || sectionDoc.id}?`)) return;
            await deleteDoc(sectionDoc.ref);
            await loadSectionsForCourse(courseId);
        });
        sectionTableBody.appendChild(row);
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
    const sectionId = `${courseId}-${sectionLabelValue}`.toUpperCase().replace(/[^A-Z0-9]+/g, "-").replace(/^-|-$/g, "");
    await setDoc(doc(db, "sections", sectionId), {
        course: courseId,
        label: sectionLabelValue,
        semester: document.getElementById("adminSecSemester").value.trim(),
        year: document.getElementById("adminSecYear").value.trim(),
        mentorName: document.getElementById("adminSecMentorName").value.trim(),
        mentorMobile: document.getElementById("adminSecMentorMobile").value.trim(),
        createdAt: new Date().toISOString()
    }, { merge: true });
    const batch = writeBatch(db);
    adminSectionExcelPayload.students.forEach(student => batch.set(doc(collection(db, "sections", sectionId, "students")), { qid: student.qid, name: student.name }));
    adminSectionExcelPayload.subjects.forEach(subject => batch.set(doc(collection(db, "sections", sectionId, "subjects")), { name: subject }));
    await batch.commit();
    closeModalById("adminAddSectionModal");
    await populateSectionDropdowns();
    await loadSectionsForCourse(courseId);
});



/*
 * When connectivity returns, Firestore automatically synchronizes pending
 * writes. Refresh only the currently visible view so the UI reflects the
 * server state after synchronization. Failures are intentionally ignored.
 */
window.addEventListener("qattend:network-restored", () => {
    setTimeout(() => {
        try {
            if (!profile) return;

            const currentPage =
                (window.location.pathname.split("/").pop() || "login.html").toLowerCase();

            if (currentPage === "admin.html") {
                if (typeof loadAdminDashboardData === "function") void loadAdminDashboardData();
                if (!requestsPanel?.classList.contains("hidden") && typeof loadRequestCenter === "function") {
                    void loadRequestCenter();
                }
                if (!adminPanel?.classList.contains("hidden") && typeof refreshManageData === "function") {
                    void refreshManageData();
                }
            } else if (currentPage === "cr.html") {
                if (typeof loadStudents === "function") void loadStudents();
                if (typeof loadSubjects === "function") void loadSubjects();
                if (typeof loadRecords === "function") void loadRecords();
            }
        } catch (error) {
            console.warn("Could not refresh QAttend after reconnect:", error);
        }
    }, 1200);
});
