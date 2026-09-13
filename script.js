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
    updateDoc,
    query,
    orderBy,
    limit
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


// HTML ELEMENTS

const loginPage =
    document.getElementById(
        "loginPage"
    );

const appPage =
    document.getElementById(
        "appPage"
    );


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

const activityMenuBtn =
    document.getElementById(
        "activityMenuBtn"
    );

const activitySidebar =
    document.getElementById(
        "activitySidebar"
    );

const activitySidebarOverlay =
    document.getElementById(
        "activitySidebarOverlay"
    );

const closeActivitySidebarBtn =
    document.getElementById(
        "closeActivitySidebarBtn"
    );

const refreshActivityBtn =
    document.getElementById(
        "refreshActivityBtn"
    );

const activityLogList =
    document.getElementById(
        "activityLogList"
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


// ACTIVITY LOG
// Every meaningful action (login, add/delete student, add/delete subject,
// assign/remove CR, save/delete attendance) writes one entry here. Admin-only
// to read (enforced in firestore.rules too, not just hidden in the UI).

const ACTION_LABELS = {
    login: { label: "Logged in", badge: "login" },
    add_student: { label: "Added student", badge: "add" },
    delete_student: { label: "Deleted student", badge: "delete" },
    edit_student: { label: "Edited student", badge: "add" },
    add_subject: { label: "Added subject", badge: "add" },
    delete_subject: { label: "Deleted subject", badge: "delete" },
    assign_cr: { label: "Assigned CR", badge: "cr" },
    remove_cr: { label: "Removed CR", badge: "delete" },
    save_attendance: { label: "Saved attendance", badge: "save" },
    delete_attendance: { label: "Deleted attendance record", badge: "delete" }
};

async function logActivity(action, details, sectionOverride) {
    try {
        const user = auth.currentUser;
        if (!user || !profile) return;

        await addDoc(collection(db, "activityLogs"), {
            action: action,
            details: details || "",
            performedBy: user.email,
            role: profile.role,
            section: sectionOverride !== undefined ? sectionOverride : (activeSection || null),
            timestamp: new Date().toISOString()
        });
    } catch (error) {
        // Never let a logging failure break the actual feature the person
        // is using — just note it in the console.
        console.error("Activity log error:", error);
    }
}

function formatLogTime(isoString) {
    try {
        const date = new Date(isoString);
        return date.toLocaleString(undefined, {
            day: "2-digit",
            month: "short",
            hour: "2-digit",
            minute: "2-digit"
        });
    } catch (error) {
        return isoString;
    }
}

async function loadActivityLog() {
    activityLogList.innerHTML = `<p class="empty-record">Loading activity...</p>`;

    try {
        const q = query(
            collection(db, "activityLogs"),
            orderBy("timestamp", "desc"),
            limit(150)
        );
        const snapshot = await getDocs(q);

        if (snapshot.empty) {
            activityLogList.innerHTML = `<p class="empty-record">No activity yet.</p>`;
            return;
        }

        activityLogList.innerHTML = "";

        snapshot.forEach(function (docSnap) {
            const entry = docSnap.data();
            const meta = ACTION_LABELS[entry.action] || { label: entry.action, badge: "add" };

            const item = document.createElement("div");
            item.className = "activity-log-item";
            item.innerHTML = `
                <div class="activity-top-row">
                    <span class="activity-badge ${meta.badge}">${meta.label}</span>
                    <span class="activity-time">${formatLogTime(entry.timestamp)}</span>
                </div>
                ${entry.details ? `<div class="activity-detail">${entry.details}</div>` : ""}
                <div class="activity-by">
                    ${entry.performedBy} (${entry.role === "admin" ? "Admin" : "CR"})
                    ${entry.section ? " — " + sectionLabelOf(entry.section) : ""}
                </div>
            `;
            activityLogList.appendChild(item);
        });

    } catch (error) {
        console.error(error);
        activityLogList.innerHTML = `<p class="empty-record">Error loading activity log.</p>`;
    }
}

function openActivitySidebar() {
    activitySidebar.classList.add("open");
    activitySidebarOverlay.classList.remove("hidden");
    requestAnimationFrame(function () {
        activitySidebarOverlay.classList.add("show");
    });
    activityMenuBtn.classList.add("open");
    loadActivityLog();
}

function closeActivitySidebar() {
    activitySidebar.classList.remove("open");
    activitySidebarOverlay.classList.remove("show");
    activityMenuBtn.classList.remove("open");
    setTimeout(function () {
        activitySidebarOverlay.classList.add("hidden");
    }, 300);
}

activityMenuBtn.addEventListener("click", function () {
    if (activitySidebar.classList.contains("open")) {
        closeActivitySidebar();
    } else {
        openActivitySidebar();
    }
});

closeActivitySidebarBtn.addEventListener("click", closeActivitySidebar);
activitySidebarOverlay.addEventListener("click", closeActivitySidebar);
refreshActivityBtn.addEventListener("click", loadActivityLog);


// TODAY DATE

attendanceDate.value =
    new Date()
        .toISOString()
        .split("T")[0];


// PASSWORD SHOW / HIDE

togglePassword.addEventListener(
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

loginForm.addEventListener(
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


// FORGOT PASSWORD

forgotPasswordLink.addEventListener(
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

backToLoginLink.addEventListener(
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

sendResetBtn.addEventListener(
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

            appPage.classList.add(
                "hidden"
            );

            loginPage.classList.remove(
                "hidden"
            );

            profile = null;

            document.body.classList.remove(
                "admin-view",
                "cr-view"
            );


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

            return;
        }


        try {

            const profileSnap =
                await getDoc(
                    doc(
                        db,
                        "users",
                        user.email
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


        loginPage.classList.add(
            "hidden"
        );

        appPage.classList.remove(
            "hidden"
        );


        loggedUser.textContent =
            user.email;


        roleBadge.textContent =
            profile.role === "admin"
                ? "ADMIN"
                : "CLASS REPRESENTATIVE";


        const isAdmin =
            profile.role === "admin";


        if (!sessionStorage.getItem("qattendLoggedThisSession")) {
            sessionStorage.setItem("qattendLoggedThisSession", "1");
            logActivity("login", `${user.email} logged in as ${profile.role}`);
        }


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

            welcomeTitle.textContent =
                "Admin Dashboard";


            welcomeSubtitle.textContent =
                "Manage every section, subject and CR from one place.";


            populateSectionDropdowns();


            activeSection =
                null;


            sectionSelect.value =
                "";


            workArea.classList.add(
                "hidden"
            );


            noSectionMessage.classList.remove(
                "hidden"
            );


            sectionLabel.textContent =
                "-- Not selected --";


            headerTotalStudents.textContent =
                "0";


            await loadCrList();

            await loadSubjects();


        } else {

            welcomeTitle.textContent =
                `${sectionLabelOf(
                    profile.section
                )} Attendance Portal`;


            welcomeSubtitle.textContent =
                "Mark, save and download attendance for your section.";


            activeSection =
                profile.section;


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


            await loadSubjects();

            await loadStudents();

            await loadRecords();
        }
    }
);


// LOGOUT

logoutBtn.addEventListener(
    "click",
    async function () {

        sessionStorage.removeItem("qattendLoggedThisSession");
        await signOut(auth);

    }
);


// SECTION DROPDOWNS

function populateSectionDropdowns() {

    sectionSelect.innerHTML =
        `<option value="">
            -- Select Section --
        </option>`;


    newCrSection.innerHTML =
        "";


    SECTIONS.forEach(
        function (section) {

            const option1 =
                document.createElement(
                    "option"
                );


            option1.value =
                section.id;


            option1.textContent =
                section.label;


            sectionSelect.appendChild(
                option1
            );


            const option2 =
                document.createElement(
                    "option"
                );


            option2.value =
                section.id;


            option2.textContent =
                section.label;


            newCrSection.appendChild(
                option2
            );
        }
    );
}


// SECTION CHANGE

sectionSelect.addEventListener(
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


        subjectSelect.value =
            "";


        attendanceDate.value =
            new Date()
                .toISOString()
                .split("T")[0];


        await loadStudents();

        await loadSubjects();

        await loadRecords();
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

studentList.addEventListener(
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

addStudentForm.addEventListener(
    "submit",
    async function (event) {

        event.preventDefault();


        if (
            !activeSection
        ) {

            alert(
                "Pick a section first (dropdown above)."
            );


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
                    activeSection,
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


            await loadStudents();

            logActivity("add_student", `${name} (${qid}) added to ${sectionLabelOf(activeSection)}`);


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


        await loadStudents();

        logActivity("edit_student", `${student.name} → ${name} (${qid}) in ${sectionLabelOf(activeSection)}`);


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


        await loadStudents();

        logActivity("delete_student", `${student ? student.name + " (" + student.qid + ")" : studentId} deleted from ${sectionLabelOf(activeSection)}`);


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

searchStudent.addEventListener(
    "input",
    displayStudents
);


// MARK ALL PRESENT

markAllPresentBtn.addEventListener(
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
    }
);


// MARK ALL ABSENT

markAllAbsentBtn.addEventListener(
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

addSubjectForm.addEventListener(
    "submit",
    async function (event) {

        event.preventDefault();


        if (
            !activeSection
        ) {

            alert(
                "Pick a section first (dropdown above)."
            );

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
                    activeSection,
                    "subjects"
                ),

                {
                    name
                }
            );


            newSubjectName.value =
                "";


            await loadSubjects();

            logActivity("add_subject", `"${name}" added to ${sectionLabelOf(activeSection)}`);


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


    const subjectToDelete =
        subjects.find(
            item => item.id === subjectId
        );


    try {

        await deleteDoc(

            doc(
                db,
                "sections",
                activeSection,
                "subjects",
                subjectId
            )
        );


        await loadSubjects();

        logActivity("delete_subject", `"${subjectToDelete ? subjectToDelete.name : subjectId}" deleted from ${sectionLabelOf(activeSection)}`);


    } catch (error) {

        console.error(error);


        alert(
            "Error deleting subject."
        );
    }
}


// LOAD CR LIST

async function loadCrList() {

    if (
        !profile ||
        profile.role !==
            "admin"
    ) {
        return;
    }


    try {

        const snapshot =
            await getDocs(
                collection(
                    db,
                    "users"
                )
            );


        crManageList.innerHTML =
            "";


        snapshot.docs

            .filter(
                docSnap =>
                    docSnap.data()
                        .role ===
                    "cr"
            )

            .forEach(
                function (docSnap) {

                    const data =
                        docSnap.data();


                    const row =
                        document.createElement(
                            "div"
                        );


                    row.className =
                        "chip-row";


                    row.innerHTML = `

                        <span>
                            ${docSnap.id}
                            —
                            ${sectionLabelOf(
                                data.section
                            )}
                        </span>


                        <button
                            type="button"
                            data-email="${docSnap.id}">

                            Remove

                        </button>

                    `;


                    crManageList.appendChild(
                        row
                    );
                }
            );


        document
            .querySelectorAll(
                "#crManageList button"
            )
            .forEach(
                function (button) {

                    button.addEventListener(
                        "click",
                        function () {

                            deleteCr(
                                button.dataset.email
                            );
                        }
                    );
                }
            );


    } catch (error) {

        console.error(error);
    }
}


// ASSIGN CR

addCrForm.addEventListener(
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


        const section =
            newCrSection.value;


        if (
            !email ||
            !section
        ) {
            return;
        }


        try {

            await setDoc(

                doc(
                    db,
                    "users",
                    email
                ),

                {
                    role:
                        "cr",

                    section:
                        section,

                    createdAt:
                        new Date()
                            .toISOString()
                }
            );


            newCrEmail.value =
                "";


            await loadCrList();

            logActivity("assign_cr", `${email} assigned as CR for ${sectionLabelOf(section)}`, section);


            alert(
                `${email} is assigned as CR for ${sectionLabelOf(section)}.\n` +
                `Make sure their Firebase Auth login uses this exact (lowercase) email.`
            );


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

        const crProfileSnap =
            await getDoc(
                doc(db, "users", email)
            );

        const crSection =
            crProfileSnap.exists()
                ? crProfileSnap.data().section
                : null;

        await deleteDoc(

            doc(
                db,
                "users",
                email
            )
        );


        await loadCrList();

        logActivity("remove_cr", `${email} removed as CR${crSection ? " of " + sectionLabelOf(crSection) : ""}`, crSection);


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


        displayStudents();


    } catch (error) {

        console.error(error);


        alert(
            "Error loading attendance."
        );
    }
}


// SUBJECT CHANGE

subjectSelect.addEventListener(
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
    }
);


// DATE CHANGE

attendanceDate.addEventListener(
    "change",
    function () {

        if (
            subjectSelect.value
        ) {

            loadAttendance();
        }
    }
);


// SAVE ATTENDANCE

saveBtn.addEventListener(
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


            alert(
                "✅ Attendance saved successfully!"
            );


            await loadRecords();

            logActivity("save_attendance", `${subject} — ${date} — ${sectionLabelOf(activeSection)}`);


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


        const serialByQid =
            new Map();


        data.students.forEach(
            function (
                student,
                index
            ) {

                serialByQid.set(
                    student.qid,
                    index + 1
                );
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
                    serialByQid.get(
                        student.qid
                    );


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


        displayStudents();


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


    const recordInfo = recordsCache[recordId];

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


        await loadRecords();

        logActivity(
            "delete_attendance",
            recordInfo ? `${recordInfo.subject} — ${recordInfo.date} — ${sectionLabelOf(activeSection)}` : recordId
        );


    } catch (error) {

        console.error(error);


        alert(
            "Error deleting record."
        );
    }
}


// REFRESH

refreshRecordsBtn.addEventListener(
    "click",
    loadRecords
);


// DOWNLOAD XLSX

downloadBtn.addEventListener(
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


        const serialByQid =
            new Map();


        students.forEach(
            function (
                student,
                index
            ) {

                serialByQid.set(
                    student.qid,
                    index + 1
                );
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
                    serialByQid.get(
                        student.qid
                    );


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
