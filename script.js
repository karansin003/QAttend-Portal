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
    limit,
    where
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
let requestFilter = "all";


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

            // Record successful login. Admin waits for this one small write so the dashboard can show it immediately.
            const loginLogPromise = writeActivityLog("LOGIN", "User signed in to QAttend.");


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

            await loginLogPromise;
            showAdminDashboard();

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

            showAttendanceView();

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
        }
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

    const manageSection = document.getElementById("manageSectionSelect");
    if (manageSection) {
        manageSection.innerHTML =
            `<option value="">-- Select Section --</option>`;
    }


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

            if (manageSection) {
                const option3 = document.createElement("option");
                option3.value = section.id;
                option3.textContent = section.label;
                manageSection.appendChild(option3);
            }
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
                    targetSection,
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

        const deletedSubject = subjects.find(item => item.id === subjectId);

        await deleteDoc(

            doc(
                db,
                "sections",
                activeSection,
                "subjects",
                subjectId
            )
        );


        void writeActivityLog("DELETE_SUBJECT",deletedSubject?.name || "",{section:activeSection});
        await loadSubjects();


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
            document.getElementById("manageSectionSelect")?.value ||
            newCrSection.value;

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


            managementSection = section;
            await loadCrListForManagementSection();
            void writeActivityLog("ASSIGN_CR",`${email} · ${sectionLabelOf(section)}`);


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

        await deleteDoc(

            doc(
                db,
                "users",
                email
            )
        );


        void writeActivityLog("REMOVE_CR",email);
        await loadCrList();


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
        syncAttendanceGate();
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
        syncAttendanceGate();
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
const activityLogList = document.getElementById("activityLogList");
const dashboardRequestList = document.getElementById("dashboardRequestList");
const requestModal = document.getElementById("requestModal");
const supportModal = document.getElementById("supportModal");
const manageSectionSelect = document.getElementById("manageSectionSelect");
const manageSectionMessage = document.getElementById("manageSectionMessage");
const manageStudentSelect = document.getElementById("manageStudentSelect");
const manageEditStudentBtn = document.getElementById("manageEditStudentBtn");
const manageDeleteStudentBtn = document.getElementById("manageDeleteStudentBtn");

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
    document.querySelectorAll(".welcome-section,.control-section,#adminPanel,#noSectionMessage,#workArea").forEach(el => {
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

    if(profile?.role === "admin" && !sectionReady){
        workArea?.classList.add("hidden");
        noSectionMessage?.classList.remove("hidden");
        return;
    }

    noSectionMessage?.classList.add("hidden");
    if(ready) workArea?.classList.remove("hidden");
    else workArea?.classList.add("hidden");
}

function showAttendanceView(){
    document.body.classList.remove("records-only-view");
    hideMainViews();
    requestsPanel?.classList.add("hidden");
    adminDashboard?.classList.add("hidden");

    // Admin attendance should only show the attendance controls.
    // CR keeps the direct attendance view with the same core controls.
    const welcome = document.querySelector(".welcome-section");
    const control = document.querySelector(".control-section");

    if(profile?.role === "admin"){
        welcome?.classList.add("hidden");
    }else{
        welcome?.classList.remove("hidden");
    }

    control?.classList.remove("hidden");
    adminPanel?.classList.add("hidden");

    if(profile?.role === "admin"){
        activeSection = sectionSelect?.value || null;
    }

    syncAttendanceGate();
}

async function loadAdminDashboardData(){
    if(profile?.role !== "admin") return;

    try{
        const usersSnap = await getDocs(collection(db,"users"));
        const crCount = usersSnap.docs.filter(d => d.data()?.role === "cr").length;

        const sectionTasks = SECTIONS.map(async section => {
            const [studentSnap, subjectSnap] = await Promise.all([
                getDocs(collection(db,"sections",section.id,"students")),
                getDocs(collection(db,"sections",section.id,"subjects"))
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

        await loadAdminActivityLogs();
        await loadAdminRequestPreview();
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

    populateSectionDropdowns();

    if(manageSectionSelect){
        manageSectionSelect.value = managementSection || "";
    }

    if(managementSection){
        activeSection = managementSection;
        void refreshManageData();
    }else{
        renderManageStudentSelect([]);
        subjectManageList.innerHTML = "";
        crManageList.innerHTML = "";
        if(manageSectionMessage) manageSectionMessage.textContent = "Please select a section to manage its data.";
    }
}

function showPreviousRecords(){
    if(!profile) return;
    document.body.classList.add("records-only-view");
    hideMainViews();
    requestsPanel?.classList.add("hidden");
    adminDashboard?.classList.add("hidden");
    document.querySelector(".welcome-section")?.classList.add("hidden");
    document.querySelector(".control-section")?.classList.add("hidden");
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

function logoutFromApp(){
    return writeActivityLog("LOGOUT","User signed out of QAttend.")
        .catch(() => {})
        .then(() => signOut(auth));
}

function populateManageStudentSelect(list){
    if(!manageStudentSelect) return;
    manageStudentSelect.innerHTML = `<option value="">-- Select Student --</option>`;
    list.forEach(student => {
        const option = document.createElement("option");
        option.value = student.id;
        option.textContent = `${student.name} — ${student.qid}`;
        manageStudentSelect.appendChild(option);
    });
}

function renderManageStudentSelect(list){
    populateManageStudentSelect(list);
    if(manageEditStudentBtn) manageEditStudentBtn.disabled = !list.length;
    if(manageDeleteStudentBtn) manageDeleteStudentBtn.disabled = !list.length;
}

async function refreshManageData(){
    if(profile?.role !== "admin" || !managementSection) return;

    activeSection = managementSection;
    sectionLabel.textContent = sectionLabelOf(managementSection);
    if(manageSectionMessage) manageSectionMessage.textContent = `${sectionLabelOf(managementSection)} selected.`;

    try{
        const studentSnap = await getDocs(collection(db,"sections",managementSection,"students"));
        const manageStudents = studentSnap.docs.map(d => ({
            id:d.id,
            qid:d.data()?.qid || "",
            name:d.data()?.name || ""
        })).sort((a,b) => compareNames(a.name,b.name));
        renderManageStudentSelect(manageStudents);

        const subjectSnap = await getDocs(collection(db,"sections",managementSection,"subjects"));
        subjects = subjectSnap.docs.map(d => ({id:d.id,...d.data()})).sort((a,b) => String(a.name).localeCompare(String(b.name)));
        renderSubjectManageList();

        await loadCrListForManagementSection();

        if(newCrSection){
            newCrSection.value = managementSection;
            newCrSection.disabled = true;
        }
    }catch(e){
        console.error(e);
        if(manageSectionMessage) manageSectionMessage.textContent = "Could not load this section. Try again.";
    }
}

async function loadCrListForManagementSection(){
    if(profile?.role !== "admin" || !managementSection || !crManageList) return;
    try{
        const snap = await getDocs(collection(db,"users"));
        crManageList.innerHTML = "";
        const crs = snap.docs.filter(d => d.data()?.role === "cr" && d.data()?.section === managementSection);
        if(!crs.length){
            crManageList.innerHTML = `<p class="empty-record">No CR assigned to this section.</p>`;
            return;
        }
        crs.forEach(d => {
            const row=document.createElement("div");
            row.className="chip-row";
            row.innerHTML=`<span>${escapeHtml(d.id)} — ${escapeHtml(sectionLabelOf(managementSection))}</span><button type="button" data-email="${escapeHtml(d.id)}">Remove</button>`;
            row.querySelector("button")?.addEventListener("click",()=>deleteCr(d.id));
            crManageList.appendChild(row);
        });
    }catch(e){
        console.error(e);
        crManageList.innerHTML=`<p class="empty-record">Could not load CRs.</p>`;
    }
}

manageSectionSelect?.addEventListener("change", async () => {
    if(profile?.role !== "admin") return;
    managementSection = manageSectionSelect.value || null;
    if(!managementSection){
        renderManageStudentSelect([]);
        subjectManageList.innerHTML = "";
        crManageList.innerHTML = "";
        if(newCrSection) newCrSection.disabled = false;
        if(manageSectionMessage) manageSectionMessage.textContent = "Please select a section to manage its data.";
        return;
    }
    await refreshManageData();
});

manageEditStudentBtn?.addEventListener("click", async () => {
    if(profile?.role !== "admin" || !managementSection) return alert("Please select a section first.");
    const id = manageStudentSelect.value;
    if(!id) return alert("Please select a student first.");
    const docSnap = await getDoc(doc(db,"sections",managementSection,"students",id));
    if(!docSnap.exists()) return alert("Student not found.");
    await editStudentUsingSection(managementSection,id,docSnap.data());
    await refreshManageData();
});

manageDeleteStudentBtn?.addEventListener("click", async () => {
    if(profile?.role !== "admin" || !managementSection) return alert("Please select a section first.");
    const id = manageStudentSelect.value;
    if(!id) return alert("Please select a student first.");
    const docSnap = await getDoc(doc(db,"sections",managementSection,"students",id));
    if(!docSnap.exists()) return alert("Student not found.");
    const data = docSnap.data();
    if(!confirm(`Delete ${data.name} (${data.qid})?`)) return;
    try{
        await deleteDoc(doc(db,"sections",managementSection,"students",id));
        await writeActivityLog("DELETE_STUDENT",`${data.qid} · ${data.name}`,{section:managementSection});
        await refreshManageData();
        await loadStudentsForCurrentAttendance();
    }catch(e){
        console.error(e);
        alert("Error deleting student.");
    }
});

async function editStudentUsingSection(section,id,data){
    const newQid = prompt("Enter Q.ID:", data.qid || "");
    if(newQid === null) return;
    const newName = prompt("Enter student name:", data.name || "");
    if(newName === null) return;
    const qid = newQid.trim();
    const name = newName.trim().toUpperCase();
    if(!qid || !name) return alert("Q.ID and student name cannot be empty.");
    try{
        const snap = await getDocs(collection(db,"sections",section,"students"));
        const duplicate = snap.docs.some(d => d.id !== id && String(d.data()?.qid || "").toLowerCase() === qid.toLowerCase());
        if(duplicate) return alert("A student with this Q.ID already exists in this section.");
        await updateDoc(doc(db,"sections",section,"students",id),{qid,name});
        await writeActivityLog("EDIT_STUDENT",`${qid} · ${name}`,{section});
    }catch(e){
        console.error(e);
        alert("Error updating student.");
    }
}

async function loadStudentsForCurrentAttendance(){
    if(!activeSection) return;
    try{
        const snapshot = await getDocs(collection(db,"sections",activeSection,"students"));
        students = snapshot.docs.map(d => ({
            id:d.id,
            qid:d.data()?.qid || "",
            name:d.data()?.name || "",
            status:null
        })).sort((a,b)=>compareNames(a.name,b.name));
        displayStudents();
        headerTotalStudents.textContent = students.length;
    }catch(e){console.error(e);}
}

document.querySelectorAll(".drawer-item").forEach(btn => {
    btn.addEventListener("click", async () => {
        const feature = btn.dataset.feature;
        closeFeatureDrawerFn();

        if(feature === "dashboard" && profile?.role === "admin") showAdminDashboard();
        if(feature === "attendance") showAttendanceView();
        if(feature === "requests") showRequestsPanel();
        if(feature === "manage" && profile?.role === "admin") showManageData();
        if(feature === "records" && profile?.role === "cr") showPreviousRecords();
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
        REQUEST_REJECTED:"Request rejected"
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
    if(!user) return;
    try{
        await addDoc(collection(db,"activityLogs"),{
            action,
            details,
            userEmail:String(user.email || ""),
            userRole:profile?.role || "unknown",
            createdAt:new Date().toISOString(),
            ...extra
        });
    }catch(e){
        console.warn("Activity log failed:",e);
    }
}

async function loadAdminActivityLogs(){
    if(profile?.role !== "admin" || !activityLogList) return;
    activityLogList.innerHTML = `<p class="empty-record">Loading activity...</p>`;
    try{
        // Client-side sort avoids a Firestore composite/index dependency.
        const snap = await getDocs(collection(db,"activityLogs"));
        const rows = snap.docs
            .map(d => ({id:d.id,...d.data()}))
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
        activityLogList.innerHTML = `<p class="empty-record">Could not load activity logs. Publish the latest Firestore Rules.</p>`;
    }
}

async function loadAdminRequestPreview(){
    if(profile?.role !== "admin" || !dashboardRequestList) return;
    try{
        const snap=await getDocs(collection(db,"requests"));
        const rows=snap.docs
            .map(d=>({id:d.id,...d.data()}))
            .filter(x=>x.status==="pending")
            .sort((a,b)=>String(b.createdAt||"").localeCompare(String(a.createdAt||"")))
            .slice(0,5);

        dashboardRequestList.innerHTML="";
        if(!rows.length){
            dashboardRequestList.innerHTML=`<p class="empty-record">No pending requests.</p>`;
            return;
        }

        rows.forEach(r=>{
            const el=document.createElement("div");
            el.className="activity-item";
            const title = r.type?.includes("student")
                ? (r.type==="add_student"?"Add Student":"Delete Student")
                : (r.type==="add_subject"?"Add Subject":"Delete Subject");
            el.innerHTML=`<span class="activity-icon">↗</span><div><b>${escapeHtml(title||"Request")}</b><small>${escapeHtml(r.requestedBy||"")} · ${escapeHtml(r.sectionLabel||sectionLabelOf(r.section||""))}</small><small>${escapeHtml(r.reason||"")}</small></div>`;
            dashboardRequestList.appendChild(el);
        });
    }catch(e){
        console.error(e);
        dashboardRequestList.innerHTML=`<p class="empty-record">Could not load requests. Publish the latest Firestore Rules.</p>`;
    }
}

document.getElementById("refreshLogsBtn")?.addEventListener("click",loadAdminActivityLogs);
document.getElementById("openLogsBtn")?.addEventListener("click",loadAdminActivityLogs);
document.getElementById("openRequestsFromDashboard")?.addEventListener("click",showRequestsPanel);
document.getElementById("backFromRequestsBtn")?.addEventListener("click",()=> profile?.role === "admin" ? showAdminDashboard() : showAttendanceView());

async function loadRequestCenter(){
    if(!requestList || !profile) return;

    requestList.innerHTML=`<p class="empty-record">Loading requests...</p>`;

    try{
        let rows;
        if(profile.role === "cr"){
            const email = String(auth.currentUser?.email || "");
            const snap = await getDocs(
                query(
                    collection(db,"requests"),
                    where("requestedBy","==",email)
                )
            );
            rows = snap.docs.map(d=>({id:d.id,...d.data()}));
        }else{
            const snap = await getDocs(collection(db,"requests"));
            rows = snap.docs.map(d=>({id:d.id,...d.data()}));
        }

        rows.sort((a,b)=>String(b.createdAt||"").localeCompare(String(a.createdAt||"")));

        const filtered=rows.filter(r=>{
            if(profile.role !== "admin") return true;
            if(requestFilter==="pending") return r.status==="pending";
            if(requestFilter==="student") return String(r.type||"").includes("student");
            if(requestFilter==="subject") return String(r.type||"").includes("subject");
            return true;
        });

        requestList.innerHTML="";

        if(profile.role === "cr"){
            const intro=document.createElement("div");
            intro.className="request-card request-intro-card";
            intro.innerHTML=`<div class="request-top"><div class="request-main"><b>Send a change request to Admin</b><small>Your request will be reviewed by an Admin. You do not receive Admin access.</small></div><span class="status-pill status-pending">REQUEST</span></div>`;
            requestList.appendChild(intro);
        }

        if(!filtered.length){
            const empty=document.createElement("p");
            empty.className="empty-record";
            empty.textContent=profile.role === "cr" ? "No requests submitted yet." : "No requests found.";
            requestList.appendChild(empty);
        }else{
            filtered.forEach(r=>{
                const card=document.createElement("div");
                card.className="request-card";
                const title = r.type === "add_student" ? "Add Student" : r.type === "delete_student" ? "Delete Student" : r.type === "add_subject" ? "Add Subject" : "Delete Subject";
                const subjectOrStudent = r.type.includes("student") ? `${r.qid||""}${r.name?` · ${r.name}`:""}` : (r.subject||"");
                card.innerHTML=`<div class="request-top"><div class="request-main"><b>${escapeHtml(title)}${subjectOrStudent?` · ${escapeHtml(subjectOrStudent)}`:""}</b><small>${profile.role === "admin" ? `CR: ${escapeHtml(r.requestedBy||"")} · ` : ""}${escapeHtml(r.sectionLabel||sectionLabelOf(r.section||""))}</small><small>${escapeHtml(r.reason||"")}</small><small>${escapeHtml(r.createdAt?new Date(r.createdAt).toLocaleString():"")}</small></div><span class="status-pill status-${escapeHtml(r.status||"pending")}">${escapeHtml(String(r.status||"pending").toUpperCase())}</span></div>`;

                if(profile.role === "admin" && r.status === "pending"){
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

                requestList.appendChild(card);
            });
        }
    }catch(e){
        console.error(e);
        requestList.innerHTML=`<p class="empty-record">Could not load requests. Publish the latest Firestore Rules.</p>`;
    }

    if(profile.role === "admin") await loadAdminRequestPreview();
}

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

function openRequestModal(){
    if(profile?.role !== "cr") return;
    populateRequestStudentOptions();
    populateRequestSubjectOptions();
    requestModal?.classList.remove("hidden");
    const type=document.getElementById("requestType");
    if(type) type.value="add_student";
    syncRequestForm();
}

function syncRequestForm(){
    const type=document.getElementById("requestType")?.value || "add_student";
    const student=type.includes("student");
    const addStudent=type === "add_student";
    const deleteStudent=type === "delete_student";
    const addSubject=type === "add_subject";
    const deleteSubject=type === "delete_subject";

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
document.getElementById("crSendRequestBtn")?.addEventListener("click",openRequestModal);
document.getElementById("closeRequestModal")?.addEventListener("click",()=>requestModal?.classList.add("hidden"));
document.getElementById("cancelRequestBtn")?.addEventListener("click",()=>requestModal?.classList.add("hidden"));
requestModal?.addEventListener("click",e=>{
    if(e.target===requestModal) requestModal.classList.add("hidden");
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
        requestedBy:String(auth.currentUser.email||""),
        status:"pending",
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
        await writeActivityLog("REQUEST_SUBMITTED",`${friendlyRequestType(type)} · ${data.qid||data.subject||""}`,{requestType:type,section:activeSection});
        requestModal?.classList.add("hidden");
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
        delete_subject:"Delete Subject"
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

async function processChangeRequest(request,approve){
    if(profile?.role!=="admin") return;

    try{
        if(approve){
            if(request.type === "add_student"){
                if(!request.qid || !request.name) throw new Error("Student details are incomplete.");
                const existing=await findStudentByQid(request.section,request.qid);
                if(existing) throw new Error("A student with this Q.ID already exists.");
                await addDoc(collection(db,"sections",request.section,"students"),{
                    qid:request.qid,
                    name:String(request.name||"").toUpperCase()
                });
            }

            if(request.type === "delete_student"){
                const d=await findStudentByQid(request.section,request.qid);
                if(!d) throw new Error("Student not found.");
                await deleteDoc(d.ref);
            }

            if(request.type === "add_subject"){
                const existing=await findSubjectByName(request.section,request.subject);
                if(existing) throw new Error("This subject already exists in the section.");
                await addDoc(collection(db,"sections",request.section,"subjects"),{
                    name:request.subject
                });
            }

            if(request.type === "delete_subject"){
                const d=await findSubjectByName(request.section,request.subject);
                if(!d) throw new Error("Subject not found.");
                await deleteDoc(d.ref);
            }
        }

        await updateDoc(doc(db,"requests",request.id),{
            status:approve ? "approved" : "rejected",
            reviewedBy:String(auth.currentUser?.email||"").toLowerCase(),
            reviewedAt:new Date().toISOString()
        });

        await writeActivityLog(
            approve ? "REQUEST_APPROVED" : "REQUEST_REJECTED",
            `${friendlyRequestType(request.type)} · ${request.qid || request.subject || ""}`,
            {requestId:request.id,section:request.section}
        );

        if(approve){
            if(request.type === "add_student"){
                await writeActivityLog("ADD_STUDENT",`${request.qid||""} · ${request.name||""}`,{section:request.section});
            }
            if(request.type === "delete_student"){
                await writeActivityLog("DELETE_STUDENT",`${request.qid||""} · ${request.name||""}`,{section:request.section});
            }
            if(request.type === "add_subject"){
                await writeActivityLog("ADD_SUBJECT",request.subject||"",{section:request.section});
            }
            if(request.type === "delete_subject"){
                await writeActivityLog("DELETE_SUBJECT",request.subject||"",{section:request.section});
            }
        }

        await loadRequestCenter();
        await loadAdminDashboardData();
        alert(approve ? "✅ Request approved and change applied." : "Request rejected.");
    }catch(e){
        console.error(e);
        alert(`Could not ${approve ? "approve" : "reject"} request: ${e.message || e}`);
    }
}

// Keep admin dashboard as the first screen for Admin, and direct attendance for CR.
// Admin dashboard statistics and logs load in the background so login remains responsive.
// CR receives only the assigned section and does not get section-switch access.

