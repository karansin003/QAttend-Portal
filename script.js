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

// Where "Contact Admin" and "Request a new section" emails are sent.
// >>> CHANGE THIS to the real admin's email before deploying. <<<
const ADMIN_CONTACT_EMAIL = "sonusin8672@gmail.com";


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
let selectedRequestIds = new Set();


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

const contactAdminLink = document.getElementById("contactAdminLink");
const contactAdminModal = document.getElementById("contactAdminModal");
const closeContactAdminModal = document.getElementById("closeContactAdminModal");
const cancelContactAdminBtn = document.getElementById("cancelContactAdminBtn");
const contactAdminForm = document.getElementById("contactAdminForm");

const requestSectionLink = document.getElementById("requestSectionLink");
const addSectionRequestModal = document.getElementById("addSectionRequestModal");
const closeAddSectionRequestModal = document.getElementById("closeAddSectionRequestModal");
const cancelAddSectionRequestBtn = document.getElementById("cancelAddSectionRequestBtn");
const addSectionRequestForm = document.getElementById("addSectionRequestForm");


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


// CONTACT ADMIN & REQUEST NEW SECTION (login page, no account needed)
// Both build a pre-filled `mailto:` link and hand off to the user's own
// mail app — no backend/Firestore write needed, and it works before login.

function openModal(modal) {
    modal.classList.remove("hidden");
}

function closeModal(modal) {
    modal.classList.add("hidden");
}

contactAdminLink.addEventListener("click", function (event) {
    event.preventDefault();
    openModal(contactAdminModal);
});

closeContactAdminModal.addEventListener("click", () => closeModal(contactAdminModal));
cancelContactAdminBtn.addEventListener("click", () => closeModal(contactAdminModal));
contactAdminModal.addEventListener("click", function (event) {
    if (event.target === contactAdminModal) closeModal(contactAdminModal);
});

contactAdminForm.addEventListener("submit", async function (event) {
    event.preventDefault();

    const name = document.getElementById("contactName").value.trim();
    const email = document.getElementById("contactEmail").value.trim();
    const mobile = document.getElementById("contactMobile").value.trim();
    const type = document.getElementById("contactType").value;
    const message = document.getElementById("contactMessage").value.trim();

    const subject = `QAttend — ${type} (${name})`;
    const body =
        `Name: ${name}\n` +
        `Email: ${email}\n` +
        `Mobile: ${mobile}\n` +
        `Request Type: ${type}\n\n` +
        `Message:\n${message}`;

    try {
        await addDoc(collection(db, "requests"), {
            type: "contact_admin",
            status: "pending",
            requestedBy: email.toLowerCase(),
            requestedByRole: "guest",
            name, email, mobile, requestType: type, reason: message,
            createdAt: new Date().toISOString()
        });
    } catch (firestoreError) {
        console.warn("Contact request could not be saved to Firestore:", firestoreError);
    }

    window.location.href =
        `mailto:${ADMIN_CONTACT_EMAIL}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

    closeModal(contactAdminModal);
    contactAdminForm.reset();
});

requestSectionLink.addEventListener("click", function (event) {
    event.preventDefault();
    resetSectionRequestForm();
    openModal(addSectionRequestModal);
});

closeAddSectionRequestModal.addEventListener("click", () => closeModal(addSectionRequestModal));
cancelAddSectionRequestBtn.addEventListener("click", () => closeModal(addSectionRequestModal));
addSectionRequestModal.addEventListener("click", function (event) {
    if (event.target === addSectionRequestModal) closeModal(addSectionRequestModal);
});

let sectionExcelPayload = null;

function resetSectionRequestForm(){
    addSectionRequestForm.reset();
    sectionExcelPayload = null;
    const preview = document.getElementById("sectionExcelPreview");
    preview?.classList.add("hidden");
    document.getElementById("sectionExcelFileLabel").textContent = "Click to upload Excel file";
    document.getElementById("sectionRequestMessage").textContent = "";
    ["excelStudentCount","excelSubjectCount","excelDuplicateCount","excelInvalidCount"].forEach(id=>{
        const el=document.getElementById(id); if(el) el.textContent="0";
    });
}

function normalizeExcelValue(value){
    return String(value ?? "").trim();
}

async function parseSectionExcel(file){
    await ensureXLSXLoaded();
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, {type:"array"});
    const sheetName = workbook.SheetNames[0];
    if(!sheetName) throw new Error("No worksheet found in the Excel file.");

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {header:1, defval:""});
    if(!rows.length) throw new Error("The Excel file is empty.");

    const headers = rows[0].map(normalizeExcelValue).map(x=>x.toLowerCase());
    const qidIndex = headers.findIndex(x=>x === "qid" || x === "q.id" || x === "q id");
    const nameIndex = headers.findIndex(x=>x === "student name" || x === "student name ");
    const subjectIndexes = headers.map((x,i)=>({x,i})).filter(({x})=>/^subject\s*\d+$/i.test(x));

    if(qidIndex < 0 || nameIndex < 0 || subjectIndexes.length === 0){
        throw new Error("Invalid template. Required columns: Qid, Student Name, Subject1, Subject2, Subject3...");
    }

    const students=[];
    const subjectSet=new Set();
    const seenQids=new Set();
    let duplicateCount=0;
    let invalidCount=0;

    for(let r=1;r<rows.length;r++){
        const row=rows[r] || [];
        const qid=normalizeExcelValue(row[qidIndex]);
        const name=normalizeExcelValue(row[nameIndex]);
        const subjects=subjectIndexes.map(({i})=>normalizeExcelValue(row[i])).filter(Boolean);
        const rowEmpty=!qid && !name && subjects.length===0;
        if(rowEmpty) continue;
        if(!qid || !name){ invalidCount++; continue; }
        const qidKey=qid.toLowerCase();
        if(seenQids.has(qidKey)){ duplicateCount++; continue; }
        seenQids.add(qidKey);
        students.push({qid,name:name.toUpperCase()});
        subjects.forEach(s=>subjectSet.add(s));
    }

    return {
        students,
        subjects:[...subjectSet].sort((a,b)=>a.localeCompare(b)),
        duplicateCount,
        invalidCount
    };
}

document.getElementById("sectionExcelFile")?.addEventListener("change", async function(){
    const file=this.files?.[0];
    if(!file) return;
    const label=document.getElementById("sectionExcelFileLabel");
    const message=document.getElementById("sectionRequestMessage");
    label.textContent=file.name;
    message.textContent="Reading Excel file...";
    try{
        sectionExcelPayload=await parseSectionExcel(file);
        document.getElementById("sectionExcelPreview")?.classList.remove("hidden");
        document.getElementById("excelStudentCount").textContent=sectionExcelPayload.students.length;
        document.getElementById("excelSubjectCount").textContent=sectionExcelPayload.subjects.length;
        document.getElementById("excelDuplicateCount").textContent=sectionExcelPayload.duplicateCount;
        document.getElementById("excelInvalidCount").textContent=sectionExcelPayload.invalidCount;
        message.textContent=sectionExcelPayload.students.length
            ? "Excel validated. You can submit the request."
            : "No valid students were found in the Excel file.";
        message.style.color = sectionExcelPayload.students.length ? "var(--green)" : "var(--red)";
    }catch(err){
        sectionExcelPayload=null;
        document.getElementById("sectionExcelPreview")?.classList.add("hidden");
        message.textContent=err.message || "Could not read the Excel file.";
        message.style.color="var(--red)";
    }
});

addSectionRequestForm.addEventListener("submit", async function(event){
    event.preventDefault();
    const message=document.getElementById("sectionRequestMessage");
    const submitButton=document.getElementById("submitAddSectionRequestBtn");

    const crQid=document.getElementById("reqCrQid").value.trim();
    const crName=document.getElementById("reqCrName").value.trim();
    const crMobile=document.getElementById("reqCrMobile").value.trim();
    const crEmail=document.getElementById("reqCrEmail").value.trim().toLowerCase();
    const mentorName=document.getElementById("reqMentorName").value.trim();
    const mentorMobile=document.getElementById("reqMentorMobile").value.trim();
    const course=document.getElementById("reqCourse").value.trim();
    const sectionName=document.getElementById("reqSectionName").value.trim();
    const semester=document.getElementById("reqSemester").value.trim();
    const year=document.getElementById("reqYear").value.trim();

    if(!/^\d{10}$/.test(crMobile) || !/^\d{10}$/.test(mentorMobile)){
        message.style.color="var(--red)";
        message.textContent="CR Mobile and Mentor No. must be exactly 10 digits.";
        return;
    }
    if(!sectionExcelPayload || !sectionExcelPayload.students.length){
        message.style.color="var(--red)";
        message.textContent="Please download, fill and upload the Excel template before submitting.";
        return;
    }

    submitButton.disabled=true;
    message.style.color="var(--gray)";
    message.textContent="Sending request to Admin...";

    const data={
        type:"add_section",
        status:"pending",
        requestedBy:crEmail,
        requestedByRole:"cr",
        crQid, crName, crMobile, crEmail,
        mentorName, mentorMobile,
        course, section:sectionName,
        sectionLabel:sectionName,
        semester, year,
        students:sectionExcelPayload.students,
        subjects:sectionExcelPayload.subjects,
        excelStats:{
            students:sectionExcelPayload.students.length,
            subjects:sectionExcelPayload.subjects.length,
            duplicates:sectionExcelPayload.duplicateCount,
            invalid:sectionExcelPayload.invalidCount
        },
        createdAt:new Date().toISOString()
    };

    try{
        await addDoc(collection(db,"requests"),data);

        // Also prepare an email so the request can be sent directly from the user's mail app.
        const subject=`QAttend — New Section Request: ${sectionName}`;
        const body=
            `CR Q.ID: ${crQid}\nCR Name: ${crName}\nCR Mobile: ${crMobile}\nCR Email: ${crEmail}\n\n`+
            `Mentor Name: ${mentorName}\nMentor No.: ${mentorMobile}\n\n`+
            `Course: ${course}\nSection: ${sectionName}\nSemester: ${semester}\nYear: ${year}\n\n`+
            `Students: ${data.students.length}\nSubjects: ${data.subjects.join(", ")}\n\n`+
            `The completed Excel data has been uploaded with this request in QAttend. Please review it from the Admin Request Center.`;

        message.style.color="var(--green)";
        message.textContent="Request sent successfully. Admin will review it.";
        window.setTimeout(()=>{
            window.location.href=`mailto:${ADMIN_CONTACT_EMAIL}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
        },150);
        window.setTimeout(()=>{
            closeModal(addSectionRequestModal);
            resetSectionRequestForm();
        },350);
    }catch(err){
        console.error(err);
        message.style.color="var(--red)";
        message.textContent=`Could not send request: ${err.message || err}. Publish the latest Firestore Rules first.`;
    }finally{
        submitButton.disabled=false;
    }
});

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

            // Record the successful login. Logging failures must never block the app.
            await writeActivityLog("LOGIN", "User signed in to QAttend.");


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

            // Admin lands on the Admin Dashboard after login.
            // Mark Attendance is opened explicitly from the hamburger menu.
            populateSectionDropdowns();

            activeSection = null;
            managementSection = null;

            sectionSelect.value = "";
            if (manageSectionSelect) manageSectionSelect.value = "";

            showAdminDashboard();

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


    if (newCrSection) {
        newCrSection.innerHTML = "";
    }

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


            if (newCrSection) {
                newCrSection.appendChild(option2);
            }

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
            await writeActivityLog(
                "ASSIGN_CR",
                `${email} · ${sectionLabelOf(section)}`,
                {section}
            );


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

    populateSectionDropdowns();

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
        activityLogList.innerHTML = `<p class="empty-record">Could not load activity logs. Make sure the latest Firestore Rules are published.</p>`;
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
        dashboardRequestList.innerHTML=`<p class="empty-record">Could not load requests. Make sure the latest Firestore Rules are published.</p>`;
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
            return true;
        });

        const visiblePendingIds = new Set(filtered.filter(r=>profile.role === "admin" && r.status === "pending").map(r=>r.id));
        selectedRequestIds = new Set([...selectedRequestIds].filter(id=>visiblePendingIds.has(id)));
        updateBulkRequestToolbar(filtered);
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

                requestList.appendChild(card);
            });
        }
    }catch(e){
        console.error(e);
        requestList.innerHTML=`<p class="empty-record">Could not load requests. Make sure the latest Firestore Rules are published.</p>`;
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

async function processChangeRequest(request,approve,options={}){
    const silent=options.silent===true;
    if(profile?.role!=="admin") return false;

    try{
        // IMPORTANT: only an approved request changes the actual students/subjects data.
        // A rejected request only changes the request status, leaving all section data untouched.
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

        // Mark the request only after the requested data operation succeeded.
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

            // Refresh the currently open section so approved student/subject changes
            // become visible immediately without requiring a page reload.
            if(activeSection === request.section){
                if(request.type.includes("student")) await loadStudents();
                if(request.type.includes("subject")) await loadSubjects();
            }
        }

        if(!silent){
            await loadRequestCenter();
            await loadAdminDashboardData();
            alert(approve ? "✅ Request approved and change applied." : "Request rejected.");
        }

        return true;
    }catch(e){
        console.error(e);
        if(!silent){
            alert(`Could not ${approve ? "approve" : "reject"} request: ${e.message || e}`);
        }
        // Let bulk processing count the request as failed instead of falsely reporting success.
        throw e;
    }
}

// Admin opens Mark Attendance by default. Dashboard is explicit from the hamburger menu.
// CR opens its assigned attendance section directly and cannot switch sections.


/* Dynamic section labels for sections created by Admin. */
sectionLabelOf = function(id){
    const found=dynamicSectionCatalog.find(x=>x.id===id);
    if(found) return found.label;
    const fixed=SECTIONS.find(x=>x.id===id);
    return fixed ? fixed.label : id;
};

/* =========================================================
   ADMIN PANEL V3
   Additive admin-only layer. Existing login, CR and attendance UI is kept.
   ========================================================= */

const adminSectionsPanel = document.getElementById("adminSectionsPanel");
const adminUsersPanel = document.getElementById("adminUsersPanel");
const adminReportsPanel = document.getElementById("adminReportsPanel");
const adminSectionModal = document.getElementById("adminSectionModal");
const adminUserModal = document.getElementById("adminUserModal");
const adminRequestDetailModal = document.getElementById("adminRequestDetailModal");
const adminRequestDetailBody = document.getElementById("adminRequestDetailBody");
let adminDetailRequest = null;
let dynamicSectionCatalog = [];

function adminSlug(value){
    return String(value || "")
        .trim()
        .toUpperCase()
        .replace(/[^A-Z0-9]+/g,"-")
        .replace(/^-+|-+$/g,"")
        .slice(0,80);
}

function adminSectionLabel(section){
    const found = dynamicSectionCatalog.find(x => x.id === section);
    return found?.label || sectionLabelOf(section);
}

function getAdminSections(){
    const base = SECTIONS.map(s => ({...s}));
    const map = new Map(base.map(s => [s.id,s]));
    dynamicSectionCatalog.forEach(s => map.set(s.id,s));
    return [...map.values()];
}

async function loadDynamicSectionCatalog(){
    if(profile?.role !== "admin") return getAdminSections();
    try{
        const snap = await getDocs(collection(db,"sections"));
        snap.docs.forEach(d => {
            const data=d.data()||{};
            if(!d.id) return;
            const item={
                id:d.id,
                label:String(data.label || data.section || d.id),
                course:String(data.course || ""),
                semester:String(data.semester || ""),
                year:String(data.year || ""),
                mentorName:String(data.mentorName || ""),
                mentorMobile:String(data.mentorMobile || "")
            };
            const existing=dynamicSectionCatalog.findIndex(x=>x.id===item.id);
            if(existing>=0) dynamicSectionCatalog[existing]=item;
            else dynamicSectionCatalog.push(item);
        });
    }catch(e){
        console.warn("Could not load dynamic section catalog:",e);
    }
    return getAdminSections();
}

function fillAdminSectionSelect(select, includeNone=true){
    if(!select) return;
    const current=select.value;
    select.innerHTML=includeNone ? `<option value="">-- Select Section --</option>` : "";
    getAdminSections().forEach(section=>{
        const o=document.createElement("option");
        o.value=section.id;
        o.textContent=section.label;
        select.appendChild(o);
    });
    if(current && [...select.options].some(o=>o.value===current)) select.value=current;
}

async function refreshAllAdminSectionSelects(){
    await loadDynamicSectionCatalog();
    fillAdminSectionSelect(sectionSelect,true);
    fillAdminSectionSelect(manageSectionSelect,true);
    fillAdminSectionSelect(document.getElementById("adminUserSection"),true);
}

function adminFriendlyRequestType(type){
    return ({
        add_student:"Add Student",
        delete_student:"Delete Student",
        add_subject:"Add Subject",
        delete_subject:"Delete Subject",
        add_section:"New Section",
        contact_admin:"Contact Admin"
    })[type] || "Request";
}

function adminRequestTarget(r){
    if(r.type?.includes("student")) return r.qid ? `${r.qid}${r.name ? ` · ${r.name}` : ""}` : "Student change";
    if(r.type?.includes("subject")) return r.subject || "Subject change";
    if(r.type === "add_section") return `${r.course || "Course"} · ${r.section || r.sectionLabel || "New Section"}`;
    if(r.type === "contact_admin") return r.requestType || "Access / support";
    return "Request";
}

function adminRequestStatusClass(status){
    return `status-pill status-${escapeHtml(status || "pending")}`;
}

function hideAdminPanels(){
    [adminSectionsPanel,adminUsersPanel,adminReportsPanel].forEach(el=>el?.classList.add("hidden"));
}

function showAdminSections(){
    if(profile?.role!=="admin") return;
    hideMainViews();
    hideAdminPanels();
    adminSectionsPanel?.classList.remove("hidden");
    void loadAdminSectionsTable();
}

function showAdminUsers(){
    if(profile?.role!=="admin") return;
    hideMainViews();
    hideAdminPanels();
    adminUsersPanel?.classList.remove("hidden");
    void loadAdminUsersTable();
}

function showAdminReports(){
    if(profile?.role!=="admin") return;
    hideMainViews();
    hideAdminPanels();
    adminReportsPanel?.classList.remove("hidden");
}

async function loadAdminSectionsTable(){
    if(profile?.role!=="admin") return;
    const body=document.getElementById("adminSectionsTableBody");
    if(!body) return;
    body.innerHTML=`<tr><td colspan="8" class="empty-cell">Loading sections...</td></tr>`;
    const sections=await refreshAllAdminSectionSelects();
    const usersSnap=await getDocs(collection(db,"users"));
    const rows=[];
    for(const s of sections){
        try{
            const [st,sub]=await Promise.all([
                getDocs(collection(db,"sections",s.id,"students")),
                getDocs(collection(db,"sections",s.id,"subjects"))
            ]);
            const cr=usersSnap.docs.find(d=>d.data()?.role==="cr" && String(d.data()?.section||"")===s.id);
            const meta=dynamicSectionCatalog.find(x=>x.id===s.id)||{};
            rows.push({
                ...s,
                ...meta,
                students:st.size,
                subjects:sub.size,
                cr:cr?.id || "(Not Assigned)"
            });
        }catch(e){
            rows.push({...s,students:0,subjects:0,cr:"(Not available)"});
        }
    }
    body.innerHTML="";
    if(!rows.length){body.innerHTML=`<tr><td colspan="8" class="empty-cell">No sections found.</td></tr>`;return;}
    rows.forEach(r=>{
        const tr=document.createElement("tr");
        tr.innerHTML=`<td>${escapeHtml(r.course||"B.Tech")}</td><td>${escapeHtml(r.label||r.id)}</td><td>${escapeHtml(r.semester||"-")}</td><td>${escapeHtml(r.year||"-")}</td><td>${r.students}</td><td>${r.subjects}</td><td>${escapeHtml(r.cr)}</td><td><div class="admin-row-actions"><button class="admin-icon-btn edit" data-edit-section="${escapeHtml(r.id)}">✎</button><button class="admin-icon-btn delete" data-delete-section="${escapeHtml(r.id)}">🗑</button></div></td>`;
        body.appendChild(tr);
    });
    body.querySelectorAll("[data-edit-section]").forEach(b=>b.onclick=()=>openAdminSectionEditor(b.dataset.editSection));
    body.querySelectorAll("[data-delete-section]").forEach(b=>b.onclick=()=>deleteAdminSection(b.dataset.deleteSection));
}

async function openAdminSectionEditor(id=""){
    const section=getAdminSections().find(x=>x.id===id);
    document.getElementById("adminSectionModalTitle").textContent=section ? "Edit Section" : "Add Section";
    document.getElementById("adminSectionCourse").value=section?.course||"";
    document.getElementById("adminSectionName").value=section?.label||"";
    document.getElementById("adminSectionSemester").value=section?.semester||"";
    document.getElementById("adminSectionYear").value=section?.year||"";
    document.getElementById("adminSectionMentor").value=section?.mentorName||"";
    document.getElementById("adminSectionMentorMobile").value=section?.mentorMobile||"";
    adminSectionModal.dataset.editId=id;
    adminSectionModal?.classList.remove("hidden");
}

document.getElementById("adminAddSectionBtn")?.addEventListener("click",()=>openAdminSectionEditor(""));
document.getElementById("closeAdminSectionModal")?.addEventListener("click",()=>adminSectionModal?.classList.add("hidden"));
document.getElementById("cancelAdminSectionBtn")?.addEventListener("click",()=>adminSectionModal?.classList.add("hidden"));

async function deleteAdminSection(id){
    if(profile?.role!=="admin") return;
    if(SECTIONS.some(s=>s.id===id) && !dynamicSectionCatalog.some(s=>s.id===id)){
        alert("This is an existing system section. Use Update / Manage Data instead of deleting it here.");
        return;
    }
    if(!confirm(`Delete section ${adminSectionLabel(id)}? This removes the section profile only. Student/subject/attendance subcollections are not automatically deleted.`)) return;
    try{
        await deleteDoc(doc(db,"sections",id));
        await writeActivityLog("DELETE_SECTION",adminSectionLabel(id),{section:id});
        await loadAdminSectionsTable();
    }catch(e){alert(`Could not delete section: ${e.message||e}`);}
}

document.getElementById("adminSectionForm")?.addEventListener("submit",async e=>{
    e.preventDefault();
    if(profile?.role!=="admin") return;
    const course=document.getElementById("adminSectionCourse").value.trim();
    const label=document.getElementById("adminSectionName").value.trim();
    const semester=document.getElementById("adminSectionSemester").value.trim();
    const year=document.getElementById("adminSectionYear").value.trim();
    const mentorName=document.getElementById("adminSectionMentor").value.trim();
    const mentorMobile=document.getElementById("adminSectionMentorMobile").value.trim();
    const oldId=adminSectionModal.dataset.editId||"";
    const id=oldId || adminSlug(label);
    if(!id) return alert("Enter a section name.");
    try{
        const sectionData={course,label,section:label,semester,year,mentorName,mentorMobile,updatedAt:new Date().toISOString()};
        if(!oldId) sectionData.createdAt=new Date().toISOString();
        await setDoc(doc(db,"sections",id),sectionData,{merge:true});
        if(!oldId) dynamicSectionCatalog.push({id,label,course,semester,year,mentorName,mentorMobile});
        await writeActivityLog(oldId?"EDIT_SECTION":"ADD_SECTION",`${course} · ${label}`,{section:id});
        adminSectionModal?.classList.add("hidden");
        await loadAdminSectionsTable();
        await refreshAllAdminSectionSelects();
    }catch(err){console.error(err);alert(`Could not save section: ${err.message||err}`);}
});

async function loadAdminUsersTable(){
    if(profile?.role!=="admin") return;
    const body=document.getElementById("adminUsersTableBody");
    if(!body) return;
    body.innerHTML=`<tr><td colspan="6" class="empty-cell">Loading users...</td></tr>`;
    try{
        const snap=await getDocs(collection(db,"users"));
        body.innerHTML="";
        snap.docs.sort((a,b)=>String(a.id).localeCompare(String(b.id))).forEach(d=>{
            const x=d.data()||{};
            const tr=document.createElement("tr");
            const active=x.active!==false;
            tr.innerHTML=`<td>${escapeHtml(x.name||d.id)}</td><td>${escapeHtml(d.id)}</td><td>${escapeHtml(x.role||"-")}</td><td>${escapeHtml(adminSectionLabel(x.section||"-"))}</td><td><span class="admin-status ${active?"active":"inactive"}">${active?"Active":"Inactive"}</span></td><td><div class="admin-row-actions"><button class="admin-icon-btn edit" data-edit-user="${escapeHtml(d.id)}">✎</button><button class="admin-icon-btn delete" data-toggle-user="${escapeHtml(d.id)}">${active?"⏸":"▶"}</button></div></td>`;
            body.appendChild(tr);
        });
        if(!snap.size) body.innerHTML=`<tr><td colspan="6" class="empty-cell">No users found.</td></tr>`;
        body.querySelectorAll("[data-edit-user]").forEach(b=>b.onclick=()=>openAdminUserEditor(b.dataset.editUser));
        body.querySelectorAll("[data-toggle-user]").forEach(b=>b.onclick=()=>toggleAdminUser(b.dataset.toggleUser));
    }catch(e){body.innerHTML=`<tr><td colspan="6" class="empty-cell">Could not load users.</td></tr>`;}
}

async function openAdminUserEditor(email=""){
    const role=document.getElementById("adminUserRole");
    const section=document.getElementById("adminUserSection");
    await refreshAllAdminSectionSelects();
    if(email){
        const snap=await getDoc(doc(db,"users",email));
        if(snap.exists()){
            const x=snap.data()||{};
            document.getElementById("adminUserName").value=x.name||"";
            document.getElementById("adminUserEmail").value=email;
            role.value=x.role||"cr";
            section.value=x.section||"";
            document.getElementById("adminUserEmail").disabled=true;
        }
    }else{
        document.getElementById("adminUserForm").reset();
        document.getElementById("adminUserEmail").disabled=false;
    }
    adminUserModal?.classList.remove("hidden");
}

document.getElementById("adminAddUserBtn")?.addEventListener("click",()=>openAdminUserEditor(""));
document.getElementById("closeAdminUserModal")?.addEventListener("click",()=>adminUserModal?.classList.add("hidden"));
document.getElementById("cancelAdminUserBtn")?.addEventListener("click",()=>adminUserModal?.classList.add("hidden"));

document.getElementById("adminUserForm")?.addEventListener("submit",async e=>{
    e.preventDefault();
    if(profile?.role!=="admin") return;
    const email=document.getElementById("adminUserEmail").value.trim().toLowerCase();
    const name=document.getElementById("adminUserName").value.trim();
    const role=document.getElementById("adminUserRole").value;
    const section=document.getElementById("adminUserSection").value||null;
    if(!email||!name) return alert("Name and email are required.");
    try{
        await setDoc(doc(db,"users",email),{name,email,role,section,active:true,updatedAt:new Date().toISOString()},{merge:true});
        await writeActivityLog("MANAGE_USER",`${email} · ${role}`,{section:section||""});
        adminUserModal?.classList.add("hidden");
        await loadAdminUsersTable();
    }catch(err){alert(`Could not save user: ${err.message||err}`);}
});

async function toggleAdminUser(email){
    try{
        const ref=doc(db,"users",email); const snap=await getDoc(ref); if(!snap.exists()) return;
        const active=snap.data()?.active!==false;
        await updateDoc(ref,{active:!active,updatedAt:new Date().toISOString()});
        await writeActivityLog(active?"DEACTIVATE_USER":"ACTIVATE_USER",email);
        await loadAdminUsersTable();
    }catch(e){alert(`Could not update user: ${e.message||e}`);}
}

function renderAdminRequestDetail(r){
    if(!adminRequestDetailBody) return;
    const rows=[];
    const add=(label,val)=>rows.push(`<div><span>${escapeHtml(label)}</span><b>${escapeHtml(val||"-")}</b></div>`);
    add("Request Type",adminFriendlyRequestType(r.type));
    add("Status",String(r.status||"pending").toUpperCase());
    add("Requested By",r.requestedBy||r.email||"");
    add("Submitted",r.createdAt?new Date(r.createdAt).toLocaleString():"");
    if(r.type==="add_section"){
        add("CR Q.ID",r.crQid);add("CR Name",r.crName);add("CR Email",r.crEmail);add("CR Mobile",r.crMobile);add("Course",r.course);add("Section",r.sectionLabel||r.section);add("Semester",r.semester);add("Year",r.year);add("Mentor Name",r.mentorName);add("Mentor Number",r.mentorMobile);add("Students",r.excelStats?.students || r.students?.length || 0);add("Subjects",r.excelStats?.subjects || r.subjects?.length || 0);
    }else if(r.type==="contact_admin"){
        add("Name",r.name);add("Email",r.email);add("Mobile",r.mobile);add("Request",r.requestType);
    }else if(r.type?.includes("student")){
        add("Q.ID",r.qid);add("Student Name",r.name);add("Section",r.sectionLabel||r.section);
    }else if(r.type?.includes("subject")){
        add("Subject",r.subject);add("Section",r.sectionLabel||r.section);
    }
    adminRequestDetailBody.innerHTML=`<div class="admin-detail-section"><div class="admin-detail-title">Request Information</div><div class="admin-detail-grid">${rows.join("")}</div></div><div class="admin-detail-section"><div class="admin-detail-title">Reason / Message</div><div class="admin-detail-message">${escapeHtml(r.reason||r.message||"No reason provided.")}</div></div>`;
    const canAct=r.status==="pending";
    document.getElementById("detailApproveBtn").classList.toggle("hidden",!canAct);
    document.getElementById("detailRejectBtn").classList.toggle("hidden",!canAct);
}

function openAdminRequestDetail(r){
    adminDetailRequest=r;
    renderAdminRequestDetail(r);
    adminRequestDetailModal?.classList.remove("hidden");
}

document.getElementById("closeAdminRequestDetail")?.addEventListener("click",()=>adminRequestDetailModal?.classList.add("hidden"));
adminRequestDetailModal?.addEventListener("click",e=>{if(e.target===adminRequestDetailModal) adminRequestDetailModal.classList.add("hidden")});
document.getElementById("detailApproveBtn")?.addEventListener("click",async()=>{if(adminDetailRequest){await processChangeRequest(adminDetailRequest,true);adminRequestDetailModal?.classList.add("hidden");}});
document.getElementById("detailRejectBtn")?.addEventListener("click",async()=>{if(adminDetailRequest){await processChangeRequest(adminDetailRequest,false);adminRequestDetailModal?.classList.add("hidden");}});

async function processNewSectionApproval(r){
    const id=adminSlug(r.section || r.sectionLabel);
    if(!id) throw new Error("Section name is missing.");
    const existing=await getDoc(doc(db,"sections",id));
    if(existing.exists()) throw new Error("A section with this ID already exists.");
    await setDoc(doc(db,"sections",id),{
        course:r.course||"",
        label:r.sectionLabel||r.section||id,
        section:r.sectionLabel||r.section||id,
        semester:r.semester||"",
        year:r.year||"",
        mentorName:r.mentorName||"",
        mentorMobile:r.mentorMobile||"",
        crQid:r.crQid||"",
        createdAt:new Date().toISOString(),
        updatedAt:new Date().toISOString()
    });
    for(const student of (r.students||[])){
        if(!student?.qid) continue;
        const existingStudent=await findStudentByQid(id,student.qid);
        if(!existingStudent) await addDoc(collection(db,"sections",id,"students"),{qid:String(student.qid),name:String(student.name||"").toUpperCase()});
    }
    for(const subject of (r.subjects||[])){
        if(!subject) continue;
        const existingSubject=await findSubjectByName(id,subject);
        if(!existingSubject) await addDoc(collection(db,"sections",id,"subjects"),{name:String(subject)});
    }
    if(r.crEmail){
        await setDoc(doc(db,"users",String(r.crEmail).toLowerCase()),{
            name:r.crName||"",email:String(r.crEmail).toLowerCase(),role:"cr",section:id,crQid:r.crQid||"",mobile:r.crMobile||"",active:true,updatedAt:new Date().toISOString()
        },{merge:true});
    }
    dynamicSectionCatalog.push({id,label:r.sectionLabel||r.section||id,course:r.course||"",semester:r.semester||"",year:r.year||"",mentorName:r.mentorName||"",mentorMobile:r.mentorMobile||""});
}

/* Override the request processor with New Section + Contact Admin support. */
const originalProcessChangeRequest = processChangeRequest;
processChangeRequest = async function(request,approve,options={}){
    if(profile?.role!=="admin") return false;
    if(approve && request.type==="add_section"){
        try{
            await processNewSectionApproval(request);
            await updateDoc(doc(db,"requests",request.id),{status:"approved",reviewedBy:String(auth.currentUser?.email||"").toLowerCase(),reviewedAt:new Date().toISOString()});
            await writeActivityLog("REQUEST_APPROVED",`New Section · ${request.sectionLabel||request.section||""}`,{requestId:request.id,section:adminSlug(request.section||request.sectionLabel)});
            await writeActivityLog("ADD_SECTION",`${request.course||""} · ${request.sectionLabel||request.section||""}`);
            if(options.silent) return true;
            await loadRequestCenter(); await loadAdminDashboardData(); await refreshAllAdminSectionSelects();
            alert("✅ New section approved and created successfully.");
            return true;
        }catch(e){
            if(!options.silent) alert(`Could not approve new section: ${e.message||e}`);
            throw e;
        }
    }
    if(request.type==="contact_admin"){
        try{
            await updateDoc(doc(db,"requests",request.id),{status:approve?"approved":"rejected",reviewedBy:String(auth.currentUser?.email||"").toLowerCase(),reviewedAt:new Date().toISOString()});
            await writeActivityLog(approve?"REQUEST_APPROVED":"REQUEST_REJECTED",`Contact Admin · ${request.requestedBy||request.email||""}`,{requestId:request.id});
            if(!options.silent){await loadRequestCenter();await loadAdminDashboardData();alert(approve?"✅ Contact request approved.":"Request rejected.");}
            return true;
        }catch(e){if(!options.silent) alert(`Could not process request: ${e.message||e}`);throw e;}
    }
    return await originalProcessChangeRequest(request,approve,options);
};

/* Override Request Center rendering so Admin gets the table-style UI shown in the reference. */
loadRequestCenter = async function(){
    if(!requestList || !profile) return;
    requestList.innerHTML=`<p class="empty-record">Loading requests...</p>`;
    try{
        let snap;
        if(profile.role==="cr"){
            snap=await getDocs(query(collection(db,"requests"),where("requestedBy","==",String(auth.currentUser?.email||"").trim())));
        }else snap=await getDocs(collection(db,"requests"));
        const rows=snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>String(b.createdAt||"").localeCompare(String(a.createdAt||"")));
        window.__qattendRequestCache=Object.fromEntries(rows.map(r=>[r.id,r]));
        let filtered=rows;
        if(profile.role==="admin"){
            if(requestFilter==="pending") filtered=rows.filter(r=>r.status==="pending");
            else if(requestFilter==="student") filtered=rows.filter(r=>String(r.type||"").includes("student"));
            else if(requestFilter==="subject") filtered=rows.filter(r=>String(r.type||"").includes("subject"));
            else if(requestFilter==="section") filtered=rows.filter(r=>r.type==="add_section");
            else if(requestFilter==="contact") filtered=rows.filter(r=>r.type==="contact_admin");
        }
        const pendingIds=new Set(filtered.filter(r=>r.status==="pending").map(r=>r.id));
        selectedRequestIds=new Set([...selectedRequestIds].filter(id=>pendingIds.has(id)));
        updateBulkRequestToolbar(filtered);
        requestList.innerHTML="";
        if(profile.role!=="admin"){
            if(!filtered.length){requestList.innerHTML=`<p class="empty-record">No requests submitted yet.</p>`;return;}
            filtered.forEach(r=>{
                const card=document.createElement("div");card.className="request-card";
                card.innerHTML=`<div class="request-top"><div class="request-main"><b>${escapeHtml(adminFriendlyRequestType(r.type))} · ${escapeHtml(adminRequestTarget(r))}</b><small>${escapeHtml(r.sectionLabel||adminSectionLabel(r.section||""))}</small><small>${escapeHtml(r.reason||r.message||"")}</small><small>${escapeHtml(r.createdAt?new Date(r.createdAt).toLocaleString():"")}</small></div><span class="${adminRequestStatusClass(r.status||"pending")}">${escapeHtml(String(r.status||"pending").toUpperCase())}</span></div>`;
                requestList.appendChild(card);
            });
            return;
        }
        if(!filtered.length){requestList.innerHTML=`<p class="empty-record">No requests found.</p>`;return;}
        const table=document.createElement("div");table.className="admin-table-card";
        table.innerHTML=`<div class="admin-table-wrap"><table class="admin-data-table request-admin-table"><thead><tr><th></th><th>Type</th><th>Details</th><th>User</th><th>Status</th><th>Date</th><th>Action</th></tr></thead><tbody></tbody></table></div>`;
        const body=table.querySelector("tbody");
        filtered.forEach(r=>{
            const tr=document.createElement("tr");
            const pending=r.status==="pending";
            tr.innerHTML=`<td>${pending?`<input type="checkbox" class="request-checkbox" data-request-id="${escapeHtml(r.id)}" ${selectedRequestIds.has(r.id)?"checked":""}>`:""}</td><td>${escapeHtml(adminFriendlyRequestType(r.type))}</td><td>${escapeHtml(adminRequestTarget(r))}<br><small>${escapeHtml(r.sectionLabel||adminSectionLabel(r.section||""))}</small></td><td>${escapeHtml(r.requestedBy||r.email||"")}</td><td><span class="${adminRequestStatusClass(r.status||"pending")}">${escapeHtml(String(r.status||"pending").toUpperCase())}</span></td><td>${escapeHtml(r.createdAt?new Date(r.createdAt).toLocaleDateString():"")}</td><td><div class="request-table-actions"><button class="request-view-btn" data-view-request="${escapeHtml(r.id)}">👁</button>${pending?`<button class="admin-icon-btn edit" data-approve-request="${escapeHtml(r.id)}">✓</button><button class="admin-icon-btn delete" data-reject-request="${escapeHtml(r.id)}">✕</button>`:""}</div></td>`;
            body.appendChild(tr);
        });
        requestList.appendChild(table);
        body.querySelectorAll(".request-checkbox").forEach(cb=>cb.onchange=()=>{if(cb.checked)selectedRequestIds.add(cb.dataset.requestId);else selectedRequestIds.delete(cb.dataset.requestId);updateBulkRequestToolbar(filtered)});
        body.querySelectorAll("[data-view-request]").forEach(b=>b.onclick=()=>openAdminRequestDetail(window.__qattendRequestCache[b.dataset.viewRequest]));
        body.querySelectorAll("[data-approve-request]").forEach(b=>b.onclick=async()=>{const r=window.__qattendRequestCache[b.dataset.approveRequest];await processChangeRequest(r,true)});
        body.querySelectorAll("[data-reject-request]").forEach(b=>b.onclick=async()=>{const r=window.__qattendRequestCache[b.dataset.rejectRequest];await processChangeRequest(r,false)});
    }catch(e){console.error(e);requestList.innerHTML=`<p class="empty-record">Could not load requests: ${escapeHtml(e.message||e)}</p>`;}
};

/* Dashboard stats: read only the known sections, and include dynamically-created sections. */
loadAdminDashboardData = async function(){
    if(profile?.role!=="admin") return;
    try{
        const sections=await refreshAllAdminSectionSelects();
        const usersSnap=await getDocs(collection(db,"users"));
        const crCount=usersSnap.docs.filter(d=>d.data()?.role==="cr" && d.data()?.active!==false).length;
        let totalStudentsCount=0,totalSubjectsCount=0;
        for(const s of sections){
            const [st,sub]=await Promise.all([getDocs(collection(db,"sections",s.id,"students")),getDocs(collection(db,"sections",s.id,"subjects"))]);
            totalStudentsCount+=st.size; totalSubjectsCount+=sub.size;
        }
        const reqSnap=await getDocs(collection(db,"requests"));
        const pending=reqSnap.docs.filter(d=>d.data()?.status==="pending").length;
        document.getElementById("dashTotalStudents").textContent=totalStudentsCount;
        document.getElementById("dashTotalCRs").textContent=crCount;
        document.getElementById("dashTotalSubjects").textContent=totalSubjectsCount;
        document.getElementById("dashPendingRequests").textContent=pending;
        await Promise.allSettled([loadAdminActivityLogs(),loadAdminRequestPreview()]);
    }catch(e){console.error("Dashboard data failed:",e);}
};

/* Extend hideMainViews/showAdminDashboard so new admin pages never disturb CR screens. */
hideMainViews = function(){
    document.querySelectorAll(".welcome-section,.control-section,#adminPanel,#noSectionMessage,#workArea").forEach(el=>el?.classList.add("hidden"));
    requestsPanel?.classList.add("hidden");adminDashboard?.classList.add("hidden");hideAdminPanels();
};
showAdminDashboard = function(){document.body.classList.remove("records-only-view");hideMainViews();adminDashboard?.classList.remove("hidden");void loadAdminDashboardData();};
showManageData = function(){
    if(profile?.role!=="admin") return;hideMainViews();adminPanel?.classList.remove("hidden");populateSectionDropdowns();if(manageSectionSelect)manageSectionSelect.value=managementSection||"";if(managementSection)void refreshManageData();
};
showRequestsPanel = function(){if(!profile)return;document.body.classList.remove("records-only-view");hideMainViews();requestsPanel?.classList.remove("hidden");void loadRequestCenter();};

/* Wire the additional Admin drawer items. */
document.querySelectorAll('.drawer-item[data-feature="sections"]').forEach(btn=>btn.addEventListener("click",()=>{closeFeatureDrawerFn();showAdminSections();}));
document.querySelectorAll('.drawer-item[data-feature="users"]').forEach(btn=>btn.addEventListener("click",()=>{closeFeatureDrawerFn();showAdminUsers();}));
document.querySelectorAll('.drawer-item[data-feature="reports"]').forEach(btn=>btn.addEventListener("click",()=>{closeFeatureDrawerFn();showAdminReports();}));

/* Reports are generated only when clicked. */
function downloadAdminCsv(filename,headers,rows){
    const csv=[headers,...rows].map(row=>row.map(v=>`"${String(v??"").replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"});
    const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=filename;a.click();URL.revokeObjectURL(url);
}
async function runAdminReport(type){
    if(profile?.role!=="admin") return;
    try{
        const sections=await refreshAllAdminSectionSelects();
        if(type==="requests"){
            const snap=await getDocs(collection(db,"requests"));
            downloadAdminCsv("QAttend_Requests.csv",["Type","Status","User","Section","Details","Reason","Created At","Reviewed At"],snap.docs.map(d=>{const r=d.data();return[adminFriendlyRequestType(r.type),r.status,r.requestedBy||r.email||"",r.sectionLabel||r.section||"",adminRequestTarget(r),r.reason||r.message||"",r.createdAt||"",r.reviewedAt||""]}));
        }else if(type==="students"){
            const rows=[];for(const s of sections){const snap=await getDocs(collection(db,"sections",s.id,"students"));snap.docs.forEach(d=>{const x=d.data();rows.push([s.course||"",s.label||s.id,x.qid||d.id,x.name||""])});}downloadAdminCsv("QAttend_Students.csv",["Course","Section","Q.ID","Student Name"],rows);
        }else if(type==="subjects"){
            const rows=[];for(const s of sections){const snap=await getDocs(collection(db,"sections",s.id,"subjects"));snap.docs.forEach(d=>rows.push([s.course||"",s.label||s.id,d.data()?.name||""]));}downloadAdminCsv("QAttend_Subjects.csv",["Course","Section","Subject"],rows);
        }else if(type==="attendance"){
            const rows=[];for(const s of sections){const snap=await getDocs(collection(db,"sections",s.id,"attendance"));snap.docs.forEach(d=>{const r=d.data();(r.students||[]).forEach(st=>rows.push([r.course||s.course||"",r.section||s.label||s.id,r.subject||"",r.date||"",st.qid||"",st.name||"",st.status||""]));});}downloadAdminCsv("QAttend_Attendance.csv",["Course","Section","Subject","Date","Q.ID","Student Name","Status"],rows);
        }
    }catch(e){alert(`Could not generate report: ${e.message||e}`);}
}
document.querySelectorAll(".admin-report-card[data-report]").forEach(b=>b.addEventListener("click",()=>runAdminReport(b.dataset.report)));

/* On Admin login, refresh the dynamic section catalog once. */
if(typeof auth !== "undefined"){
    onAuthStateChanged(auth, async user=>{
        if(user && profile?.role==="admin") await refreshAllAdminSectionSelects();
    });
}
