import { initializeApp } from
"https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js";

import {
    getAuth,
    signInWithEmailAndPassword,
    signOut,
    onAuthStateChanged
} from
"https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js";

import {
    getFirestore,
    collection,
    doc,
    setDoc,
    getDoc,
    getDocs,
    deleteDoc,
    query,
    orderBy
} from
"https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";


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


// INITIALIZE

const app = initializeApp(firebaseConfig);

const auth = getAuth(app);

const db = getFirestore(app);


// STUDENTS

const students = [

    { qid: "24030101", name: "AAKASH DHIMAN", status: null },
    { qid: "24030571", name: "AARYAN RANA", status: null },
    { qid: "24030545", name: "ADARSH KUMAR", status: null },
    { qid: "24030160", name: "ADDRI GHOSH", status: null },
    { qid: "24030596", name: "ADITYA MITTAL", status: null },
    { qid: "24030307", name: "ADITYA UPRETI", status: null },
    { qid: "25030193", name: "AHZAM KHAN", status: null },
    { qid: "24030640", name: "AKHIL CHANDRA", status: null },
    { qid: "24030649", name: "AKSHAY PUNDIR", status: null },
    { qid: "24030178", name: "ALOK GUPTA", status: null },
    { qid: "24030736", name: "AMAN KUMAR", status: null },
    { qid: "24030400", name: "ANIL KUMAR YADAW", status: null },
    { qid: "24030524", name: "ANIMESH KUMAR DAS", status: null },
    { qid: "24030115", name: "ANINDITA DEY", status: null },
    { qid: "24030617", name: "ANKIT KUMAR", status: null },
    { qid: "24030568", name: "ANKIT SINGH", status: null },
    { qid: "24030035", name: "ANUBHAV KANNAUJIYA", status: null },
    { qid: "24030100", name: "ANUJ KUMAR", status: null },
    { qid: "24030020", name: "ANURAG ROHILA", status: null },
    { qid: "24030137", name: "ANUSHKA", status: null },
    { qid: "24030664", name: "ARYUSH RAJ LAMAN", status: null },
    { qid: "24030514", name: "AVINASH KUMAR", status: null },
    { qid: "24030332", name: "AYUSHI DEY", status: null },
    { qid: "24030089", name: "BAMERRYSHA MARNGAR", status: null },
    { qid: "24030496", name: "BIPIN KUMAR SHARMA", status: null },
    { qid: "24030278", name: "BISHAL KABI", status: null },

    {
        qid: "24030731",
        name: "CHAMANA VEERA SURYA SAI MANIKANTA DASARI BHARATH KUMAR",
        status: null
    },

    { qid: "24030723", name: "DAVANG YADAV", status: null },
    { qid: "24030598", name: "DEEPANSHU CHAUDHARY", status: null },
    { qid: "24030242", name: "DEVANK", status: null },
    { qid: "24030150", name: "DIKSHIT CHAUHAN", status: null },
    { qid: "24030431", name: "DISHA BARUI", status: null },
    { qid: "24030550", name: "GOPAL AGARWAL", status: null },
    { qid: "24030955", name: "HARSH KUMAR GAUTAM", status: null },
    { qid: "24030551", name: "HARSHIT SONI", status: null },
    { qid: "24030459", name: "HEMANT SAINI", status: null },
    { qid: "24030269", name: "JAKKULA LIKHITHA RAO", status: null },
    { qid: "24030699", name: "JANNAT", status: null },
    { qid: "24030734", name: "JATAN KUMAR", status: null },
    { qid: "24030762", name: "KALUGURI MUKKANTI THEERDHARAM", status: null },
    { qid: "24030216", name: "KANISHK TRIPATHI", status: null },
    { qid: "24030110", name: "KARAN KUMAR", status: null },
    { qid: "24030054", name: "KASHISH KUMARI", status: null },
    { qid: "24030583", name: "KAUSHAL KUMAR GUPTA", status: null },
    { qid: "24030585", name: "KRISH KUMAR GUPTA", status: null },
    { qid: "24030341", name: "MUKUND VISHWAKARMA", status: null },
    { qid: "24030009", name: "PRATEEK PATEL", status: null },
    { qid: "24030572", name: "RAKESH KUMAR", status: null },
    { qid: "24030290", name: "RAUNAK RAMESH YADAV", status: null },
    { qid: "24030228", name: "RISHU RAJ", status: null },
    { qid: "24030616", name: "ROHIT KUMAR", status: null },
    { qid: "24030518", name: "SAKSHAM", status: null },
    { qid: "24030410", name: "SATYAM KUMAR", status: null },
    { qid: "24030277", name: "SEJAL SINGH", status: null },
    { qid: "24030713", name: "SHAURYA", status: null },
    { qid: "24030007", name: "SHIVAM", status: null },
    { qid: "24030580", name: "SHIVAM", status: null },
    { qid: "24030652", name: "SHIVAM KUMAR SHARMA", status: null },
    { qid: "24030059", name: "SHOBHIT SINGH", status: null },
    { qid: "24030239", name: "SHREYA DUTTA", status: null },
    { qid: "24030509", name: "SHUBHAM KUMAR", status: null },
    { qid: "24030408", name: "SIDHANT KUMAR", status: null },
    { qid: "24030502", name: "SOHEL ANSARI", status: null },
    { qid: "24030189", name: "SUNDRAM", status: null },
    { qid: "24030320", name: "SWATANTRA SHUKLA", status: null },
    { qid: "24030409", name: "SWATI JAISWAL", status: null },
    { qid: "24030471", name: "TANISHQ GAUTAM", status: null },
    { qid: "24030268", name: "TARANG SAINI", status: null },

    {
        qid: "24030304",
        name: "THALLAPUREDDY CHARAN MANI TEJA REDDY",
        status: null
    },

    {
        qid: "24030303",
        name: "THORAM JISHNU VENKATA SAI VINAY",
        status: null
    },

    { qid: "24030439", name: "UJJWAL KUMAR", status: null },
    { qid: "24030206", name: "VANSH", status: null },
    { qid: "24030241", name: "VANSH JAAT", status: null },
    { qid: "24030095", name: "VANSH KAMBOJ", status: null },
    { qid: "24030131", name: "VANSH TOMAR", status: null },
    { qid: "24030039", name: "VICKY KUMAR", status: null },
    { qid: "24030030", name: "VINIT KUMAR PATEL", status: null },
    { qid: "24030543", name: "VISHAL RAJ", status: null },
    { qid: "24030554", name: "VISHAL SINGH", status: null },
    { qid: "24030672", name: "YASHRAJ KUMAR YADAV", status: null },
    { qid: "24030724", name: "YASHWANT", status: null },
    { qid: "25030688", name: "YUSHA AKHTAR", status: null }

];


// HTML ELEMENTS

const loginPage = document.getElementById("loginPage");
const appPage = document.getElementById("appPage");

const loginForm = document.getElementById("loginForm");

const emailInput = document.getElementById("email");
const passwordInput = document.getElementById("password");
const togglePassword = document.getElementById("togglePassword");

const loginMessage = document.getElementById("loginMessage");

const logoutBtn = document.getElementById("logoutBtn");
const loggedUser = document.getElementById("loggedUser");

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

        await signInWithEmailAndPassword(
            auth,
            email,
            password
        );

        loginMessage.textContent = "";

    } catch (error) {

        console.error(error);

        loginMessage.textContent =
            "Invalid email or password.";

    }

});


// AUTH

onAuthStateChanged(auth, async function (user) {

    if (user) {

        loginPage.classList.add("hidden");
        appPage.classList.remove("hidden");

        loggedUser.textContent = user.email;

        displayStudents();

        await loadRecords();

    } else {

        appPage.classList.add("hidden");
        loginPage.classList.remove("hidden");

    }

});


// LOGOUT

logoutBtn.addEventListener("click", async function () {

    await signOut(auth);

});


// DISPLAY STUDENTS

function displayStudents() {

    const search = searchStudent.value
        .trim()
        .toLowerCase();

    studentList.innerHTML = "";

    students.forEach(function (student, index) {

        const matchesName =
            student.name.toLowerCase().includes(search);

        const matchesQid =
            student.qid.includes(search);

        if (!matchesName && !matchesQid) {
            return;
        }

        const row = document.createElement("div");

        row.className = "student";

        row.innerHTML = `
            <div class="qid-number">
                ${student.qid}
            </div>

            <div class="student-name">
                ${student.name}
            </div>

            <div class="buttons">

                <button
                    class="present ${
                        student.status === "Present"
                            ? "active-present"
                            : ""
                    }"
                    data-index="${index}"
                    data-status="Present"
                >
                    ✓ Present
                </button>

                <button
                    class="absent ${
                        student.status === "Absent"
                            ? "active-absent"
                            : ""
                    }"
                    data-index="${index}"
                    data-status="Absent"
                >
                    ✕ Absent
                </button>

            </div>
        `;

        studentList.appendChild(row);

    });


    document.querySelectorAll("[data-status]")
        .forEach(function (button) {

            button.addEventListener("click", function () {

                const index =
                    Number(button.dataset.index);

                students[index].status =
                    button.dataset.status;

                displayStudents();

            });

        });


    updateStats();

}


// UPDATE STATS

function updateStats() {

    const total = students.length;

    const present = students.filter(
        student => student.status === "Present"
    ).length;

    const absent = students.filter(
        student => student.status === "Absent"
    ).length;

    totalStudents.textContent = total;

    headerTotalStudents.textContent = total;

    presentStudents.textContent = present;

    absentStudents.textContent = absent;

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

        students.forEach(function (student) {
            student.status = "Present";
        });

        displayStudents();

    }
);


// MARK ALL ABSENT

markAllAbsentBtn.addEventListener(
    "click",
    function () {

        students.forEach(function (student) {
            student.status = "Absent";
        });

        displayStudents();

    }
);


// GET ID

function getAttendanceId() {

    const date = attendanceDate.value;

    const subject = subjectSelect.value;

    const cleanSubject =
        subject.replace(/[^a-zA-Z0-9]/g, "_");

    return `${date}_${cleanSubject}`;

}


// LOAD ATTENDANCE

async function loadAttendance() {

    const user = auth.currentUser;

    if (!user) return;

    const date = attendanceDate.value;
    const subject = subjectSelect.value;

    if (!date || !subject) return;


    try {

        const attendanceRef = doc(
            db,
            "users",
            user.uid,
            "attendance",
            getAttendanceId()
        );

        const snapshot =
            await getDoc(attendanceRef);


        if (snapshot.exists()) {

            const data = snapshot.data();

            students.forEach(function (student) {

                const savedStudent =
                    data.students.find(
                        item => item.qid === student.qid
                    );

                student.status =
                    savedStudent
                        ? savedStudent.status
                        : null;

            });

        }

        /*
        IMPORTANT:

        Agar Firebase mein record nahi mila,
        to current marked attendance reset nahi hoga.

        Isliye subject change karne par
        tumhara current marking delete nahi hoga.
        */

        displayStudents();

    } catch (error) {

        console.error(error);

        alert("Error loading attendance.");

    }

}


// SUBJECT CHANGE

subjectSelect.addEventListener(
    "change",
    function () {

        if (!attendanceDate.value) {

            alert("Please select date first.");

            subjectSelect.value = "";

            return;

        }

        loadAttendance();

    }
);


// DATE CHANGE

attendanceDate.addEventListener(
    "change",
    function () {

        if (subjectSelect.value) {
            loadAttendance();
        }

    }
);


// SAVE

saveBtn.addEventListener(
    "click",
    async function () {

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


        const unmarked =
            students.filter(
                student => student.status === null
            );


        if (unmarked.length > 0) {

            alert(
                `${unmarked.length} students are not marked.`
            );

            return;

        }


        try {

            saveBtn.textContent = "Saving...";


            const attendanceRef = doc(
                db,
                "users",
                user.uid,
                "attendance",
                getAttendanceId()
            );


            await setDoc(
                attendanceRef,
                {

                    university:
                        "Quantum University",

                    course:
                        "B.Tech Artificial Intelligence & Machine Learning",

                    section:
                        "AIML - 2",

                    subject: subject,

                    date: date,

                    students: students,

                    updatedAt:
                        new Date().toISOString()

                }
            );


            alert(
                "✅ Attendance saved successfully!"
            );


            await loadRecords();

        } catch (error) {

            console.error(error);

            alert("❌ Error saving attendance.");

        } finally {

            saveBtn.textContent =
                "💾 Save Attendance";

        }

    }
);


// LOAD RECORDS

async function loadRecords() {

    const user = auth.currentUser;

    if (!user) return;


    recordsList.innerHTML =
        `<p class="empty-record">
            Loading records...
        </p>`;


    try {

        const attendanceCollection =
            collection(
                db,
                "users",
                user.uid,
                "attendance"
            );


        const q =
            query(
                attendanceCollection,
                orderBy("updatedAt", "desc")
            );


        const snapshot =
            await getDocs(q);


        recordsList.innerHTML = "";


        if (snapshot.empty) {

            recordsList.innerHTML =
                `<p class="empty-record">
                    No previous attendance records found.
                </p>`;

            return;

        }


        snapshot.forEach(function (documentSnapshot) {

            const data =
                documentSnapshot.data();


            const present =
                data.students.filter(
                    student =>
                        student.status === "Present"
                ).length;


            const absent =
                data.students.filter(
                    student =>
                        student.status === "Absent"
                ).length;


            const card =
                document.createElement("div");


            card.className = "record-card";


            card.innerHTML = `

                <div>

                    <h3>
                        ${data.subject}
                    </h3>

                    <p>
                        📅 ${data.date}
                    </p>

                    <p>
                        🟢 Present: ${present}
                        |
                        🔴 Absent: ${absent}
                    </p>

                </div>


                <div class="record-actions">

                    <button
                        class="view-record-btn"
                        data-id="${documentSnapshot.id}"
                    >
                        View
                    </button>


                    <button
                        class="delete-record-btn"
                        data-id="${documentSnapshot.id}"
                    >
                        Delete
                    </button>

                </div>

            `;


            recordsList.appendChild(card);

        });


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

    document.querySelectorAll(".view-record-btn")
        .forEach(function (button) {

            button.addEventListener(
                "click",
                async function () {

                    await viewRecord(
                        button.dataset.id
                    );

                }
            );

        });


    document.querySelectorAll(".delete-record-btn")
        .forEach(function (button) {

            button.addEventListener(
                "click",
                async function () {

                    await deleteRecord(
                        button.dataset.id
                    );

                }
            );

        });

}


// VIEW RECORD

async function viewRecord(recordId) {

    const user = auth.currentUser;

    if (!user) return;


    try {

        const recordRef = doc(
            db,
            "users",
            user.uid,
            "attendance",
            recordId
        );


        const snapshot =
            await getDoc(recordRef);


        if (!snapshot.exists()) {

            alert("Record not found.");

            return;

        }


        const data = snapshot.data();


        attendanceDate.value = data.date;

        subjectSelect.value = data.subject;


        students.forEach(function (student) {

            const savedStudent =
                data.students.find(
                    item => item.qid === student.qid
                );


            student.status =
                savedStudent
                    ? savedStudent.status
                    : null;

        });


        displayStudents();


        window.scrollTo({
            top: 0,
            behavior: "smooth"
        });

    } catch (error) {

        console.error(error);

        alert("Error loading record.");

    }

}


// DELETE RECORD

async function deleteRecord(recordId) {

    const user = auth.currentUser;

    if (!user) return;


    const confirmDelete = confirm(
        "Are you sure you want to delete this attendance record?"
    );


    if (!confirmDelete) return;


    try {

        await deleteDoc(
            doc(
                db,
                "users",
                user.uid,
                "attendance",
                recordId
            )
        );


        alert("Attendance record deleted.");

        await loadRecords();

    } catch (error) {

        console.error(error);

        alert("Error deleting record.");

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
    function () {

        const date = attendanceDate.value;
        const subject = subjectSelect.value;


        if (!date || !subject) {

            alert("Select date and subject first.");

            return;

        }


        const unmarked =
            students.filter(
                student => student.status === null
            );


        if (unmarked.length > 0) {

            alert(
                "Please mark all students before downloading."
            );

            return;

        }


        const sortedStudents =
            [...students].sort(function (a, b) {

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

            });


        const excelData = [

            ["QATTEND ATTENDANCE REPORT"],

            [],

            ["University", "Quantum University"],

            [
                "Course",
                "B.Tech Artificial Intelligence & Machine Learning"
            ],

            ["Section", "AIML - 2"],

            ["Subject", subject],

            ["Date", date],

            [],

            [
                "Q.ID",
                "Name of Student",
                "Attendance"
            ]

        ];


        sortedStudents.forEach(function (student) {

            excelData.push([
                student.qid,
                student.name,

                student.status === "Present"
                    ? 1
                    : 0
            ]);

        });


        const worksheet =
            XLSX.utils.aoa_to_sheet(
                excelData
            );


        worksheet["!cols"] = [

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
            `QAttend_${cleanSubject}_${date}.xlsx`
        );

    }
);


// INITIAL DISPLAY

displayStudents();