import { initializeApp } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js";
import { getAuth, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js";
import { getFirestore, collection, addDoc, getDocs } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";

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

const statusEl = document.getElementById("status");
const runBtn = document.getElementById("runBtn");

const AIML2_STUDENTS = [
    { qid: "24030101", name: "AAKASH DHIMAN" },
    { qid: "24030571", name: "AARYAN RANA" },
    { qid: "24030545", name: "ADARSH KUMAR" },
    { qid: "24030160", name: "ADDRI GHOSH" },
    { qid: "24030596", name: "ADITYA MITTAL" },
    { qid: "24030307", name: "ADITYA UPRETI" },
    { qid: "25030193", name: "AHZAM KHAN" },
    { qid: "24030640", name: "AKHIL CHANDRA" },
    { qid: "24030649", name: "AKSHAY PUNDIR" },
    { qid: "24030178", name: "ALOK GUPTA" },
    { qid: "24030736", name: "AMAN KUMAR" },
    { qid: "24030400", name: "ANIL KUMAR YADAW" },
    { qid: "24030524", name: "ANIMESH KUMAR DAS" },
    { qid: "24030115", name: "ANINDITA DEY" },
    { qid: "24030617", name: "ANKIT KUMAR" },
    { qid: "24030568", name: "ANKIT SINGH" },
    { qid: "24030035", name: "ANUBHAV KANNAUJIYA" },
    { qid: "24030100", name: "ANUJ KUMAR" },
    { qid: "24030020", name: "ANURAG ROHILA" },
    { qid: "24030137", name: "ANUSHKA" },
    { qid: "24030664", name: "ARYUSH RAJ LAMAN" },
    { qid: "24030514", name: "AVINASH KUMAR" },
    { qid: "24030332", name: "AYUSHI DEY" },
    { qid: "24030089", name: "BAMERRYSHA MARNGAR" },
    { qid: "24030496", name: "BIPIN KUMAR SHARMA" },
    { qid: "24030278", name: "BISHAL KABI" },
    { qid: "24030731", name: "CHAMANA VEERA SURYA SAI MANIKANTA DASARI BHARATH KUMAR" },
    { qid: "24030723", name: "DAVANG YADAV" },
    { qid: "24030598", name: "DEEPANSHU CHAUDHARY" },
    { qid: "24030242", name: "DEVANK" },
    { qid: "24030150", name: "DIKSHIT CHAUHAN" },
    { qid: "24030431", name: "DISHA BARUI" },
    { qid: "24030550", name: "GOPAL AGARWAL" },
    { qid: "24030955", name: "HARSH KUMAR GAUTAM" },
    { qid: "24030551", name: "HARSHIT SONI" },
    { qid: "24030459", name: "HEMANT SAINI" },
    { qid: "24030269", name: "JAKKULA LIKHITHA RAO" },
    { qid: "24030699", name: "JANNAT" },
    { qid: "24030734", name: "JATAN KUMAR" },
    { qid: "24030762", name: "KALUGURI MUKKANTI THEERDHARAM" },
    { qid: "24030216", name: "KANISHK TRIPATHI" },
    { qid: "24030110", name: "KARAN KUMAR" },
    { qid: "24030054", name: "KASHISH KUMARI" },
    { qid: "24030583", name: "KAUSHAL KUMAR GUPTA" },
    { qid: "24030585", name: "KRISH KUMAR GUPTA" },
    { qid: "24030341", name: "MUKUND VISHWAKARMA" },
    { qid: "24030009", name: "PRATEEK PATEL" },
    { qid: "24030572", name: "RAKESH KUMAR" },
    { qid: "24030290", name: "RAUNAK RAMESH YADAV" },
    { qid: "24030228", name: "RISHU RAJ" },
    { qid: "24030616", name: "ROHIT KUMAR" },
    { qid: "24030518", name: "SAKSHAM" },
    { qid: "24030410", name: "SATYAM KUMAR" },
    { qid: "24030277", name: "SEJAL SINGH" },
    { qid: "24030713", name: "SHAURYA" },
    { qid: "24030007", name: "SHIVAM" },
    { qid: "24030580", name: "SHIVAM" },
    { qid: "24030652", name: "SHIVAM KUMAR SHARMA" },
    { qid: "24030059", name: "SHOBHIT SINGH" },
    { qid: "24030239", name: "SHREYA DUTTA" },
    { qid: "24030509", name: "SHUBHAM KUMAR" },
    { qid: "24030408", name: "SIDHANT KUMAR" },
    { qid: "24030502", name: "SOHEL ANSARI" },
    { qid: "24030189", name: "SUNDRAM" },
    { qid: "24030320", name: "SWATANTRA SHUKLA" },
    { qid: "24030409", name: "SWATI JAISWAL" },
    { qid: "24030471", name: "TANISHQ GAUTAM" },
    { qid: "24030268", name: "TARANG SAINI" },
    { qid: "24030304", name: "THALLAPUREDDY CHARAN MANI TEJA REDDY" },
    { qid: "24030303", name: "THORAM JISHNU VENKATA SAI VINAY" },
    { qid: "24030439", name: "UJJWAL KUMAR" },
    { qid: "24030206", name: "VANSH" },
    { qid: "24030241", name: "VANSH JAAT" },
    { qid: "24030095", name: "VANSH KAMBOJ" },
    { qid: "24030131", name: "VANSH TOMAR" },
    { qid: "24030039", name: "VICKY KUMAR" },
    { qid: "24030030", name: "VINIT KUMAR PATEL" },
    { qid: "24030543", name: "VISHAL RAJ" },
    { qid: "24030554", name: "VISHAL SINGH" },
    { qid: "24030672", name: "YASHRAJ KUMAR YADAV" },
    { qid: "24030724", name: "YASHWANT" },
    { qid: "25030688", name: "YUSHA AKHTAR" }
];

const SUBJECTS = [
    "Design and Analysis of Algorithm",
    "Mini Project - III",
    "Design and Analysis of Algorithm Lab",
    "Corporate Communication",
    "Summer Internship II",
    "Technical Skills Development-IV",
    "Scala for Data Science - Tools and Techniques",
    "Foundation of Cloud Computing",
    "Advance Machine Learning Practical with Python, Scikit-learn, TensorFlow",
    "Robotic Industry 4.0"
];

onAuthStateChanged(auth, function (user) {
    if (!user) {
        statusEl.textContent = "Please log in as Admin first (go to index.html, log in, then come back to this page).";
        runBtn.disabled = true;
    }
});

runBtn.addEventListener("click", async function () {
    runBtn.disabled = true;
    statusEl.textContent = "Importing...";

    try {
        // Skip if already imported, to avoid duplicates on re-run
        const existingStudents = await getDocs(collection(db, "sections", "AIML-2", "students"));
        if (existingStudents.empty) {
            for (const student of AIML2_STUDENTS) {
                await addDoc(collection(db, "sections", "AIML-2", "students"), student);
            }
            statusEl.textContent += `\nImported ${AIML2_STUDENTS.length} AIML-2 students.`;
        } else {
            statusEl.textContent += `\nAIML-2 already has students — skipped to avoid duplicates.`;
        }

        const existingSubjects = await getDocs(collection(db, "subjects"));
        if (existingSubjects.empty) {
            for (const name of SUBJECTS) {
                await addDoc(collection(db, "subjects"), { name });
            }
            statusEl.textContent += `\nImported ${SUBJECTS.length} subjects.`;
        } else {
            statusEl.textContent += `\nSubjects already exist — skipped to avoid duplicates.`;
        }

        statusEl.textContent += `\n\nDone. You can delete seed.html and seed.js now.`;

    } catch (error) {
        console.error(error);
        statusEl.textContent += `\n\nError: ${error.message}\n(Are you logged in as Admin?)`;
        runBtn.disabled = false;
    }
});
