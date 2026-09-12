import { initializeApp } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js";
import {
  getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged,
  sendPasswordResetEmail
} from "https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js";
import {
  getFirestore, collection, doc, addDoc, setDoc, getDoc, getDocs, deleteDoc,
  updateDoc, query, orderBy, where, limit, writeBatch
} from "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";

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

const SECTIONS = [
  ["SECTION-1","Section 1"],["SECTION-2","Section 2"],["SECTION-3","Section 3"],["SECTION-4","Section 4"],
  ["SECTION-5","Section 5"],["SECTION-6","Section 6"],["SECTION-7","Section 7"],["SECTION-8","Section 8"],
  ["AIML-1","AIML - 1"],["AIML-2","AIML - 2"],["CSCQ","CSCQ"],["DATA-SCIENCE","Data Science"],
  ["FULL-STACK-DEV","Full Stack Development"],["CLOUD-TECH-INFOSEC","Cloud Technology & Information Security"]
];
const sectionLabelOf = id => (SECTIONS.find(([value]) => value === id)?.[1] || id);
const today = () => new Date().toISOString().slice(0,10);

let profile = null, activeSection = null, students = [], subjects = [], recordsCache = {};
let adminRequests = [], requestFilter = "all", xlsxPromise = null;

const $ = id => document.getElementById(id);
const loginPage = $("loginPage"), appPage = $("appPage");
const loginForm = $("loginForm"), emailInput = $("email"), passwordInput = $("password");
const loginMessage = $("loginMessage"), togglePassword = $("togglePassword");
const forgotPasswordLink = $("forgotPasswordLink"), resetPasswordBox = $("resetPasswordBox");
const resetEmail = $("resetEmail"), sendResetBtn = $("sendResetBtn"), resetMessage = $("resetMessage");
const backToLoginLink = $("backToLoginLink"), logoutBtn = $("logoutBtn");

$("attendanceDate").value = today();

togglePassword.addEventListener("click", () => {
  const show = passwordInput.type === "password";
  passwordInput.type = show ? "text" : "password";
  togglePassword.textContent = show ? "🙈" : "👁";
});

loginForm.addEventListener("submit", async e => {
  e.preventDefault();
  const email = emailInput.value.trim().toLowerCase(), password = passwordInput.value;
  loginMessage.textContent = "Logging in...";
  try {
    await signInWithEmailAndPassword(auth, email, password);
    loginMessage.textContent = "";
  } catch (err) {
    console.error(err);
    loginMessage.textContent = err.code === "auth/too-many-requests"
      ? "Too many attempts. Please try again later."
      : "Invalid email or password.";
  }
});

forgotPasswordLink.addEventListener("click", e => {
  e.preventDefault();
  resetEmail.value = emailInput.value.trim();
  $("loginForm").classList.add("hidden");
  forgotPasswordLink.parentElement.classList.add("hidden");
  resetPasswordBox.classList.remove("hidden");
  loginMessage.textContent = "";
});
backToLoginLink.addEventListener("click", e => {
  e.preventDefault();
  resetPasswordBox.classList.add("hidden");
  $("loginForm").classList.remove("hidden");
  forgotPasswordLink.parentElement.classList.remove("hidden");
  resetMessage.textContent = "";
});
sendResetBtn.addEventListener("click", async () => {
  const email = resetEmail.value.trim().toLowerCase();
  if (!email) return;
  resetMessage.textContent = "Sending...";
  sendResetBtn.disabled = true;
  try {
    await sendPasswordResetEmail(auth, email);
    resetMessage.className = "login-message login-message-success";
    resetMessage.textContent = "Reset link sent. Check your inbox and Spam folder.";
  } catch (err) {
    console.error(err);
    resetMessage.className = "login-message";
    resetMessage.textContent = err.code === "auth/user-not-found"
      ? "No account found for this email."
      : "Could not send reset email. Try again.";
  } finally { sendResetBtn.disabled = false; }
});

function setRoleUI() {
  document.body.classList.toggle("admin-view", profile?.role === "admin");
  document.body.classList.toggle("cr-view", profile?.role === "cr");
  document.querySelectorAll("[data-admin-only]").forEach(el => {
    el.classList.toggle("hidden", profile?.role !== "admin");
  });
  $("sectionSwitchBox").classList.toggle("hidden", profile?.role !== "admin");
  $("newRequestBtn").classList.toggle("hidden", profile?.role !== "cr");
  $("adminRequestTabs").classList.toggle("hidden", profile?.role !== "admin");
  $("requestsTitle").textContent = profile?.role === "admin" ? "Student & Subject Requests" : "My Requests";
  $("requestsSubtitle").textContent = profile?.role === "admin"
    ? "Approve or reject CR requests without giving CRs Admin access."
    : "Send a request to Admin for student or subject changes.";
  $("portalLabel").textContent = profile?.role === "admin" ? "ADMIN ATTENDANCE" : "DIRECT ATTENDANCE";
  $("welcomeTitle").textContent = profile?.role === "admin" ? "Mark Attendance" : "Mark Attendance";
}

function showView(viewId) {
  document.querySelectorAll(".view").forEach(v => v.classList.add("hidden"));
  const target = $(viewId);
  if (target) target.classList.remove("hidden");
  document.querySelectorAll(".nav-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.view === viewId));
  if (viewId === "dashboardView" && profile?.role === "admin") loadAdminDashboard();
  if (viewId === "recordsView") loadRecords();
  if (viewId === "requestsView") loadRequests();
  if (viewId === "adminToolsView" && profile?.role === "admin") {
    populateSectionDropdowns(); loadSubjects(); loadCrList();
  }
}
document.querySelectorAll(".nav-btn, .quick-action, .text-btn, .more-menu-item").forEach(btn => {
  btn.addEventListener("click", () => {
    const view = btn.dataset.view;
    if (!view) return;
    showView(view);
    const target = btn.dataset.scrollTarget;
    if (target) setTimeout(() => $(target)?.scrollIntoView({behavior:"smooth", block:"start"}), 40);
    closeMoreMenu();
  });
});

function closeMoreMenu(){
  $("moreMenu")?.classList.add("hidden");
  $("moreMenuBtn")?.setAttribute("aria-expanded","false");
}
$("moreMenuBtn")?.addEventListener("click", e => {
  e.stopPropagation();
  const menu = $("moreMenu");
  menu.classList.toggle("hidden");
  $("moreMenuBtn").setAttribute("aria-expanded", String(!menu.classList.contains("hidden")));
});
document.addEventListener("click", e => {
  if (!e.target.closest(".more-menu-wrap")) closeMoreMenu();
});

function populateSectionDropdowns() {
  $("sectionSelect").innerHTML = `<option value="">-- Select Section --</option>`;
  $("newCrSection").innerHTML = "";
  SECTIONS.forEach(([id,label]) => {
    $("sectionSelect").insertAdjacentHTML("beforeend", `<option value="${id}">${label}</option>`);
    $("newCrSection").insertAdjacentHTML("beforeend", `<option value="${id}">${label}</option>`);
  });
}
$("sectionSelect").addEventListener("change", async () => {
  activeSection = $("sectionSelect").value || null;
  $("attendanceWorkArea").classList.toggle("hidden", !activeSection);
  $("noSectionMessage").classList.toggle("hidden", !!activeSection);
  $("sectionLabel").textContent = activeSection ? sectionLabelOf(activeSection) : "-- Not selected --";
  $("subjectSelect").value = "";
  $("attendanceDate").value = today();
  if (!activeSection) return resetAttendanceUI();
  await loadSectionData();
});

async function loadSectionData() {
  $("headerTotalStudents").textContent = "Loading...";
  await Promise.all([loadStudents(), loadSubjects()]);
  $("headerTotalStudents").textContent = students.length;
  // Never auto-select a subject. The user chooses the subject explicitly.
  $("subjectSelect").value = "";
}

async function loadStudents() {
  if (!activeSection) return;
  try {
    const snap = await getDocs(collection(db,"sections",activeSection,"students"));
    students = snap.docs.map(d => ({id:d.id, ...d.data(), status:null}))
      .sort((a,b) => String(a.name).localeCompare(String(b.name)));
    displayStudents();
    $("headerTotalStudents").textContent = students.length;
  } catch (e) {
    console.error(e);
    $("studentList").innerHTML = `<p class="empty-record">Error loading students.</p>`;
  }
}

function displayStudents() {
  const search = $("searchStudent").value.trim().toLowerCase();
  const isAdmin = profile?.role === "admin";
  $("studentList").innerHTML = "";
  for (const student of students) {
    if (search && !String(student.name).toLowerCase().includes(search) && !String(student.qid).toLowerCase().includes(search)) continue;
    const row = document.createElement("div");
    row.className = "student";
    row.dataset.id = student.id;
    row.innerHTML = `
      <div class="qid-number">${escapeHtml(student.qid)}</div>
      <div class="student-name">${escapeHtml(student.name)}</div>
      <div class="buttons">
        <button type="button" class="absent ${student.status === "Absent" ? "active-absent":""}" data-id="${student.id}" data-status="Absent">✕ Absent</button>
        <button type="button" class="present ${student.status === "Present" ? "active-present":""}" data-id="${student.id}" data-status="Present">✓ Present</button>
      </div>
      ${isAdmin ? `<div class="student-menu-wrapper">
        <button type="button" class="student-menu-btn" data-menu-id="${student.id}" aria-label="Student options">⋮</button>
        <div class="student-menu" data-menu="${student.id}">
          <button type="button" class="edit-student-btn" data-id="${student.id}">✏️ Edit</button>
          <button type="button" class="delete-student-menu-btn" data-id="${student.id}">🗑️ Delete</button>
        </div>
      </div>` : ""}
    `;
    $("studentList").appendChild(row);
  }
  updateStats();
}
$("searchStudent").addEventListener("input", displayStudents);

$("studentList").addEventListener("click", async e => {
  const menuBtn = e.target.closest(".student-menu-btn");
  if (menuBtn) {
    document.querySelectorAll(".student-menu.show").forEach(m => m.classList.remove("show"));
    document.querySelector(`.student-menu[data-menu="${menuBtn.dataset.menuId}"]`)?.classList.add("show");
    return;
  }
  const editBtn = e.target.closest(".edit-student-btn");
  if (editBtn) { await editStudent(editBtn.dataset.id); return; }
  const delBtn = e.target.closest(".delete-student-menu-btn");
  if (delBtn) { await deleteStudent(delBtn.dataset.id); return; }

  const statusBtn = e.target.closest("[data-status]");
  if (!statusBtn) return;
  if (!$("subjectSelect").value) { alert("Please select a subject first."); return; }
  const student = students.find(s => s.id === statusBtn.dataset.id);
  if (student) student.status = statusBtn.dataset.status;
  displayStudents();
});
document.addEventListener("click", e => {
  if (!e.target.closest(".student-menu-wrapper")) document.querySelectorAll(".student-menu.show").forEach(m => m.classList.remove("show"));
});

function updateStats() {
  const total = students.length, present = students.filter(s => s.status === "Present").length, absent = students.filter(s => s.status === "Absent").length;
  $("totalStudents").textContent = total; $("presentStudents").textContent = present; $("absentStudents").textContent = absent;
}

$("markAllPresentBtn").addEventListener("click", () => { students.forEach(s=>s.status="Present"); displayStudents(); });
$("markAllAbsentBtn").addEventListener("click", () => { students.forEach(s=>s.status="Absent"); displayStudents(); });

async function editStudent(studentId) {
  if (profile?.role !== "admin") return;
  const student = students.find(s => s.id === studentId); if (!student) return;
  const name = prompt("Enter student name:", student.name); if (name === null) return;
  const qid = prompt("Enter Q.ID:", student.qid); if (qid === null) return;
  if (!name.trim() || !qid.trim()) return alert("Name and Q.ID are required.");
  try {
    await updateDoc(doc(db,"sections",activeSection,"students",studentId), {name:name.trim().toUpperCase(),qid:qid.trim()});
    await loadStudents();
  } catch (e) { console.error(e); alert("Error updating student."); }
}
async function deleteStudent(studentId) {
  if (profile?.role !== "admin") return;
  const student = students.find(s=>s.id===studentId);
  if (!student || !confirm(`Delete ${student.name} (${student.qid})?`)) return;
  try {
    await deleteDoc(doc(db,"sections",activeSection,"students",studentId));
    await loadStudents();
  } catch (e) { console.error(e); alert("Error deleting student."); }
}

async function loadSubjects() {
  if (!activeSection) return;
  try {
    const snap = await getDocs(collection(db,"sections",activeSection,"subjects"));
    subjects = snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>String(a.name).localeCompare(String(b.name)));
    $("subjectSelect").innerHTML = `<option value="">-- Select Subject --</option>`;
    subjects.forEach(s => $("subjectSelect").insertAdjacentHTML("beforeend", `<option value="${escapeHtml(s.name)}">${escapeHtml(s.name)}</option>`));
    // Keep the subject selector blank until the user selects one.
    $("subjectSelect").value = "";
    renderSubjectManageList();
  } catch (e) { console.error(e); }
}
function renderSubjectManageList() {
  $("subjectManageList").innerHTML = "";
  subjects.forEach(s => {
    const row = document.createElement("div"); row.className="manage-row";
    row.innerHTML = `<span>${escapeHtml(s.name)}</span><button class="small-danger" data-subject-delete="${s.id}">Delete</button>`;
    $("subjectManageList").appendChild(row);
  });
  $("subjectManageList").querySelectorAll("[data-subject-delete]").forEach(b => b.addEventListener("click", ()=>deleteSubject(b.dataset.subjectDelete)));
}
async function deleteSubject(subjectId) {
  if (profile?.role !== "admin" || !activeSection) return;
  const s = subjects.find(x=>x.id===subjectId);
  if (!s || !confirm(`Delete subject "${s.name}" from ${sectionLabelOf(activeSection)}?`)) return;
  try { await deleteDoc(doc(db,"sections",activeSection,"subjects",subjectId)); await loadSubjects(); }
  catch(e){ console.error(e); alert("Error deleting subject.");}
}
$("addStudentForm").addEventListener("submit", async e => {
  e.preventDefault();
  if(profile?.role!=="admin" || !activeSection) return alert("Select a section in Mark Attendance first.");
  const qid=$("newStudentQid").value.trim(), name=$("newStudentName").value.trim().toUpperCase();
  if(!qid||!name) return;
  try { await addDoc(collection(db,"sections",activeSection,"students"),{qid,name}); $("newStudentQid").value=""; $("newStudentName").value=""; await loadSectionData(); alert("Student added."); }
  catch(e){ console.error(e); alert("Error adding student.");}
});
$("addSubjectForm").addEventListener("submit", async e => {
  e.preventDefault();
  if(profile?.role!=="admin" || !activeSection) return alert("Select a section in Mark Attendance first.");
  const name=$("newSubjectName").value.trim(); if(!name) return;
  try { await addDoc(collection(db,"sections",activeSection,"subjects"),{name}); $("newSubjectName").value=""; await loadSubjects(); alert("Subject added."); }
  catch(e){ console.error(e); alert("Error adding subject.");}
});

async function loadCrList() {
  if(profile?.role!=="admin") return;
  try {
    const snap=await getDocs(collection(db,"users"));
    $("crManageList").innerHTML="";
    snap.docs.filter(d=>d.data().role==="cr").forEach(d=>{
      const data=d.data(), row=document.createElement("div"); row.className="manage-row";
      row.innerHTML=`<span>${escapeHtml(d.id)}<br><small>${escapeHtml(sectionLabelOf(data.section))}</small></span><button class="small-danger" data-remove-cr="${escapeHtml(d.id)}">Remove</button>`;
      $("crManageList").appendChild(row);
    });
    $("crManageList").querySelectorAll("[data-remove-cr]").forEach(b=>b.addEventListener("click",()=>removeCr(b.dataset.removeCr)));
  } catch(e){console.error(e);}
}
$("addCrForm").addEventListener("submit", async e=>{
  e.preventDefault();
  if(profile?.role!=="admin") return;
  const email=$("newCrEmail").value.trim().toLowerCase(), section=$("newCrSection").value; if(!email||!section)return;
  try{
    await setDoc(doc(db,"users",email),{role:"cr",section,createdAt:new Date().toISOString()});
    $("newCrEmail").value=""; await loadCrList(); await loadAdminDashboard(); alert(`${email} is assigned as CR for ${sectionLabelOf(section)}.`);
  }catch(err){console.error(err);alert("Error assigning CR.");}
});
async function removeCr(email){
  if(profile?.role!=="admin" || !confirm(`Remove CR access for ${email}?`)) return;
  try{await deleteDoc(doc(db,"users",email));await loadCrList();await loadAdminDashboard();}catch(e){console.error(e);alert("Error removing CR.");}
}

function getAttendanceId() {
  const date=$("attendanceDate").value, subject=$("subjectSelect").value.replace(/[^a-zA-Z0-9]/g,"_");
  return `${date}_${subject}`;
}
async function loadAttendance() {
  if(!activeSection || !$("attendanceDate").value || !$("subjectSelect").value) return;
  try{
    const snap=await getDoc(doc(db,"sections",activeSection,"attendance",getAttendanceId()));
    const data=snap.exists()?snap.data():null;
    students.forEach(s=>s.status=null);
    if(data?.students) students.forEach(s=>{const found=data.students.find(x=>x.qid===s.qid);if(found)s.status=found.status;});
    displayStudents();
  }catch(e){console.error(e);alert("Error loading attendance.");}
}
$("subjectSelect").addEventListener("change",loadAttendance);
$("attendanceDate").addEventListener("change",loadAttendance);

$("saveBtn").addEventListener("click",async()=>{
  const user=auth.currentUser;if(!user||!activeSection)return;
  const date=$("attendanceDate").value, subject=$("subjectSelect").value;
  if(!date)return alert("Please select date.");
  if(!subject)return alert("Please select subject.");
  const unmarked=students.filter(s=>!s.status); if(unmarked.length)return alert(`${unmarked.length} students are not marked.`);
  $("saveBtn").textContent="Saving...";$("saveBtn").disabled=true;
  try{
    await setDoc(doc(db,"sections",activeSection,"attendance",getAttendanceId()),{
      university:"Quantum University",course:"B.Tech",section:sectionLabelOf(activeSection),subject,date,
      students:students.map(({id,...rest})=>rest),markedBy:user.email,updatedAt:new Date().toISOString()
    });
    alert("✅ Attendance saved successfully!"); await loadRecords(); await loadAdminDashboard(true);
  }catch(e){console.error(e);alert("❌ Error saving attendance.");}
  finally{$("saveBtn").textContent="💾 Save Attendance";$("saveBtn").disabled=false;}
});

async function loadRecords() {
  if(!activeSection){$("recordsList").innerHTML=`<p class="empty-record">Select a section to view records.</p>`;return;}
  $("recordsList").innerHTML=`<p class="empty-record">Loading records...</p>`;
  try{
    const snap=await getDocs(collection(db,"sections",activeSection,"attendance"));
    recordsCache={};
    const rows=snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
    if(!rows.length){$("recordsList").innerHTML=`<p class="empty-record">No saved attendance records.</p>`;return;}
    rows.forEach(r=>recordsCache[r.id]=r);
    $("recordsList").innerHTML="";
    rows.forEach(r=>{
      const card=document.createElement("div");card.className="record-card";
      card.innerHTML=`<div class="record-info"><b>${escapeHtml(r.subject||"Subject")}</b><span>${escapeHtml(r.date||"")} · ${escapeHtml(r.section||sectionLabelOf(activeSection))} · by ${escapeHtml(r.markedBy||"")}</span></div>
      <div class="record-actions">
        <button class="view-record-btn" data-record="${r.id}">View</button>
        <button class="download-record-btn" data-record="${r.id}">Download</button>
        <button class="share-record-btn" data-record="${r.id}">Share</button>
        <button class="delete-record-btn" data-record="${r.id}">Delete</button>
      </div>`;
      $("recordsList").appendChild(card);
    });
  }catch(e){console.error(e);$("recordsList").innerHTML=`<p class="empty-record">Error loading records.</p>`;}
}
$("refreshRecordsBtn").addEventListener("click",loadRecords);
$("recordsList").addEventListener("click",async e=>{
  const id=e.target.dataset.record;if(!id)return;
  if(e.target.classList.contains("view-record-btn")) await viewRecord(id);
  if(e.target.classList.contains("delete-record-btn")) await deleteRecord(id);
  if(e.target.classList.contains("download-record-btn")) await exportRecord(id);
  if(e.target.classList.contains("share-record-btn")) await shareRecord(id);
});

async function viewRecord(id){
  const r=recordsCache[id];if(!r)return;
  $("attendanceDate").value=r.date||today();$("subjectSelect").value=r.subject||"";
  await loadAttendance();showView("attendanceView");window.scrollTo({top:0,behavior:"smooth"});
}
async function deleteRecord(id){
  if(!confirm("Are you sure you want to delete this attendance record?"))return;
  try{await deleteDoc(doc(db,"sections",activeSection,"attendance",id));await loadRecords();}catch(e){console.error(e);alert("Error deleting record.");}
}

async function ensureXLSXLoaded(){
  if(window.XLSX)return;
  if(!xlsxPromise){
    xlsxPromise=new Promise((resolve,reject)=>{
      const s=document.createElement("script");s.src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";
      s.onload=resolve;s.onerror=()=>reject(new Error("Could not load Excel library."));document.head.appendChild(s);
    });
  }
  return xlsxPromise;
}
function buildWorkbook(record){
  const rows=[["QATTEND ATTENDANCE REPORT"],[],["University",record.university||"Quantum University"],["Course",record.course||"B.Tech"],["Section",record.section],["Semester","5"],["Subject",record.subject],["Date",record.date],[],["S.No","Q.ID","Student Name","Attendance"]];
  const sorted=[...record.students].sort((a,b)=>(a.status==="Present"?0:1)-(b.status==="Present"?0:1));
  sorted.forEach((s,i)=>rows.push([i+1,s.qid,s.name,s.status==="Present"?1:0]));
  return rows;
}
async function exportRecord(id){
  try{
    const record=recordsCache[id]||await getDoc(doc(db,"sections",activeSection,"attendance",id)).then(s=>s.data());
    await ensureXLSXLoaded();
    const ws=XLSX.utils.aoa_to_sheet(buildWorkbook(record));ws["!cols"]=[{wch:8},{wch:18},{wch:36},{wch:12}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Attendance");
    const name=`QAttend_${record.section}_${String(record.subject).replace(/[^a-zA-Z0-9]/g,"_")}_${record.date}.xlsx`;
    XLSX.writeFile(wb,name);
  }catch(e){console.error(e);alert("Error creating Excel file.");}
}
async function shareRecord(id){
  try{
    const record=recordsCache[id]||await getDoc(doc(db,"sections",activeSection,"attendance",id)).then(s=>s.data());
    await ensureXLSXLoaded();
    const ws=XLSX.utils.aoa_to_sheet(buildWorkbook(record));ws["!cols"]=[{wch:8},{wch:18},{wch:36},{wch:12}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Attendance");
    const fileName=`QAttend_${record.section}_${String(record.subject).replace(/[^a-zA-Z0-9]/g,"_")}_${record.date}.xlsx`;
    const blob=new Blob([XLSX.write(wb,{bookType:"xlsx",type:"array"})],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const file=new File([blob],fileName,{type:blob.type});
    if(navigator.share&&navigator.canShare?.({files:[file]})){await navigator.share({files:[file],title:`${record.subject} Attendance`,text:`${record.subject} · ${record.date} · ${record.section}`});}
    else {XLSX.writeFile(wb,fileName);alert("Sharing is not supported here, so the file was downloaded instead.");}
  }catch(e){console.error(e);alert("Error sharing record.");}
}
$("downloadBtn").addEventListener("click",async()=>{
  const subject=$("subjectSelect").value;if(!subject)return alert("Please select subject first.");
  const id=getAttendanceId();const snap=await getDoc(doc(db,"sections",activeSection,"attendance",id));
  if(!snap.exists())return alert("Save attendance first or open a saved record.");
  recordsCache[id]=snap.data();await exportRecord(id);
});
$("shareCurrentBtn").addEventListener("click",async()=>{
  const subject=$("subjectSelect").value;if(!subject)return alert("Please select subject first.");
  const id=getAttendanceId();const snap=await getDoc(doc(db,"sections",activeSection,"attendance",id));
  if(!snap.exists())return alert("Save attendance first or open a saved record.");
  recordsCache[id]=snap.data();await shareRecord(id);
});

async function submitChangeRequest(data){
  await addDoc(collection(db,"requests"),{
    ...data,
    requestedBy:auth.currentUser.email,
    requestedByName:auth.currentUser.email.split("@")[0],
    section:activeSection,
    sectionLabel:sectionLabelOf(activeSection),
    status:"pending",
    createdAt:new Date().toISOString()
  });
}
function openRequestModal(){
  $("requestModal").classList.remove("hidden"); syncRequestFields();
}
function closeRequestModal(){ $("requestModal").classList.add("hidden"); $("requestForm").reset(); syncRequestFields(); }
$("newRequestBtn").addEventListener("click",openRequestModal);
$("closeRequestModal").addEventListener("click",closeRequestModal);
$("cancelRequestBtn").addEventListener("click",closeRequestModal);
$("requestType").addEventListener("change",syncRequestFields);
function syncRequestFields(){
  const type=$("requestType").value, subject=type.includes("subject");
  $("requestStudentFields").classList.toggle("hidden",subject);
  $("requestSubjectFields").classList.toggle("hidden",!subject);
  if(type==="add_student"){ $("requestStudentName").parentElement.style.display="flex"; $("requestQid").placeholder="2403...."; }
  else if(type==="delete_student"){ $("requestStudentName").parentElement.style.display="none"; }
  else if(type==="edit_student"){ $("requestStudentName").parentElement.style.display="flex"; }
  else $("requestStudentName").parentElement.style.display="none";
}
$("requestForm").addEventListener("submit",async e=>{
  e.preventDefault();
  if(profile?.role!=="cr"||!activeSection)return;
  const type=$("requestType").value, reason=$("requestReason").value.trim();
  if(!reason)return alert("Please enter a reason.");
  const data={type,reason};
  if(type.includes("student")){
    data.qid=$("requestQid").value.trim();
    data.name=$("requestStudentName").value.trim().toUpperCase();
    if(!data.qid || ((type!=="delete_student")&&!data.name)) return alert("Enter the required student details.");
  }else{
    data.subject=$("requestSubjectName").value.trim();
    if(!data.subject)return alert("Enter subject name.");
  }
  try{await submitChangeRequest(data);closeRequestModal();alert("✅ Request sent to Admin.");await loadRequests();}
  catch(err){console.error(err);alert("Could not send request. Check Firestore Rules and try again.");}
});

async function loadRequests(){
  if(!profile)return;
  if(profile.role==="admin") await loadAdminRequests();
  else $("requestList").innerHTML=`<div class="panel-card"><p class="empty-record">Requests are sent directly to Admin. You do not need Admin access.</p></div>`;
}
async function loadAdminRequests(){
  try{
    const snap=await getDocs(collection(db,"requests"));
    adminRequests=snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>String(b.createdAt||"").localeCompare(String(a.createdAt||"")));
    renderRequests();
    const pending=adminRequests.filter(r=>r.status==="pending").length;
    $("pendingBadge").textContent=pending;$("pendingBadge").classList.toggle("hidden",pending===0);
    $("dashPendingRequests").textContent=pending;
    $("dashboardPendingList").innerHTML="";
    adminRequests.filter(r=>r.status==="pending").slice(0,4).forEach(r=>{
      const el=document.createElement("div");el.className="activity-item";
      el.innerHTML=`<span class="activity-icon">${r.type.includes("student")?"👤":"📚"}</span><div><b>${friendlyRequestType(r.type)}</b><small>${escapeHtml(r.sectionLabel||r.section)} · ${escapeHtml(r.requestedBy||"")}</small></div>`;
      $("dashboardPendingList").appendChild(el);
    });
    if(!$("dashboardPendingList").children.length)$("dashboardPendingList").innerHTML=`<p class="empty-record">No pending requests.</p>`;
  }catch(e){console.error(e);$("requestList").innerHTML=`<p class="empty-record">Could not load requests.</p>`;}
}
function friendlyRequestType(type){return ({add_student:"Add Student",edit_student:"Edit Student",delete_student:"Delete Student",add_subject:"Add Subject",delete_subject:"Delete Subject"})[type]||type;}
function renderRequests(){
  const list=adminRequests.filter(r=>{
    if(requestFilter==="pending")return r.status==="pending";
    if(requestFilter==="student")return r.type.includes("student");
    if(requestFilter==="subject")return r.type.includes("subject");
    return true;
  });
  $("requestList").innerHTML="";
  if(!list.length){$("requestList").innerHTML=`<p class="empty-record">No requests found.</p>`;return;}
  list.forEach(r=>{
    const card=document.createElement("div");card.className="request-card";
    const title=r.type.includes("student") ? `${friendlyRequestType(r.type)}${r.qid?` · ${escapeHtml(r.qid)}`:""}` : `${friendlyRequestType(r.type)}${r.subject?` · ${escapeHtml(r.subject)}`:""}`;
    card.innerHTML=`<div class="request-top"><div class="request-main"><b>${title}</b><small>CR: ${escapeHtml(r.requestedBy||"")} · ${escapeHtml(r.sectionLabel||r.section)}</small><small>${escapeHtml(r.reason||"")}</small></div><span class="status-pill status-${r.status||"pending"}">${escapeHtml((r.status||"pending").toUpperCase())}</span></div>`;
    if(r.status==="pending"&&profile?.role==="admin"){
      const actions=document.createElement("div");actions.className="request-actions";
      const approve=document.createElement("button");approve.className="approve-btn";approve.textContent="Approve";
      const reject=document.createElement("button");reject.className="reject-btn";reject.textContent="Reject";
      approve.onclick=()=>processRequest(r,true);reject.onclick=()=>processRequest(r,false);
      actions.append(approve,reject);card.appendChild(actions);
    }
    $("requestList").appendChild(card);
  });
}
document.querySelectorAll(".request-tab").forEach(btn=>btn.addEventListener("click",()=>{
  document.querySelectorAll(".request-tab").forEach(b=>b.classList.remove("active"));btn.classList.add("active");
  requestFilter=btn.dataset.requestFilter;renderRequests();
}));

async function processRequest(r,approve){
  if(profile?.role!=="admin")return;
  try{
    if(approve) await applyApprovedRequest(r);
    await updateDoc(doc(db,"requests",r.id),{status:approve?"approved":"rejected",reviewedBy:auth.currentUser.email,reviewedAt:new Date().toISOString()});
    await loadAdminRequests(); await loadAdminDashboard();
    alert(approve?"✅ Request approved and change applied.":"Request rejected.");
  }catch(e){console.error(e);alert(`Could not ${approve?"approve":"reject"} request.`);}
}
async function findStudentDoc(section,qid){
  const snap=await getDocs(collection(db,"sections",section,"students"));
  return snap.docs.find(d=>String(d.data().qid)===String(qid));
}
async function findSubjectDoc(section,name){
  const snap=await getDocs(collection(db,"sections",section,"subjects"));
  return snap.docs.find(d=>String(d.data().name).toLowerCase()===String(name).toLowerCase());
}
async function applyApprovedRequest(r){
  const section=r.section;
  if(r.type==="add_student") await addDoc(collection(db,"sections",section,"students"),{qid:r.qid,name:String(r.name).toUpperCase()});
  if(r.type==="edit_student"){
    const d=await findStudentDoc(section,r.qid);if(!d)throw new Error("Student not found");
    await updateDoc(d.ref,{name:String(r.name).toUpperCase(),qid:r.qid});
  }
  if(r.type==="delete_student"){
    const d=await findStudentDoc(section,r.qid);if(!d)throw new Error("Student not found");
    await deleteDoc(d.ref);
  }
  if(r.type==="add_subject") await addDoc(collection(db,"sections",section,"subjects"),{name:r.subject});
  if(r.type==="delete_subject"){
    const d=await findSubjectDoc(section,r.subject);if(!d)throw new Error("Subject not found");
    await deleteDoc(d.ref);
  }
}

async function loadAdminDashboard(force=false){
  if(profile?.role!=="admin")return;
  if(!force && $("dashboardView").classList.contains("hidden") && !document.body.classList.contains("admin-view"))return;
  try{
    const usersSnap=await getDocs(collection(db,"users"));
    const crs=usersSnap.docs.filter(d=>d.data().role==="cr");
    let totalStudents=0,totalSubjects=0;
    const tasks=SECTIONS.map(async([id])=>{
      const [ss,sub]=await Promise.all([getDocs(collection(db,"sections",id,"students")),getDocs(collection(db,"sections",id,"subjects"))]);
      totalStudents+=ss.size;totalSubjects+=sub.size;
    });
    await Promise.all(tasks);
    $("dashTotalStudents").textContent=totalStudents;
    $("dashTotalSubjects").textContent=totalSubjects;
    $("dashTotalCRs").textContent=crs.length;
    const reqSnap=await getDocs(collection(db,"requests"));
    adminRequests=reqSnap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>String(b.createdAt||"").localeCompare(String(a.createdAt||"")));
    const pending=adminRequests.filter(r=>r.status==="pending").length;
    $("dashPendingRequests").textContent=pending;$("pendingBadge").textContent=pending;$("pendingBadge").classList.toggle("hidden",pending===0);
    $("recentActivity").innerHTML="";
    const recent=adminRequests.slice(0,5);
    recent.forEach(r=>{
      const el=document.createElement("div");el.className="activity-item";
      el.innerHTML=`<span class="activity-icon">${r.type.includes("student")?"👤":"📚"}</span><div><b>${friendlyRequestType(r.type)}</b><small>${escapeHtml(r.requestedBy||"")} · ${escapeHtml(r.sectionLabel||r.section)} · ${escapeHtml(r.status||"pending")}</small></div>`;
      $("recentActivity").appendChild(el);
    });
    if(!recent.length)$("recentActivity").innerHTML=`<p class="empty-record">No recent activity yet.</p>`;
    $("dashboardPendingList").innerHTML="";
    adminRequests.filter(r=>r.status==="pending").slice(0,4).forEach(r=>{
      const el=document.createElement("div");el.className="activity-item";
      el.innerHTML=`<span class="activity-icon">${r.type.includes("student")?"👤":"📚"}</span><div><b>${friendlyRequestType(r.type)}</b><small>${escapeHtml(r.requestedBy||"")} · ${escapeHtml(r.sectionLabel||r.section)}</small></div>`;
      $("dashboardPendingList").appendChild(el);
    });
    if(!pending)$("dashboardPendingList").innerHTML=`<p class="empty-record">No pending requests.</p>`;
  }catch(e){console.error(e);}
}

async function resetAttendanceUI(){
  $("studentList").innerHTML="";students=[];subjects=[];$("subjectSelect").innerHTML=`<option value="">-- Select Subject --</option>`;
  updateStats();
}
function escapeHtml(value=""){
  return String(value).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]));
}

onAuthStateChanged(auth, async user=>{
  if(!user){
    appPage.classList.add("hidden");loginPage.classList.remove("hidden");profile=null;activeSection=null;
    return;
  }
  try{
    const snap=await getDoc(doc(db,"users",user.email.toLowerCase()));
    if(!snap.exists()){alert("Your account is not set up yet. Ask the Admin to assign you a role.");await signOut(auth);return;}
    profile=snap.data();
  }catch(e){console.error(e);alert("Could not load your account role.");await signOut(auth);return;}

  loginPage.classList.add("hidden");appPage.classList.remove("hidden");
  $("loggedUser").textContent=user.email;
  $("roleBadge").textContent=profile.role==="admin"?"ADMIN":"CLASS REPRESENTATIVE";
  setRoleUI();

  populateSectionDropdowns();
  $("attendanceDate").value=today();
  if(profile.role==="cr"){
    activeSection=profile.section;
    $("sectionLabel").textContent=sectionLabelOf(activeSection);
    $("attendanceWorkArea").classList.remove("hidden");
    $("noSectionMessage").classList.add("hidden");
    $("welcomeSubtitle").textContent=`${sectionLabelOf(activeSection)} is automatically selected for you.`;
    showView("attendanceView");
    await loadSectionData();     // CR only loads its assigned section
  }else{
    activeSection=null;
    $("attendanceWorkArea").classList.add("hidden");
    $("noSectionMessage").classList.remove("hidden");
    $("sectionSelect").value="";
    showView("dashboardView");   // Admin lands on dashboard
    await Promise.all([loadAdminDashboard(),loadCrList()]);
  }
});

logoutBtn.addEventListener("click",()=>signOut(auth));
