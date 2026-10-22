let participantData = [];
let clearTimer = null;
let lastId = null;

function toggleMenu() {
    const navList = document.getElementById('nav-list');
    navList.classList.toggle('active');
}

function saveToLocal() {
    localStorage.setItem('acl_data_2026', JSON.stringify(participantData));
}

function loadFromLocal() {
    const savedData = localStorage.getItem('acl_data_2026');
    if (savedData) {
        participantData = JSON.parse(savedData);
        document.getElementById('download-btn').style.display = 'block';
        updateRealtimeTable();
    }
}

function systemReset() {
    const confirmReset = confirm("คำเตือน: ข้อมูลรายชื่อและการเช็คชื่อรวมถึงโน้ตทั้งหมดจะถูกลบถาวร ต้องการดำเนินการต่อหรือไม่?");
    if (confirmReset) {
        localStorage.removeItem('acl_data_2026');
        localStorage.removeItem('acl_note_2026');
        location.reload();
    }
}

function showPage(pageId) {
    const pages = ['home-page', 'realtime-page', 'note-page'];
    const navs = ['nav-home', 'nav-realtime', 'nav-note'];
    pages.forEach(p => document.getElementById(p).style.display = 'none');
    navs.forEach(n => document.getElementById(n).classList.remove('active'));
    document.getElementById(pageId + '-page').style.display = 'block';
    document.getElementById('nav-' + pageId).classList.add('active');
    document.getElementById('nav-list').classList.remove('active');
    if (pageId === 'realtime') { updateRealtimeTable(); }
    else if (pageId === 'note') { loadNote(); }
}

function saveNote() {
    const content = document.getElementById('note-input').value;
    localStorage.setItem('acl_note_2026', content);
    document.getElementById('note-status').innerText = 'กำลังบันทึก...';
    setTimeout(() => { document.getElementById('note-status').innerText = 'บันทึกข้อมูลแล้ว'; }, 500);
}

function loadNote() {
    const savedContent = localStorage.getItem('acl_note_2026');
    if (savedContent !== null) { document.getElementById('note-input').value = savedContent; }
}

document.getElementById('excel-upload').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        participantData = XLSX.utils.sheet_to_json(sheet);
        saveToLocal();
        alert('โหลดข้อมูลเรียบร้อยแล้วจำนวน ' + participantData.length + ' รายชื่อ');
        document.getElementById('download-btn').style.display = 'block';
        updateRealtimeTable();
        document.getElementById('id-input').focus();
    };
    reader.readAsBinaryString(file);
});

let scanTimeout = null;
let lastKeyTime = 0;

function handleScanner(event) {
    if (event.key === "Enter") {
        const id = event.target.value.trim();
        if (id) {
            clearTimeout(scanTimeout);
            findUser(id);
            event.target.value = "";
        }
        return;
    }

    // ตรวจจับว่าเป็น QR Scanner (พิมพ์เร็วมาก < 30ms ต่อตัวอักษร)
    const now = Date.now();
    const timeDiff = now - lastKeyTime;
    lastKeyTime = now;

    if (timeDiff < 30) {
        // เป็น QR Scanner — ใช้ timeout สั้น
        clearTimeout(scanTimeout);
        scanTimeout = setTimeout(() => {
            const id = event.target.value.trim();
            if (id.length > 3) {
                findUser(id);
                event.target.value = "";
            }
        }, 150);
    }
    // ถ้าพิมพ์มือ (timeDiff >= 30ms) — ให้กด Enter เอง
}

function manualSearch() {
    const id = document.getElementById('id-input').value.trim();
    if (id) {
        findUser(id);
        document.getElementById('id-input').value = "";
    }
    else { alert("กรุณากรอกรหัส REF"); }
}

function formatDateTime() {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, '0');
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const y = now.getFullYear();
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');
    return `${d}/${m}/${y} ${hh}:${mm}:${ss} (UTC +07:00)`;
}

function findUser(id) {
    if (!id) return;
    if (participantData.length === 0) {
        alert("กรุณาอัปโหลดไฟล์ Excel ก่อนทำการค้นหา");
        return;
    }

    const user = participantData.find(u => String(u["REF"]).trim() === id.trim());

    if (user) {
        lastId = id;
        user["Check Attendance"] = "Present";
        if (!user["Registration On"]) { user["Registration On"] = formatDateTime(); }
        saveToLocal();

        document.getElementById('placeholder-text').style.display = 'none';
        document.getElementById('card').style.display = 'block';

        document.getElementById('display-id').innerText = user["Order"] || 'N/A';
        document.getElementById('display-name').innerText = user["Name Surname"] || '-';
        document.getElementById('display-school').innerText = user["School Name"] || '-';
        document.getElementById('display-grade').innerText = user["Grade"] || '-';
        document.getElementById('display-subject').innerText = user["Subject"] || '-';
        document.getElementById('display-examvenue').innerText = user["Exam Venue"] || '-';
        document.getElementById('display-room').innerText = user["Room"] || '-';
        document.getElementById('display-floor').innerText = user["Floor"] || '-';
        document.getElementById('display-seat').innerText = user["Seat No."] || '-';

        if (clearTimer) clearTimeout(clearTimer);
        clearTimer = setTimeout(() => { clearResult(); lastId = null; }, 12000);
    } else {
        alert("ไม่พบข้อมูลในระบบ (REF: " + id + ")");
        document.getElementById('id-input').focus();
    }
}

function updateRealtimeTable() {
    const tbody = document.getElementById('table-body');
    const searchTerm = document.getElementById('realtime-search').value.toLowerCase();
    tbody.innerHTML = '';

    if (participantData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;">ยังไม่ได้อัปโหลดข้อมูลรายชื่อ</td></tr>';
        document.getElementById('stat-total').innerText = '0';
        document.getElementById('stat-present').innerText = '0';
        document.getElementById('stat-absent').innerText = '0';
        return;
    }

    const total = participantData.length;
    const presentCount = participantData.filter(u => u["Check Attendance"] === "Present").length;
    const absentCount = total - presentCount;
    document.getElementById('stat-total').innerText = total;
    document.getElementById('stat-present').innerText = presentCount;
    document.getElementById('stat-absent').innerText = absentCount;

    const filteredData = participantData.filter(user => {
        return Object.values(user).some(val => String(val).toLowerCase().includes(searchTerm));
    });

    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;">ไม่พบข้อมูลที่ค้นหา</td></tr>';
        return;
    }

    filteredData.forEach((user, index) => {
        const tr = document.createElement('tr');
        const statusClass = user["Check Attendance"] === "Present" ? "status-present" : "status-none";
        const statusText = user["Check Attendance"] === "Present" ? "Present" : "No record";

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td>${user["Order"] || '-'}</td>
            <td>${user["Name Surname"] || '-'}</td>
            <td>${user["School Name"] || '-'}</td>
            <td>${user["Grade"] || '-'} / ${user["Subject"] || '-'}</td>
            <td>${user["Registration On"] || "-"}</td>
            <td class="${statusClass}">${statusText}</td>
        `;
        tbody.appendChild(tr);
    });
}

function exportToExcel() {
    if (participantData.length === 0) return alert("ไม่มีข้อมูลสำหรับการดาวน์โหลด");
    const finalData = participantData.map(item => ({
        ...item,
        "Check Attendance": item["Check Attendance"] === "Present" ? "Present" : "Absent",
        "Registration On": item["Registration On"] || "-"
    }));
    const worksheet = XLSX.utils.json_to_sheet(finalData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance_Report");
    XLSX.writeFile(workbook, "Attendance_Report_2026.xlsx");
}

function clearResult() {
    document.getElementById('card').style.display = 'none';
    document.getElementById('placeholder-text').style.display = 'block';
    document.getElementById('id-input').focus();
}

window.onload = function() {
    loadFromLocal();
    loadNote();
    document.getElementById('id-input').focus();
};