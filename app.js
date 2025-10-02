// Data storage
let masterData = [];
let nilaiData = [];
let users = [];
let currentUser = null;
let currentPage = 1;
let pageSize = 20;
let filteredNilaiData = [];

// ===== Charts (Chart.js) =====
let _chartTop10 = null;
let _chartDistribusi = null;

// ===== Session (3 hari) =====
const SESSION_TTL_MS = 3 * 24 * 60 * 60 * 1000;

function saveSession(user) {
  const payload = { user, expiresAt: Date.now() + SESSION_TTL_MS };
  localStorage.setItem('sessionUser', JSON.stringify(payload));
}

function loadSession() {
  try {
    const raw = localStorage.getItem('sessionUser');
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj || !obj.expiresAt || obj.expiresAt < Date.now()) {
      localStorage.removeItem('sessionUser');
      return null;
    }
    return obj.user;
  } catch {
    return null;
  }
}

function clearSession() {
  localStorage.removeItem('sessionUser');
}


// Konfigurasi Google Apps Script
// GANTI URL_INI dengan URL Web App Google Apps Script Anda
const GAS_URL =
  "https://script.google.com/macros/s/AKfycbwu5QTgQ_S7g5OjnHpBrdxtICaZ1_ZQ2rV5dOU7iJjn1ENLn1CuPgKFQvwEbb98tbQ6/exec";

// Initialize application
document.addEventListener("DOMContentLoaded", function () {
    // Auto-login jika sesi valid
  const sessionUser = loadSession();
  if (sessionUser) {
    currentUser = sessionUser;
    document.getElementById("current-user").textContent = currentUser.username;
    const cuMobile = document.getElementById("current-user-mobile");
    if (cuMobile) cuMobile.textContent = currentUser.username;

    showPage("dashboard");
    updateUIForUserRole();
    initializeDefaultAdmin();
    loadDataFromStorage();
    setupEventListeners();
    initializeUI();
    return; // tidak perlu tampilkan modal login
  }

  // TIDAK ada sesi → Show login modal
  const loginModal = new bootstrap.Modal(document.getElementById("loginModal"));
  loginModal.show();

  initializeDefaultAdmin();
  loadDataFromStorage();
  setupEventListeners();
  initializeUI();
});

// Initialize default admin user
function initializeDefaultAdmin() {
  const defaultAdmin = {
    username: "admin",
    password: "user123",
    role: "admin",
    region: "",
    unit: "",
    status: "active",
  };

  // Save to localStorage if not exists
  if (!localStorage.getItem("users")) {
    users = [defaultAdmin];
    localStorage.setItem("users", JSON.stringify(users));
  }
}

// Load data from localStorage
function loadDataFromStorage() {
  // Load master data
  const storedMasterData = localStorage.getItem("masterData");
  if (storedMasterData) {
    masterData = normalizeMasterData(JSON.parse(storedMasterData));
  } else {
    masterData = normalizeMasterData(parseMasterData());
    localStorage.setItem("masterData", JSON.stringify(masterData));
  }

  // Load nilai data
  const storedNilaiData = localStorage.getItem("nilaiData");
  if (storedNilaiData) {
    nilaiData = JSON.parse(storedNilaiData);
  }

  // Load users
  const storedUsers = localStorage.getItem("users");
  if (storedUsers) {
    users = JSON.parse(storedUsers);
  }
}

// Parse the provided master data
function parseMasterData() {
  // This would parse the Excel data provided in the question
  // For now, we'll create a simplified version
  const data = [];

  // Sample data structure based on the provided Excel
  const sampleData = [
    {
      Unit: "KNPA",
      Region: "Kenepai",
      NIP: "24007",
      Nama: "GREGORIUS",
      Divisi: "01 DIV1",
      KodeJabatan: "006 MANDOR PRODUKSI",
      Grade: "",
    },
    {
      Unit: "KNPA",
      Region: "Kenepai",
      NIP: "24014",
      Nama: "VALENTINA ERNI",
      Divisi: "01 DIV1",
      KodeJabatan: "009 MANDOR PERAWATAN",
      Grade: "",
    },
    {
      Unit: "KNPA",
      Region: "Kenepai",
      NIP: "24015",
      Nama: "RINI",
      Divisi: "01 DIV1",
      KodeJabatan: "014 KERANI DIVISI",
      Grade: "",
    },
    // Add more sample data as needed
  ];

  return sampleData;
}

function normalizeMasterData(arr) {
  if (!Array.isArray(arr)) return [];
  return arr.map((it) => ({
    Unit: (it.Unit ?? "").toString().trim(),
    Region: (it.Region ?? "").toString().trim(),
    NIP: (it.NIP ?? "").toString().trim(),
    Nama: (it.Nama ?? "").toString().trim(),
    Divisi: (it.Divisi ?? "").toString().trim(),
    KodeJabatan: (it.KodeJabatan ?? "").toString().trim(),
    Grade: (it.Grade ?? "").toString().trim(),
  }));
}

// Set up event listeners
function setupEventListeners() {
  // Login form
  document.getElementById("login-form").addEventListener("submit", handleLogin);

  // Sidebar navigation
  document.querySelectorAll(".sidebar .nav-link").forEach((link) => {
    link.addEventListener("click", function (e) {
      e.preventDefault();
      const page = this.getAttribute("data-page");
      showPage(page);

      // Close sidebar on mobile after clicking a link
      if (window.innerWidth <= 768) {
        document.getElementById("sidebar").classList.remove("show");
      }
    });
  });

  // Mobile menu button
  document
    .querySelector(".mobile-menu-btn")
    .addEventListener("click", function () {
      document.getElementById("sidebar").classList.toggle("show");
    });

  // Logout button
  document.getElementById("logout-btn").addEventListener("click", handleLogout);

  // Input form
  document
    .getElementById("input-form")
    .addEventListener("submit", handleInputFormSubmit);
  document
    .getElementById("input-nilai-isian")
    .addEventListener("input", calculateTotalNilai);
  document
    .getElementById("input-nilai-bkm")
    .addEventListener("input", calculateTotalNilai);

  // Autocomplete for nama input
  document
    .getElementById("input-nama")
    .addEventListener("input", handleNamaAutocomplete);
  document
    .getElementById("input-nip")
    .addEventListener("input", handleNipAutocomplete);

  // Report page
  document
    .getElementById("search-report")
    .addEventListener("input", filterReportTable);
  document.getElementById("page-size").addEventListener("change", function () {
    pageSize = parseInt(this.value);
    currentPage = 1;
    renderReportTable();
  });
  document
    .getElementById("select-all")
    .addEventListener("change", toggleSelectAll);
  document
    .getElementById("export-excel")
    .addEventListener("click", exportToExcel);
  document
    .getElementById("sync-selected")
    .addEventListener("click", syncSelectedData);
  document.getElementById("sync-all").addEventListener("click", syncAllData);

  // Sync page
  document
    .getElementById("pull-master-data")
    .addEventListener("click", pullMasterData);
  document
    .getElementById("upload-master-data")
    .addEventListener("click", uploadMasterData);
  const clearBtn = document.getElementById("clear-local-data");
  if (clearBtn) clearBtn.addEventListener("click", clearLocalDataWithPassword);
  setupUserManagementEvents();

  // Setting page
  document
    .getElementById("add-user-btn")
    .addEventListener("click", showAddUserModal);
  document.getElementById("save-user").addEventListener("click", saveUser);

  // Dashboard filters
  document
    .getElementById("filter-region")
    .addEventListener("change", updateDashboard);
  document
    .getElementById("filter-unit")
    .addEventListener("change", updateDashboard);
  document
    .getElementById("filter-jabatan")
    .addEventListener("change", updateDashboard);
  document
    .getElementById("filter-count")
    .addEventListener("change", updateDashboard);
}

// Initialize UI
function initializeUI() {
  // Populate filter dropdowns
  populateFilterDropdowns();

  // Update dashboard
  updateDashboard();

  // Update sync info
  updateSyncInfo();
}

// Fungsi login yang terintegrasi dengan GAS
async function handleLogin(e) {
  e.preventDefault();
  showSpinner();

  const username = document.getElementById("login-username").value;
  const password = document.getElementById("login-password").value;

  try {
    const response = await fetch(GAS_URL, {
      method: "POST",
      headers: {
        // PENTING: pakai text/plain agar tidak memicu preflight CORS
        "Content-Type": "text/plain;charset=utf-8",
      },
      body: JSON.stringify({
        action: "login",
        username: username,
        password: password,
      }),
    });

    const result = await response.json();

    if (result.status === "success") {
      currentUser = result.data;
      document.getElementById("current-user").textContent =
        currentUser.username;

        const cuMobile = document.getElementById("current-user-mobile");
        if (cuMobile) cuMobile.textContent = currentUser.username;

      // simpan sesi 3 hari
        saveSession(currentUser);

      // Hide login modal
      const loginModal = bootstrap.Modal.getInstance(
        document.getElementById("loginModal")
      );
      loginModal.hide();

      // Show dashboard
      showPage("dashboard");

      // Update UI based on user role
      updateUIForUserRole();

      Swal.fire({
        icon: "success",
        title: "Login Berhasil",
        text: `Selamat datang ${currentUser.username}!`,
        timer: 2000,
        showConfirmButton: false,
      });
    } else {
      Swal.fire({
        icon: "error",
        title: "Login Gagal",
        text: result.message,
      });
    }
  } catch (error) {
    console.error("Login error:", error);
    Swal.fire({
      icon: "error",
      title: "Login Gagal",
      text: "Terjadi kesalahan saat login: " + error.message,
    });
  } finally {
    hideSpinner();
  }
}

// Update event listener login form
document.getElementById("login-form").addEventListener("submit", handleLogin);

// Handle logout
function handleLogout() {
  clearSession();
  currentUser = null;
  document.getElementById("login-username").value = "";
  document.getElementById("login-password").value = "";

  const loginModal = new bootstrap.Modal(document.getElementById("loginModal"));
  loginModal.show();
}

// Show page
function showPage(pageName) {
  // Hide all pages
  document.querySelectorAll(".page").forEach((page) => {
    page.classList.add("d-none");
  });

  // Remove active class from all nav links
  document.querySelectorAll(".sidebar .nav-link").forEach((link) => {
    link.classList.remove("active");
  });

  // Show selected page
  document.getElementById(`${pageName}-page`).classList.remove("d-none");

  // Add active class to selected nav link
  document
    .querySelector(`.sidebar .nav-link[data-page="${pageName}"]`)
    .classList.add("active");

  // Update page-specific content
  if (pageName === "report") {
    renderReportTable();
  } else if (pageName === "setting") {
    renderUserTable();
  }
}

// Update UI based on user role
function updateUIForUserRole() {
  if (currentUser.role === "kerani") {
    // Hide admin-only elements
    document.querySelectorAll(".admin-only").forEach((el) => {
      el.style.display = "none";
    });

    // Apply region/unit filters if specified
    if (currentUser.region) {
      document.getElementById("filter-region").value = currentUser.region;
    }
    if (currentUser.unit) {
      document.getElementById("filter-unit").value = currentUser.unit;
    }

    updateDashboard();
  } else {
    // Show admin-only elements
    document.querySelectorAll(".admin-only").forEach((el) => {
      el.style.display = "block";
    });
  }
}

// Populate filter dropdowns
function populateFilterDropdowns() {
  const regionSelect = document.getElementById("filter-region");
  const unitSelect = document.getElementById("filter-unit");
  const jabatanSelect = document.getElementById("filter-jabatan");

  // Get unique values from master data
  const regions = [...new Set(masterData.map((item) => item.Region))].filter(
    Boolean
  );
  const units = [...new Set(masterData.map((item) => item.Unit))].filter(
    Boolean
  );
  const jabatans = [
    ...new Set(masterData.map((item) => item.KodeJabatan)),
  ].filter(Boolean);

  // Populate region dropdown
  regions.forEach((region) => {
    const option = document.createElement("option");
    option.value = region;
    option.textContent = region;
    regionSelect.appendChild(option);
  });

  // Populate unit dropdown
  units.forEach((unit) => {
    const option = document.createElement("option");
    option.value = unit;
    option.textContent = unit;
    unitSelect.appendChild(option);
  });

  // Populate jabatan dropdown
  jabatans.forEach((jabatan) => {
    const option = document.createElement("option");
    option.value = jabatan;
    option.textContent = jabatan;
    jabatanSelect.appendChild(option);
  });

  // Also populate user form dropdowns
  const userRegionSelect = document.getElementById("user-region");
  const userUnitSelect = document.getElementById("user-unit");

  regions.forEach((region) => {
    const option = document.createElement("option");
    option.value = region;
    option.textContent = region;
    userRegionSelect.appendChild(option);
  });

  units.forEach((unit) => {
    const option = document.createElement("option");
    option.value = unit;
    option.textContent = unit;
    userUnitSelect.appendChild(option);
  });
}

// Handle input form submission
async function handleInputFormSubmit(e) {
  e.preventDefault();
  showSpinner();

  const newItem = {
    id: Date.now().toString(),
    nip: document.getElementById("input-nip").value.trim(),
    nama: document.getElementById("input-nama").value.trim(),
    region: document.getElementById("input-region").value.trim(),
    unit: document.getElementById("input-unit").value.trim(),
    divisi: document.getElementById("input-divisi").value.trim(),
    jabatan: document.getElementById("input-jabatan").value.trim(),
    grade: document.getElementById("input-grade").value.trim(),
    nilaiIsian: parseInt(document.getElementById("input-nilai-isian").value),
    nilaiBKM: parseInt(document.getElementById("input-nilai-bkm").value),
    totalNilai: parseInt(document.getElementById("input-total-nilai").value),
    inputBy: currentUser?.username || '',
    timestamp: new Date().toISOString(),
    synced: false
  };

  // Validasi dasar
  if (!newItem.nip || !newItem.nama || isNaN(newItem.nilaiIsian) || isNaN(newItem.nilaiBKM)) {
    hideSpinner();
    await Swal.fire({ icon: "error", title: "Data Tidak Lengkap", text: "Harap isi semua field yang wajib!" });
    return;
  }

  // Cek duplikasi by NIP
  const existIdx = nilaiData.findIndex(x => x.nip === newItem.nip && x.unit === newItem.unit && x.region === newItem.region);
  if (existIdx >= 0) {
    hideSpinner();
    const exist = nilaiData[existIdx];
    const { isConfirmed } = await Swal.fire({
      icon: "warning",
      title: "Data sudah ada",
      html: `
        <div class="text-start">
          <p>Data untuk NIP <b>${exist.nip}</b> (${exist.nama}) sudah pernah diinput.</p>
          <ul class="small">
            <li>Total lama: <b>${exist.totalNilai}</b> (${exist.nilaiIsian}+${exist.nilaiBKM})</li>
            <li>Status sync: <b>${exist.synced ? 'Synced' : 'Belum Sync'}</b></li>
          </ul>
          <p>Apakah Anda ingin <b>mengganti/mengedit</b> data tersebut dengan nilai baru?</p>
        </div>
      `,
      showCancelButton: true,
      confirmButtonText: "Ya, ganti",
      cancelButtonText: "Tidak"
    });
    if (!isConfirmed) return;

    // Jika ganti: gunakan ID lama agar upsert ke baris yang sama saat sync
    newItem.id = exist.id;
    newItem.synced = false;         // perubahan lokal → perlu sync ulang
    nilaiData[existIdx] = newItem;

  } else {
    // Baru
    nilaiData.push(newItem);
  }

  localStorage.setItem("nilaiData", JSON.stringify(nilaiData));

  // Reset form
  document.getElementById("input-form").reset();
  document.getElementById("input-total-nilai").value = "";

  hideSpinner();

  await Swal.fire({
    icon: "success",
    title: "Data Tersimpan",
    text: "Data nilai berhasil disimpan!",
    // opsional: auto close agar cepat kembali input
    // timer: 1400,
    // showConfirmButton: false,
    didClose: () => {
      // Pindahkan fokus saat modal benar-benar sudah tertutup
      focusNipField();
    }
  });

  // Fallback tambahan kalau user menutup via tombol & event timing berbeda
  focusNipField();

  // Refresh dashboard/report
  renderReportTable();
  updateDashboard();
}


// Calculate total nilai
function calculateTotalNilai() {
  const nilaiIsian =
    parseInt(document.getElementById("input-nilai-isian").value) || 0;
  const nilaiBKM =
    parseInt(document.getElementById("input-nilai-bkm").value) || 0;
  const totalNilai = nilaiIsian + nilaiBKM;

  document.getElementById("input-total-nilai").value = totalNilai;
}

// Autocomplete Nama
function handleNamaAutocomplete() {
  const input = document.getElementById("input-nama");
  const value = (input.value || "").toLowerCase();
  const list = document.getElementById("autocomplete-nama-list");
  list.innerHTML = "";
  if (value.length < 2) return;

  const matches = masterData
    .filter((item) => item.Nama && item.Nama.toLowerCase().includes(value))
    .slice(0, 10);

  matches.forEach((item) => {
    const div = document.createElement("div");
    div.textContent = `${item.Nama} (${item.NIP}) - ${item.Unit}`;
    div.addEventListener("click", function () {
      input.value = item.Nama;
      document.getElementById("input-nip").value = item.NIP;
      document.getElementById("input-region").value = item.Region || "";
      document.getElementById("input-unit").value = item.Unit || "";
      document.getElementById("input-divisi").value = item.Divisi || "";
      document.getElementById("input-jabatan").value = item.KodeJabatan || "";
      document.getElementById("input-grade").value = item.Grade || "";
      list.innerHTML = "";
    });
    list.appendChild(div);
  });
}

// Autocomplete NIP
function handleNipAutocomplete() {
  const input = document.getElementById("input-nip");
  const raw = input.value || "";
  const value = raw.toString().trim();
  const list = document.getElementById("autocomplete-nip-list");
  list.innerHTML = "";
  if (value.length < 2) return;

  const matches = masterData
    .filter((item) => {
      const nipStr = (item.NIP ?? "").toString();
      return nipStr.includes(value);
    })
    .slice(0, 10);

  matches.forEach((item) => {
    const div = document.createElement("div");
    div.textContent = `${item.NIP} - ${item.Nama} - ${item.Unit}`;
    div.addEventListener("click", function () {
      input.value = item.NIP;
      document.getElementById("input-nama").value = item.Nama;
      document.getElementById("input-region").value = item.Region || "";
      document.getElementById("input-unit").value = item.Unit || "";
      document.getElementById("input-divisi").value = item.Divisi || "";
      document.getElementById("input-jabatan").value = item.KodeJabatan || "";
      document.getElementById("input-grade").value = item.Grade || "";
      list.innerHTML = "";
    });
    list.appendChild(div);
  });
}

// Helper: aman-destroy chart lama
function destroyChartIfAny(ref) {
  if (ref && typeof ref.destroy === 'function') {
    ref.destroy();
  }
}


// Render report table
function renderReportTable() {
  const tableBody = document
    .getElementById("report-table")
    .querySelector("tbody");
  tableBody.innerHTML = "";

  // Apply filters if any
  filteredNilaiData = [...nilaiData];
  const searchTerm = document
    .getElementById("search-report")
    .value.toLowerCase();

  if (searchTerm) {
    filteredNilaiData = filteredNilaiData.filter((item) =>
      Object.values(item).some(
        (val) => val && val.toString().toLowerCase().includes(searchTerm)
      )
    );
  }

  // Apply pagination
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, filteredNilaiData.length);
  const pageData = filteredNilaiData.slice(startIndex, endIndex);

  // Render table rows
  pageData.forEach((item) => {
    const row = document.createElement("tr");
    row.innerHTML = `
                    <td><input type="checkbox" class="row-select" data-id="${
                      item.id
                    }"></td>
                    <td>${item.nip}</td>
                    <td>${item.nama}</td>
                    <td>${item.region}</td>
                    <td>${item.unit}</td>
                    <td>${item.divisi}</td>
                    <td>${item.jabatan}</td>
                    <td>${item.grade}</td>
                    <td>${item.nilaiIsian}</td>
                    <td>${item.nilaiBKM}</td>
                    <td>${item.totalNilai}</td>
                    <td>
                        <span class="sync-status ${
                          item.synced ? "synced" : "not-synced"
                        }"></span>
                        ${item.synced ? "Synced" : "Not Synced"}
                    </td>
                    <td>${item.inputBy}</td>
                    <td>${new Date(item.timestamp).toLocaleString("id-ID")}</td>
                    <td>
                        <button class="btn btn-sm btn-warning edit-btn" data-id="${
                          item.id
                        }">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn btn-sm btn-danger delete-btn" data-id="${
                          item.id
                        }">
                            <i class="fas fa-trash"></i>
                        </button>
                    </td>
                `;
    tableBody.appendChild(row);
  });

  // Add event listeners to action buttons
  document.querySelectorAll(".edit-btn").forEach((btn) => {
    btn.addEventListener("click", function () {
      const id = this.getAttribute("data-id");
      editNilaiData(id);
    });
  });

  document.querySelectorAll(".delete-btn").forEach((btn) => {
    btn.addEventListener("click", function () {
      const id = this.getAttribute("data-id");
      deleteNilaiData(id);
    });
  });

  // Render pagination
  renderPagination();
}

// Render pagination
function renderPagination() {
  const pagination = document.getElementById("pagination");
  pagination.innerHTML = "";

  const totalPages = Math.ceil(filteredNilaiData.length / pageSize);

  // Previous button
  const prevLi = document.createElement("li");
  prevLi.className = `page-item ${currentPage === 1 ? "disabled" : ""}`;
  prevLi.innerHTML = `<a class="page-link" href="#">Previous</a>`;
  prevLi.addEventListener("click", function (e) {
    e.preventDefault();
    if (currentPage > 1) {
      currentPage--;
      renderReportTable();
    }
  });
  pagination.appendChild(prevLi);

  // Page numbers
  for (let i = 1; i <= totalPages; i++) {
    const li = document.createElement("li");
    li.className = `page-item ${i === currentPage ? "active" : ""}`;
    li.innerHTML = `<a class="page-link" href="#">${i}</a>`;
    li.addEventListener("click", function (e) {
      e.preventDefault();
      currentPage = i;
      renderReportTable();
    });
    pagination.appendChild(li);
  }

  // Next button
  const nextLi = document.createElement("li");
  nextLi.className = `page-item ${
    currentPage === totalPages ? "disabled" : ""
  }`;
  nextLi.innerHTML = `<a class="page-link" href="#">Next</a>`;
  nextLi.addEventListener("click", function (e) {
    e.preventDefault();
    if (currentPage < totalPages) {
      currentPage++;
      renderReportTable();
    }
  });
  pagination.appendChild(nextLi);
}

// Filter report table
function filterReportTable() {
  currentPage = 1;
  renderReportTable();
}

// Toggle select all
function toggleSelectAll() {
  const checked = document.getElementById("select-all").checked;
  document.querySelectorAll(".row-select").forEach((checkbox) => {
    checkbox.checked = checked;
  });
}

// Edit nilai data
function editNilaiData(id) {
  const item = nilaiData.find((item) => item.id === id);
  if (!item) return;

  // Populate input form with data
  document.getElementById("input-nip").value = item.nip;
  document.getElementById("input-nama").value = item.nama;
  document.getElementById("input-region").value = item.region;
  document.getElementById("input-unit").value = item.unit;
  document.getElementById("input-divisi").value = item.divisi;
  document.getElementById("input-jabatan").value = item.jabatan;
  document.getElementById("input-grade").value = item.grade;
  document.getElementById("input-nilai-isian").value = item.nilaiIsian;
  document.getElementById("input-nilai-bkm").value = item.nilaiBKM;
  document.getElementById("input-total-nilai").value = item.totalNilai;

  // Remove the item from nilaiData
  nilaiData = nilaiData.filter((item) => item.id !== id);
  localStorage.setItem("nilaiData", JSON.stringify(nilaiData));

  // Switch to input page
  showPage("input");

  Swal.fire({
    icon: "info",
    title: "Data Diedit",
    text: "Data telah dimuat ke form. Silakan edit dan simpan kembali.",
  });
}

// Delete nilai data
function deleteNilaiData(id) {
  Swal.fire({
    title: "Hapus Data?",
    text: "Data yang dihapus tidak dapat dikembalikan!",
    icon: "warning",
    showCancelButton: true,
    confirmButtonColor: "#d33",
    cancelButtonColor: "#3085d6",
    confirmButtonText: "Ya, Hapus!",
    cancelButtonText: "Batal",
  }).then((result) => {
    if (result.isConfirmed) {
      nilaiData = nilaiData.filter((item) => item.id !== id);
      localStorage.setItem("nilaiData", JSON.stringify(nilaiData));
      renderReportTable();

      Swal.fire("Terhapus!", "Data telah dihapus.", "success");
    }
  });
}

// Export to Excel
function exportToExcel() {
  // Create worksheet
  const ws = XLSX.utils.json_to_sheet(
    nilaiData.map((item) => ({
      NIP: item.nip,
      Nama: item.nama,
      Region: item.region,
      Unit: item.unit,
      Divisi: item.divisi,
      "Kode Jabatan": item.jabatan,
      Grade: item.grade,
      "Nilai Isian": item.nilaiIsian,
      "Nilai BKM": item.nilaiBKM,
      "Total Nilai": item.totalNilai,
      "Input By": item.inputBy,
      Timestamp: new Date(item.timestamp).toLocaleString("id-ID"),
    }))
  );

  // Create workbook
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data Nilai");

  // Export to file
  XLSX.writeFile(wb, "data_nilai_uji_kompetensi.xlsx");
}

// Fungsi sync yang sebenarnya ke Google Sheets
async function syncDataToGoogleSheets(dataToSync) {
  try {
    showProgressModal("Mengirim data ke Google Sheet...");

    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < dataToSync.length; i++) {
      const item = dataToSync[i];
      const progress = Math.round(((i + 1) / dataToSync.length) * 100);
      updateProgress(progress);

      try {
        // Kirim data ke Google Apps Script
        const response = await fetch(GAS_URL, {
          method: "POST",
          headers: {
            "Content-Type": "text/plain;charset=utf-8",
          },
          body: JSON.stringify({
            action: "pushScore",
            row: item,
          }),
        });

        const result = await response.json();
        if (result.status === "success") {
          const itemIndex = nilaiData.findIndex((n) => n.id === item.id);
          if (itemIndex !== -1) {
            nilaiData[itemIndex].synced = true;
            nilaiData[itemIndex].syncedAt = new Date().toISOString();
            successCount++;
          }
        } else {
          errorCount++;
          console.error("Sync error:", result.message);
        }
      } catch (error) {
        errorCount++;
        console.error("Network error:", error);
      }
    }

    // Simpan perubahan ke local storage
    localStorage.setItem("nilaiData", JSON.stringify(nilaiData));

    hideProgressModal();

    // Tampilkan hasil
    if (errorCount === 0) {
      Swal.fire({
        icon: "success",
        title: "Sync Berhasil",
        text: `Semua data (${successCount}) berhasil disinkronisasi!`,
      });
    } else {
      Swal.fire({
        icon: "warning",
        title: "Sync Sebagian Berhasil",
        text: `${successCount} data berhasil, ${errorCount} gagal disinkronisasi.`,
      });
    }

    renderReportTable();
    updateSyncInfo();
  } catch (error) {
    hideProgressModal();
    Swal.fire({
      icon: "error",
      title: "Sync Gagal",
      text: "Terjadi kesalahan saat sinkronisasi: " + error.message,
    });
  }
}

// Sync selected data
function syncSelectedData() {
  const selectedIds = [];
  document.querySelectorAll(".row-select:checked").forEach((checkbox) => {
    selectedIds.push(checkbox.getAttribute("data-id"));
  });

  if (selectedIds.length === 0) {
    Swal.fire({
      icon: "warning",
      title: "Tidak Ada Data Dipilih",
      text: "Pilih data yang akan disinkronisasi!",
    });
    return;
  }

  const selectedData = nilaiData.filter(
    (item) => selectedIds.includes(item.id) && !item.synced
  );

  if (selectedData.length === 0) {
    Swal.fire({
      icon: "info",
      title: "Data Sudah Disinkronisasi",
      text: "Data yang dipilih sudah disinkronisasi sebelumnya.",
    });
    return;
  }

  syncDataToGoogleSheets(selectedData);
}

// Sync all data
function syncAllData() {
  const unsyncedData = nilaiData.filter((item) => !item.synced);

  if (unsyncedData.length === 0) {
    Swal.fire({
      icon: "info",
      title: "Tidak Ada Data Baru",
      text: "Semua data sudah disinkronisasi!",
    });
    return;
  }

  syncDataToGoogleSheets(unsyncedData);
}

// Fungsi untuk menarik master data dari Google Sheet
async function pullMasterDataFromGoogleSheets() {
  try {
    showProgressModal("Menarik master & data aktual dari Google Sheet...");

    // ===== 1) MASTER
    const qsMaster = new URLSearchParams({ action: "getMasterData" });
    if (currentUser?.region) qsMaster.set("region", currentUser.region);
    if (currentUser?.unit)   qsMaster.set("unit",   currentUser.unit);

    const respMaster = await fetch(`${GAS_URL}?${qsMaster.toString()}`);
    if (!respMaster.ok) throw new Error(`HTTP Master ${respMaster.status}`);
    const resMaster = await respMaster.json();
    if (!(resMaster.status === "success" && Array.isArray(resMaster.data))) {
      throw new Error(resMaster.message || "Gagal ambil master");
    }
    masterData = normalizeMasterData(resMaster.data);
    localStorage.setItem("masterData", JSON.stringify(masterData));
    localStorage.setItem("masterDataLastUpdate", new Date().toISOString());

    // ===== 2) SCORES (data aktual)
    updateProgress(35);
    const qsScores = new URLSearchParams({ action: "getScores" });
    if (currentUser?.region) qsScores.set("region", currentUser.region);
    if (currentUser?.unit)   qsScores.set("unit",   currentUser.unit);

    const respScores = await fetch(`${GAS_URL}?${qsScores.toString()}`);
    if (!respScores.ok) throw new Error(`HTTP Scores ${respScores.status}`);
    const resScores = await respScores.json();
    if (!(resScores.status === "success" && Array.isArray(resScores.data))) {
      throw new Error(resScores.message || "Gagal ambil scores");
    }

    // Map data Scores dari server → format nilaiData frontend
    const serverItems = resScores.data.map(r => ({
      id:        String(r._id || r.id),
      nip:       String(r.NIP || ''),
      nama:      String(r.Nama || ''),
      region:    String(r.Region || ''),
      unit:      String(r.Unit || ''),
      divisi:    String(r.Divisi || ''),
      jabatan:   String(r.KodeJabatan || ''),
      grade:     String(r.Grade || ''),
      nilaiIsian: Number(r.NilaiIsian || 0),
      nilaiBKM:   Number(r.NilaiBKM   || 0),
      totalNilai: Number(r.Total      || 0),
      inputBy:    String(r.createdBy  || ''),
      timestamp:  String(r.createdAt  || new Date().toISOString()),
      synced:     true,                         // penting: tandai sebagai sudah sinkron
      syncedAt:   r.syncAt ? String(r.syncAt) : new Date().toISOString()
    }));

    updateProgress(65);

    // Merge: pertahankan item lokal yang belum synced, hindari duplikat id
    const localUnsynced = (nilaiData || []).filter(x => !x.synced);
    const byId = new Map();
    serverItems.forEach(it => byId.set(it.id, it));
    localUnsynced.forEach(it => {
      if (!byId.has(it.id)) byId.set(it.id, it);
    });

    nilaiData = Array.from(byId.values());
    localStorage.setItem("nilaiData", JSON.stringify(nilaiData));

    hideProgressModal();

    Swal.fire({
      icon: "success",
      title: "Data Terunduh",
      text: `Master: ${masterData.length} baris, Scores: ${serverItems.length} baris (sinkron).`
    });

    updateSyncInfo();
    // Bersihkan & isi ulang dropdown filter
    ["filter-region","filter-unit","filter-jabatan","user-region","user-unit"].forEach(id => {
      const el = document.getElementById(id);
      if (el) for (let i = el.options.length - 1; i >= 1; i--) el.remove(i);
    });
    populateFilterDropdowns();

    // Refresh Report & Dashboard
    renderReportTable();
    updateDashboard();
  } catch (error) {
    hideProgressModal();
    console.error("Error pulling data:", error);
    Swal.fire({ icon: "error", title: "Gagal Tarik Data", text: String(error.message) });
  }
}


// Update fungsi pullMasterData
function pullMasterData() {
  pullMasterDataFromGoogleSheets();
}

// Fungsi untuk upload master data dari Excel
function uploadMasterData() {
  const fileInput = document.getElementById("upload-master-file");
  const file = fileInput.files[0];

  if (!file) {
    Swal.fire({
      icon: "warning",
      title: "Pilih File",
      text: "Pilih file Excel master data!",
    });
    return;
  }

  showProgressModal("Mengupload dan memproses file master data...");

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, {
        type: "array",
      });

      // Ambil sheet pertama
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Konversi ke JSON
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
      });

      // Proses data (header di baris pertama)
      const headers = jsonData[0];
      const rows = jsonData.slice(1);

      // Mapping data ke format yang diinginkan
      masterData = rows
        .map((row) => {
          const item = {};
          headers.forEach((header, index) => {
            if (header && row[index] !== undefined) {
              // Mapping header ke field yang sesuai
              switch (header.toLowerCase()) {
                case "unit":
                  item.Unit = row[index];
                  break;
                case "region":
                  item.Region = row[index];
                  break;
                case "nip":
                  item.NIP = row[index]?.toString();
                  break;
                case "nama":
                  item.Nama = row[index];
                  break;
                case "divisi":
                  item.Divisi = row[index];
                  break;
                case "kode jabatan":
                case "kode_jabatan":
                case "kodejabatan":
                  item.KodeJabatan = row[index];
                  break;
                case "grade":
                  item.Grade = row[index];
                  break;
                default:
                  item[header] = row[index];
              }
            }
          });
          return item;
        })
        .filter((item) => item.NIP && item.Nama); // Hanya data yang valid

      // Simpan ke localStorage
      localStorage.setItem("masterData", JSON.stringify(masterData));
      localStorage.setItem("masterDataLastUpdate", new Date().toISOString());

      hideProgressModal();

      Swal.fire({
        icon: "success",
        title: "Upload Berhasil",
        text: `Berhasil memproses ${masterData.length} data master dari file Excel!`,
      });

      updateSyncInfo();
      populateFilterDropdowns();
    } catch (error) {
      hideProgressModal();
      Swal.fire({
        icon: "error",
        title: "Upload Gagal",
        text: "Terjadi kesalahan saat memproses file: " + error.message,
      });
    }
  };

  reader.onerror = function () {
    hideProgressModal();
    Swal.fire({
      icon: "error",
      title: "Upload Gagal",
      text: "Gagal membaca file!",
    });
  };

  reader.readAsArrayBuffer(file);
}

// Update sync info
function updateSyncInfo() {
  document.getElementById("master-data-count").textContent = masterData.length;
  document.getElementById("nilai-data-count").textContent = nilaiData.length;

  const unsyncedCount = nilaiData.filter((item) => !item.synced).length;
  document.getElementById("unsynced-data-count").textContent = unsyncedCount;

  const lastUpdate = localStorage.getItem("masterDataLastUpdate") || "-";
  document.getElementById("master-data-last-update").textContent = lastUpdate;
}

// Fungsi untuk sync users ke Google Sheet
async function syncUsersToGoogleSheets(patch) {
  try {
    showProgressModal("Menyinkronisasi data user ke Google Sheet...");

    const toSend = Array.isArray(patch) && patch.length ? patch : users;

    const response = await fetch(GAS_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "syncUsers", users: toSend })
    });

    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
    const result = await response.json();

    hideProgressModal();
    if (result.status === "success") return true;
    throw new Error(result.message || "Gagal sync users");
  } catch (error) {
    hideProgressModal();
    console.error("Error syncing users:", error);
    Swal.fire({ icon: "error", title: "Sync User Gagal", text: "Terjadi kesalahan: " + error.message });
    return false;
  }
}

// =====  clearLocalDataWithPassword =====
async function clearLocalDataWithPassword() {
  if (!currentUser || !currentUser.username) {
    await Swal.fire({ icon: 'error', title: 'Tidak Bisa', text: 'Silakan login dahulu.' });
    return;
  }

  // 1) Verifikasi password via GAS
  const { value: verifiedPwd } = await Swal.fire({
    title: 'Verifikasi Password',
    html: `
      <p class="mb-1">Masukkan password untuk menghapus data lokal:</p>
      <p class="small text-muted mb-2">
        Data yang akan dihapus: <strong>Master Data</strong>, <strong>Data Nilai</strong>, dan
        <strong>Master Last Update</strong>.
      </p>
    `,
    input: 'password',
    inputPlaceholder: 'Password',
    inputAttributes: { autocapitalize: 'off', autocorrect: 'off' },
    showCancelButton: true,
    confirmButtonText: 'Verifikasi',
    cancelButtonText: 'Batal',
    showLoaderOnConfirm: true,
    allowOutsideClick: () => !Swal.isLoading(),
    preConfirm: async (pwd) => {
      try {
        if (!pwd) throw new Error('Password wajib diisi');
        const resp = await fetch(GAS_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ action: 'login', username: currentUser.username, password: pwd })
        });
        const res = await resp.json();
        if (res.status !== 'success') throw new Error(res.message || 'Password salah');
        return pwd;
      } catch (err) {
        Swal.showValidationMessage(err.message || 'Verifikasi gagal');
        return false;
      }
    }
  });
  if (!verifiedPwd) return;

  // 2) Pilihan mode hapus
  const mode = await Swal.fire({
    icon: 'warning',
    title: 'Hapus Data Lokal?',
    html: `
      <div class="text-start">
        <p>Anda dapat memilih:</p>
        <ul class="mb-0">
          <li><b>Hapus Saja</b> – hanya mengosongkan data lokal.</li>
          <li><b>Hapus & Tarik Ulang</b> – setelah kosong, sistem akan menarik kembali <i>Master</i> dan <i>Scores</i> dari Google Sheet.</li>
        </ul>
      </div>
    `,
    showCancelButton: true,
    showDenyButton: true,
    confirmButtonText: 'Hapus & Tarik Ulang',
    denyButtonText: 'Hapus Saja',
    cancelButtonText: 'Batal'
  });
  if (mode.isDismissed) return;

  const rePull = mode.isConfirmed; // true jika "Hapus & Tarik Ulang"

  // 3) Eksekusi hapus + progress modal
  try {
    // Hapus storage
    localStorage.removeItem('masterData');
    localStorage.removeItem('masterDataLastUpdate');
    localStorage.removeItem('nilaiData');

    // Reset variabel runtime
    masterData = [];
    nilaiData  = [];

    // Reset halaman & refresh UI kosong
    currentPage = 1;
    updateSyncInfo();
    renderReportTable();
    updateDashboard();

    if (!rePull) {
      await Swal.fire({ icon: 'success', title: 'Berhasil', text: 'Data lokal telah dihapus.' });
    }

  } catch (err) {
    await Swal.fire({ icon: 'error', title: 'Gagal', text: 'Terjadi kesalahan saat menghapus data lokal.' });
    return;
  }

  // 4) Opsional: tarik ulang (Master + Scores) agar Report/Dashboard terisi lagi
  if (rePull) {
    // fungsi ini sudah menampilkan Progress Modal sendiri dan akan update UI
    await pullMasterDataFromGoogleSheets();
  }
}


// Fungsi untuk load users dari Google Sheet
async function loadUsersFromGoogleSheets() {
  try {
    const response = await fetch(`${GAS_URL}?action=getUsers`);

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const result = await response.json();

    if (result.status === "success" && result.data) {
      users = result.data;
      localStorage.setItem("users", JSON.stringify(users));
      return true;
    } else {
      throw new Error(result.message || "Gagal mengambil data user");
    }
  } catch (error) {
    console.error("Error loading users:", error);
    // Fallback ke local storage jika gagal
    const storedUsers = localStorage.getItem("users");
    if (storedUsers) {
      users = JSON.parse(storedUsers);
    }
    return false;
  }
}

// Show add user modal
function showAddUserModal() {
  document.getElementById("userModalLabel").textContent = "Tambah User";
  document.getElementById("user-form").reset();
  document.getElementById("user-id").value = "";
  document.getElementById("user-username").readOnly = false;

  const userModal = new bootstrap.Modal(document.getElementById("userModal"));
  userModal.show();
}

// Render user table dengan load dari Google Sheet
async function renderUserTable() {
  showSpinner();

  try {
    // Load users dari Google Sheet terlebih dahulu
    await loadUsersFromGoogleSheets();

    const tableBody = document
      .getElementById("user-table")
      .querySelector("tbody");
    tableBody.innerHTML = "";

    if (users.length === 0) {
      tableBody.innerHTML = `
                <tr>
                    <td colspan="6" class="text-center text-muted py-4">
                        <i class="fas fa-users fa-2x mb-3"></i>
                        <p>Tidak ada data user</p>
                    </td>
                </tr>
            `;
    } else {
      users.forEach((user) => {
        const row = document.createElement("tr");
        row.innerHTML = `
                    <td>
                        <div class="d-flex align-items-center">
                            <i class="fas fa-user-circle me-2 text-muted"></i>
                            ${user.username}
                            ${
                              user.username === "admin"
                                ? '<span class="badge bg-primary ms-2">Admin</span>'
                                : ""
                            }
                        </div>
                    </td>
                    <td>
                        <span class="badge ${
                          user.role === "admin" ? "bg-primary" : "bg-info"
                        }">
                            ${user.role}
                        </span>
                    </td>
                    <td>${
                      user.region || '<span class="text-muted">Semua</span>'
                    }</td>
                    <td>${
                      user.unit || '<span class="text-muted">Semua</span>'
                    }</td>
                    <td>
                        <span class="badge ${
                          user.status === "active"
                            ? "bg-success"
                            : "bg-secondary"
                        }">
                            <i class="fas ${
                              user.status === "active" ? "fa-check" : "fa-times"
                            } me-1"></i>
                            ${user.status === "active" ? "Active" : "Inactive"}
                        </span>
                    </td>
                    <td>
                        <div class="btn-group btn-group-sm">
                            <button class="btn btn-warning edit-user-btn" data-id="${
                              user.username
                            }" 
                                ${user.username === "admin" ? "disabled" : ""}
                                title="Edit User">
                                <i class="fas fa-edit"></i>
                            </button>
                            <button class="btn btn-danger delete-user-btn" data-id="${
                              user.username
                            }" 
                                ${user.username === "admin" ? "disabled" : ""}
                                title="Hapus User">
                                <i class="fas fa-trash"></i>
                            </button>
                        </div>
                    </td>
                `;
        tableBody.appendChild(row);
      });
    }

    // Add event listeners
    document
      .querySelectorAll(".edit-user-btn:not(:disabled)")
      .forEach((btn) => {
        btn.addEventListener("click", function () {
          const username = this.getAttribute("data-id");
          editUser(username);
        });
      });

    document
      .querySelectorAll(".delete-user-btn:not(:disabled)")
      .forEach((btn) => {
        btn.addEventListener("click", function () {
          const username = this.getAttribute("data-id");
          deleteUser(username);
        });
      });
  } catch (error) {
    console.error("Error rendering user table:", error);
    Swal.fire({
      icon: "error",
      title: "Gagal Memuat Data",
      text: "Terjadi kesalahan saat memuat data user",
    });
  } finally {
    hideSpinner();
  }
}

// Edit user
function editUser(username) {
  const user = users.find((u) => u.username === username);
  if (!user) return;

  document.getElementById("userModalLabel").textContent = "Edit User";
  document.getElementById("user-id").value = user.username;
  document.getElementById("user-username").value = user.username;
  document.getElementById("user-username").readOnly = true;
  document.getElementById("user-password").value = "";
  document.getElementById("user-password").placeholder =
    "Kosongkan jika tidak ingin mengubah password";
  document.getElementById("user-role").value = user.role;
  document.getElementById("user-region").value = user.region || "";
  document.getElementById("user-unit").value = user.unit || "";
  document.getElementById("user-status").value = user.status;

  const userModal = new bootstrap.Modal(document.getElementById("userModal"));
  userModal.show();
}

// Delete user dengan sync ke Google Sheet
async function deleteUser(username) {
  if (!username) return;

  if (username === "admin") {
    Swal.fire({ icon: "error", title: "Tidak Dapat Menghapus", text: "User admin tidak dapat dihapus!" });
    return;
  }

  const userToDelete = users.find((u) => u.username === username);
  if (!userToDelete) {
    Swal.fire({ icon: "error", title: "Tidak Ditemukan", text: "User tidak ditemukan di daftar." });
    return;
  }

  const confirm = await Swal.fire({
    title: "Hapus User?",
    html: `
      <div class="text-start">
        <p>User yang dihapus <b>benar-benar</b> akan dihapus dari Google Sheet.</p>
        <div class="alert alert-warning">
          <strong>Detail User:</strong><br>
          Username: <strong>${userToDelete.username}</strong><br>
          Role: <strong>${userToDelete.role}</strong><br>
          Region: <strong>${userToDelete.region || "Semua"}</strong><br>
          Unit: <strong>${userToDelete.unit || "Semua"}</strong>
        </div>
      </div>
    `,
    icon: "warning",
    showCancelButton: true,
    confirmButtonColor: "#d33",
    cancelButtonColor: "#3085d6",
    confirmButtonText: "Ya, Hapus!",
    cancelButtonText: "Batal"
  });

  if (!confirm.isConfirmed) return;

  try {
    showProgressModal("Menghapus user...");
    const resp = await fetch(GAS_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "deleteUsers",
        usernames: [username]
      })
    });
    const json = await resp.json();
    hideProgressModal();

    if (json.status !== "success") {
      throw new Error(json.message || "Gagal menghapus user.");
    }

    const { deleted, skipped, details } = json.data || {};
    const info = details && details[0];

    if (deleted >= 1 && info?.ok) {
      // Hapus dari cache lokal
      users = users.filter(u => u.username !== username);
      localStorage.setItem("users", JSON.stringify(users));

      await Swal.fire({ icon: "success", title: "Terhapus!", text: `User ${username} telah dihapus.` });

      // Refresh tampilan dari server agar konsisten
      await renderUserTable();
    } else {
      await Swal.fire({
        icon: "error",
        title: "Gagal Menghapus",
        text: info?.reason ? `Alasan: ${info.reason}` : "Tidak ada baris yang terhapus."
      });
    }
  } catch (err) {
    hideProgressModal();
    console.error(err);
    Swal.fire({ icon: "error", title: "Gagal", text: String(err.message || err) });
  }
}

// Save user dengan sync ke Google Sheet
async function saveUser() {
  const id = document.getElementById("user-id").value;
  const username = document.getElementById("user-username").value.trim();
  const password = document.getElementById("user-password").value;
  const role = document.getElementById("user-role").value;
  const region = document.getElementById("user-region").value;
  const unit = document.getElementById("user-unit").value;
  const status = document.getElementById("user-status").value;

  // Validasi input
  if (!username) {
    Swal.fire({
      icon: "error",
      title: "Data Tidak Lengkap",
      text: "Username harus diisi!",
    });
    return;
  }

  if (username.length < 3) {
    Swal.fire({
      icon: "error",
      title: "Username Tidak Valid",
      text: "Username harus minimal 3 karakter!",
    });
    return;
  }

  if (!id && !password) {
    Swal.fire({
      icon: "error",
      title: "Data Tidak Lengkap",
      text: "Password harus diisi untuk user baru!",
    });
    return;
  }

  showSpinner();

  try {
    let userUpdated = false;

    if (id) {
      // Edit existing user
      const userIndex = users.findIndex((u) => u.username === id);
      if (userIndex !== -1) {
        users[userIndex].role = role;
        users[userIndex].region = region;
        users[userIndex].unit = unit;
        users[userIndex].status = status;

        if (password) {
          users[userIndex].password = password;
        }
        userUpdated = true;
      }
    } else {
      // Add new user
      // Check if username already exists
      if (users.find((u) => u.username === username)) {
        hideSpinner();
        Swal.fire({
          icon: "error",
          title: "Username Sudah Ada",
          text: "Username sudah digunakan!",
        });
        return;
      }

      users.push({
        username,
        password,
        role,
        region,
        unit,
        status,
      });
      userUpdated = true;
    }

    if (userUpdated) {
      // Simpan ke localStorage
      localStorage.setItem("users", JSON.stringify(users));

        // Tentukan patch user yang dikirim
        const patchUser = id
    ? users.find(u => u.username === id)              // edit
    : { username, password, role, region, unit, status }; // user baru

        // Sync ke Google Sheet
      const syncSuccess = await syncUsersToGoogleSheets([patchUser]);

      if (syncSuccess) {
        const userModal = bootstrap.Modal.getInstance(
          document.getElementById("userModal")
        );
        userModal.hide();

        // Refresh table
        await renderUserTable();

        Swal.fire({
          icon: "success",
          title: "Berhasil!",
          text: `User ${
            id ? "diperbarui" : "ditambahkan"
          } dan disinkronisasi ke Google Sheet`,
          timer: 2000,
          showConfirmButton: false,
        });
      } else {
        throw new Error("Gagal menyinkronisasi ke Google Sheet");
      }
    }
  } catch (error) {
    console.error("Error saving user:", error);
    Swal.fire({
      icon: "error",
      title: "Gagal Menyimpan",
      text: "Terjadi kesalahan saat menyimpan user: " + error.message,
    });
  } finally {
    hideSpinner();
  }
}

// Fungsi untuk refresh user table (bisa dipanggil dari luar)
async function refreshUserTable() {
  await renderUserTable();
}

// Event listener untuk refresh button (tambahkan di setupEventListeners)
function setupUserManagementEvents() {
  // Refresh button untuk user table
  const refreshBtn = document.getElementById("refresh-users-btn");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", async () => {
      await refreshUserTable();
    });
  }

  // Auto-load user data saat membuka halaman setting
  document
    .querySelector('.nav-link[data-page="setting"]')
    .addEventListener("click", async () => {
      await refreshUserTable();
    });
}

// Tambahkan fungsi ini di setupEventListeners yang sudah ada
// setupUserManagementEvents();

// Update dashboard
function updateDashboard() {
  const regionFilter = document.getElementById("filter-region").value;
  const unitFilter = document.getElementById("filter-unit").value;
  const jabatanFilter = document.getElementById("filter-jabatan").value;
  const countFilter = parseInt(document.getElementById("filter-count").value);

  // Filter nilai data
  let filteredData = [...nilaiData];

  if (regionFilter) {
    filteredData = filteredData.filter((item) => item.region === regionFilter);
  }

  if (unitFilter) {
    filteredData = filteredData.filter((item) => item.unit === unitFilter);
  }

  if (jabatanFilter) {
    filteredData = filteredData.filter(
      (item) => item.jabatan === jabatanFilter
    );
  }

  // Update statistics
  document.getElementById("total-mandor").textContent = filteredData.length;

  if (filteredData.length > 0) {
    const avgNilai =
      filteredData.reduce((sum, item) => sum + item.totalNilai, 0) /
      filteredData.length;
    document.getElementById("avg-nilai").textContent = avgNilai.toFixed(2);

    const maxNilai = Math.max(...filteredData.map((item) => item.totalNilai));
    document.getElementById("nilai-tertinggi").textContent = maxNilai;

    const minNilai = Math.min(...filteredData.map((item) => item.totalNilai));
    document.getElementById("nilai-terendah").textContent = minNilai;
  } else {
    document.getElementById("avg-nilai").textContent = "0";
    document.getElementById("nilai-tertinggi").textContent = "0";
    document.getElementById("nilai-terendah").textContent = "0";
  }

  // Update top scores table
  updateTopScoresTable(filteredData, countFilter);

  // Update charts
  updateCharts(filteredData);
}

// Update top scores table
function updateTopScoresTable(data, count) {
  const tableBody = document
    .getElementById("top-scores-table")
    .querySelector("tbody");
  tableBody.innerHTML = "";

  // Sort by total nilai (descending) and take top N
  const topData = [...data]
    .sort((a, b) => b.totalNilai - a.totalNilai)
    .slice(0, count);

  topData.forEach((item, index) => {
    const row = document.createElement("tr");
    row.innerHTML = `
                    <td>${index + 1}</td>
                    <td>${item.nip}</td>
                    <td>${item.nama}</td>
                    <td>${item.unit}</td>
                    <td>${item.region}</td>
                    <td>${item.divisi}</td>
                    <td>${item.jabatan}</td>
                    <td>${item.nilaiIsian}</td>
                    <td>${item.nilaiBKM}</td>
                    <td>${item.totalNilai}</td>
                `;
    tableBody.appendChild(row);
  });
}

// Update charts
function updateCharts(data) {
  // Ambil Top 10 berdasarkan totalNilai
  const top10 = [...data]
    .sort((a,b) => b.totalNilai - a.totalNilai)
    .slice(0, 10);

  // ===== Bar: Top 10 Nilai Tertinggi =====
  const topLabels = top10.map(x => `${x.nama} (${x.nip})`);
  const topValues = top10.map(x => x.totalNilai);

  const ctxTop = document.getElementById('topScoreCanvas').getContext('2d');
  destroyChartIfAny(_chartTop10);
  _chartTop10 = new Chart(ctxTop, {
    type: 'bar',
    data: {
      labels: topLabels,
      datasets: [{ label: 'Total Nilai', data: topValues }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => ` ${c.parsed.y}` } }
      },
      scales: {
        x: { ticks: { maxRotation: 60, minRotation: 0, autoSkip: false } },
        y: { beginAtZero: true, suggestedMax: 100 }
      }
    }
  });

  // ===== Pie/Doughnut: Distribusi Nilai (bucket) =====
  // Kelompokkan Total Nilai ke bucket: <50, 50-59, 60-69, 70-79, 80-89, >=90
  const buckets = [
    { label: '< 50',    min: -Infinity, max: 49 },
    { label: '50–59',   min: 50, max: 59 },
    { label: '60–69',   min: 60, max: 69 },
    { label: '70–79',   min: 70, max: 79 },
    { label: '80–89',   min: 80, max: 89 },
    { label: '≥ 90',    min: 90, max: Infinity },
  ];
  const bucketCounts = buckets.map(b =>
    data.filter(x => x.totalNilai >= b.min && x.totalNilai <= b.max).length
  );

  const ctxPie = document.getElementById('scoreDistributionCanvas').getContext('2d');
  destroyChartIfAny(_chartDistribusi);
  _chartDistribusi = new Chart(ctxPie, {
    type: 'doughnut',
    data: {
      labels: buckets.map(b => b.label),
      datasets: [{ data: bucketCounts }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom' }
      },
      cutout: '55%'
    }
  });
}

function focusNipField() {
  const nip = document.getElementById('input-nip');
  if (!nip) return;

  // Bersihkan dulu agar mobile keyboard mau muncul lagi
  nip.blur();

  // Coba fokus segera, lalu ulang via rAF bila belum aktif
  const tryFocus = () => {
    nip.focus({ preventScroll: false });
    if (document.activeElement !== nip) {
      requestAnimationFrame(() => {
        nip.focus({ preventScroll: false });
      });
    }
  };

  // Sedikit jeda supaya layout settle
  setTimeout(tryFocus, 0);
}

// Show spinner
function showSpinner() {
  document.getElementById("spinner-overlay").classList.remove("d-none");
}

// Hide spinner
function hideSpinner() {
  document.getElementById("spinner-overlay").classList.add("d-none");
}

// Update progress
function updateProgress(percent) {
  document.querySelector(".progress-bar").style.width = `${percent}%`;
}

// Simpan instance modal secara global
let _progressModalInstance = null;

function showProgressModal(message) {
  const el = document.getElementById('progressModal');
  document.getElementById('progress-message').textContent = message || 'Sedang memproses...';
  const bar = el.querySelector('.progress-bar');
  if (bar) bar.style.width = '0%';

  // Pakai getOrCreateInstance agar selalu ada instance yang valid
  _progressModalInstance = bootstrap.Modal.getOrCreateInstance(el, {
    backdrop: 'static',
    keyboard: false
  });
  _progressModalInstance.show();
}

function hideProgressModal() {
  const el = document.getElementById('progressModal');
  // Ambil instance yang ada, atau buat kalau hilang (defensif)
  const instance = _progressModalInstance || bootstrap.Modal.getInstance(el) || bootstrap.Modal.getOrCreateInstance(el);
  try {
    instance.hide();
  } catch(e) {
    // noop
  } finally {
    // Tunggu event hidden agar cleanup backdrop/scroll tepat waktu
    const cleanup = () => {
      setTimeout(() => {
        document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
        document.body.classList.remove('modal-open');
        document.body.style.removeProperty('overflow');
        document.body.style.removeProperty('padding-right');
        _progressModalInstance = null;
      }, 50);
      el.removeEventListener('hidden.bs.modal', cleanup);
    };
    el.addEventListener('hidden.bs.modal', cleanup);

    // Fallback kalau event tidak terpanggil (misal kondisi edge)
    setTimeout(() => {
      document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
      document.body.classList.remove('modal-open');
      document.body.style.removeProperty('overflow');
      document.body.style.removeProperty('padding-right');
      _progressModalInstance = null;
    }, 300);
  }
}