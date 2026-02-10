(() => {
  // --- DOM refs: Sections ---
  const homeSection = document.getElementById('home-section');
  const uploadSection = document.getElementById('upload-section');
  const optinPromptSection = document.getElementById('optin-prompt-section');
  const filterSection = document.getElementById('filter-section');
  const resultsSection = document.getElementById('results-section');
  const optinManagerSection = document.getElementById('optin-manager-section');

  const allSections = [homeSection, uploadSection, optinPromptSection, filterSection, resultsSection, optinManagerSection];

  // --- DOM refs: Home ---
  const homeGenerate = document.getElementById('home-generate');
  const homeOptin = document.getElementById('home-optin');

  // --- DOM refs: Upload ---
  const uploadZone = document.getElementById('upload-zone');
  const fileInput = document.getElementById('file-input');
  const fileInfo = document.getElementById('file-info');
  const fileName = document.getElementById('file-name');
  const clearFileBtn = document.getElementById('clear-file');
  const btnUploadHome = document.getElementById('btn-upload-home');

  // --- DOM refs: Opt-in Prompt ---
  const optinUploadZone = document.getElementById('optin-upload-zone');
  const optinFileInput = document.getElementById('optin-file-input');
  const optinFileInfo = document.getElementById('optin-file-info');
  const optinFileName = document.getElementById('optin-file-name');
  const optinClearFile = document.getElementById('optin-clear-file');
  const btnOptinBack = document.getElementById('btn-optin-back');
  const btnOptinSkip = document.getElementById('btn-optin-skip');
  const btnOptinContinue = document.getElementById('btn-optin-continue');

  // --- DOM refs: Prefix Filter ---
  const prefixList = document.getElementById('prefix-list');
  const btnBack = document.getElementById('btn-back');
  const btnGenerate = document.getElementById('btn-generate');

  // --- DOM refs: Results ---
  const resultsMeta = document.getElementById('results-meta');
  const resultsBody = document.getElementById('results-body');
  const btnExport = document.getElementById('btn-export');
  const btnNewReport = document.getElementById('btn-new-report');

  // --- DOM refs: Opt-in Manager ---
  const optinMgrUploadZone = document.getElementById('optin-mgr-upload-zone');
  const optinMgrFileInput = document.getElementById('optin-mgr-file-input');
  const addNameInput = document.getElementById('add-name-input');
  const addCompanyInput = document.getElementById('add-company-input');
  const btnAddEmployee = document.getElementById('btn-add-employee');
  const optinBody = document.getElementById('optin-body');
  const optinTableWrapper = document.getElementById('optin-table-wrapper');
  const optinEmpty = document.getElementById('optin-empty');
  const btnOptinHome = document.getElementById('btn-optin-home');
  const btnOptinExport = document.getElementById('btn-optin-export');

  // --- DOM refs: Opt-in Manager extra ---
  const btnOptinClear = document.getElementById('btn-optin-clear');

  // --- DOM refs: Modal & Loading ---
  const modalOverlay = document.getElementById('modal-overlay');
  const modalTitle = document.getElementById('modal-title');
  const modalMessage = document.getElementById('modal-message');
  const modalOk = document.getElementById('modal-ok');
  const confirmOverlay = document.getElementById('confirm-overlay');
  const confirmTitle = document.getElementById('confirm-title');
  const confirmMessage = document.getElementById('confirm-message');
  const confirmCancel = document.getElementById('confirm-cancel');
  const confirmYes = document.getElementById('confirm-yes');
  const loadingOverlay = document.getElementById('loading-overlay');

  // ============================================================
  //  STATE
  // ============================================================
  // Report flow
  let parsedRows = [];           // { name, prefix, date, isGeneric }
  let prefixMap = {};            // prefix -> Set of names
  let selectedPrefixes = new Set();
  let includeGenericCards = false;
  let genericCardNames = new Set();
  let reportData = [];           // { name, company, days }
  let reportMonth = '';

  // Opt-in reference (used during report filtering)
  let optinRef = null;           // null = not loaded, Map<normalizedName, { name, company }>

  // Opt-in manager list
  let optinManagerList = [];     // [{ name, company }]

  // ============================================================
  //  UTILITIES
  // ============================================================
  function showModal(title, message) {
    modalTitle.textContent = title;
    modalMessage.textContent = message;
    modalOverlay.classList.remove('hidden');
    modalOk.focus();
  }

  modalOk.addEventListener('click', () => modalOverlay.classList.add('hidden'));
  modalOverlay.addEventListener('click', (e) => {
    if (e.target === modalOverlay) modalOverlay.classList.add('hidden');
  });

  let confirmCallback = null;

  function showConfirm(title, message, onConfirm) {
    confirmTitle.textContent = title;
    confirmMessage.textContent = message;
    confirmCallback = onConfirm;
    confirmOverlay.classList.remove('hidden');
    confirmCancel.focus();
  }

  confirmCancel.addEventListener('click', () => confirmOverlay.classList.add('hidden'));
  confirmOverlay.addEventListener('click', (e) => {
    if (e.target === confirmOverlay) confirmOverlay.classList.add('hidden');
  });
  confirmYes.addEventListener('click', () => {
    confirmOverlay.classList.add('hidden');
    if (confirmCallback) confirmCallback();
  });

  function showSection(section) {
    allSections.forEach(s => s.classList.add('hidden'));
    section.classList.remove('hidden');
  }

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function toLastFirst(name) {
    const parts = name.trim().split(/\s+/);
    if (parts.length < 2) return name;
    const last = parts.pop();
    return `${last}, ${parts.join(' ')}`;
  }

  function normalizeName(name) {
    return name.trim().toLowerCase().replace(/\s+/g, ' ');
  }

  // Generic card names that get grouped under "Guest & Spare Cards"
  const GENERIC_NAMES = new Set(['guest', 'spare', 'master', 'test', 'unknown']);

  function setupDropZone(zone, input, onFile) {
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      zone.classList.remove('dragover');
      if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]);
    });
    input.addEventListener('change', () => { if (input.files[0]) onFile(input.files[0]); });
  }

  function readWorkbook(file, callback) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext !== 'xls' && ext !== 'xlsx') {
      showModal('Invalid File', 'Please upload a .xls or .xlsx file.');
      return;
    }
    loadingOverlay.classList.remove('hidden');
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        loadingOverlay.classList.add('hidden');
        callback(rows, file.name);
      } catch (err) {
        loadingOverlay.classList.add('hidden');
        showModal('Parse Error', 'Failed to read the file. Make sure it is a valid spreadsheet.');
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ============================================================
  //  HOME
  // ============================================================
  homeGenerate.addEventListener('click', () => {
    resetReportState();
    showSection(uploadSection);
  });

  homeOptin.addEventListener('click', () => {
    showSection(optinManagerSection);
    renderOptinManager();
  });

  // ============================================================
  //  REPORT: Upload Access Log
  // ============================================================
  function resetReportState() {
    parsedRows = [];
    prefixMap = {};
    selectedPrefixes = new Set();
    includeGenericCards = false;
    genericCardNames = new Set();
    reportData = [];
    reportMonth = '';
    optinRef = null;
    fileInput.value = '';
    fileInfo.classList.add('hidden');
    uploadZone.classList.remove('hidden');
    optinFileInput.value = '';
    optinFileInfo.classList.add('hidden');
    optinUploadZone.classList.remove('hidden');
    btnOptinContinue.classList.add('hidden');
  }

  setupDropZone(uploadZone, fileInput, handleAccessLogFile);

  clearFileBtn.addEventListener('click', () => {
    resetReportState();
  });

  btnUploadHome.addEventListener('click', () => showSection(homeSection));

  function handleAccessLogFile(file) {
    readWorkbook(file, (rows, name) => {
      fileName.textContent = name;
      fileInfo.classList.remove('hidden');
      uploadZone.classList.add('hidden');

      parseAccessLog(rows);

      if (parsedRows.length === 0) {
        showModal('No Data Found', 'Could not find any access events in this file.');
        return;
      }

      // Move to opt-in prompt step
      showSection(optinPromptSection);
    });
  }

  // ============================================================
  //  REPORT: Parse Access Log
  // ============================================================
  function parseAccessLog(rows) {
    parsedRows = [];
    prefixMap = {};
    genericCardNames = new Set();

    let headerIdx = -1;
    let dateCol = -1;
    let descCol = -1;

    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const row = rows[i];
      for (let j = 0; j < row.length; j++) {
        const cell = String(row[j]).trim();
        if (cell === 'Date and Time' || cell === 'Date and time') dateCol = j;
        if (cell === 'Description #2') descCol = j;
      }
      if (dateCol >= 0 && descCol >= 0) { headerIdx = i; break; }
      dateCol = -1;
      descCol = -1;
    }

    if (headerIdx === -1) { headerIdx = 5; dateCol = 1; descCol = 7; }

    let detectedMonth = null;

    for (let i = headerIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length <= Math.max(dateCol, descCol)) continue;

      const rawDate = String(row[dateCol]).trim();
      const rawName = String(row[descCol]).trim();
      if (!rawDate || !rawName) continue;

      const date = parseDate(rawDate);
      if (!date) continue;

      if (!detectedMonth) {
        const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
        reportMonth = `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
        detectedMonth = true;
      }

      const dateStr = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
      const { prefix, name, isGeneric } = parseName(rawName);
      if (!name) continue;

      parsedRows.push({ name, prefix, date: dateStr, isGeneric });

      if (isGeneric) {
        genericCardNames.add(name);
      } else {
        if (!prefixMap[prefix]) prefixMap[prefix] = new Set();
        prefixMap[prefix].add(name);
      }
    }
  }

  function parseDate(raw) {
    if (typeof raw === 'number' || /^\d+(\.\d+)?$/.test(raw)) {
      const serial = typeof raw === 'number' ? raw : parseFloat(raw);
      const utcDays = Math.floor(serial) - 25569;
      return new Date(utcDays * 86400 * 1000);
    }
    const d = new Date(raw);
    return isNaN(d.getTime()) ? null : d;
  }

  function parseName(raw) {
    if (!raw || raw === '0') return { prefix: '', name: '', isGeneric: false };
    const dashIdx = raw.indexOf(' - ');
    if (dashIdx === -1) return { prefix: '', name: '', isGeneric: false };

    const prefix = raw.substring(0, dashIdx).trim();
    const remainder = raw.substring(dashIdx + 3).trim();
    const parenMatch = remainder.match(/\(([^)]+)\)/);
    const name = parenMatch ? parenMatch[1].trim() : remainder;
    const isGeneric = GENERIC_NAMES.has(name.toLowerCase());
    return { prefix, name, isGeneric };
  }

  // ============================================================
  //  REPORT: Opt-in Prompt
  // ============================================================
  setupDropZone(optinUploadZone, optinFileInput, handleOptinRefFile);

  optinClearFile.addEventListener('click', () => {
    optinRef = null;
    optinFileInput.value = '';
    optinFileInfo.classList.add('hidden');
    optinUploadZone.classList.remove('hidden');
    btnOptinContinue.classList.add('hidden');
  });

  btnOptinBack.addEventListener('click', () => showSection(uploadSection));
  btnOptinSkip.addEventListener('click', () => {
    optinRef = null;
    showSection(filterSection);
    buildPrefixFilter();
  });
  btnOptinContinue.addEventListener('click', () => {
    showSection(filterSection);
    buildPrefixFilter();
  });

  function handleOptinRefFile(file) {
    readWorkbook(file, (rows, name) => {
      optinRef = new Map();

      // Find header row with "Employee Name" or "Name"
      let nameCol = -1;
      let companyCol = -1;
      let startRow = 0;

      for (let i = 0; i < Math.min(rows.length, 10); i++) {
        for (let j = 0; j < (rows[i] || []).length; j++) {
          const cell = String(rows[i][j]).trim().toLowerCase();
          if (cell === 'employee name' || cell === 'name') nameCol = j;
          if (cell === 'company') companyCol = j;
        }
        if (nameCol >= 0) { startRow = i + 1; break; }
      }

      // Fallback: assume col 0 = name, col 1 = company
      if (nameCol === -1) { nameCol = 0; companyCol = 1; startRow = 1; }

      for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        const empName = String(row[nameCol] || '').trim();
        if (!empName) continue;
        const company = companyCol >= 0 ? String(row[companyCol] || '').trim() : '';
        optinRef.set(normalizeName(empName), { name: empName, company });
      }

      if (optinRef.size === 0) {
        showModal('No Names Found', 'Could not find any employee names in the uploaded file.');
        optinRef = null;
        return;
      }

      optinFileName.textContent = `${name} (${optinRef.size} employees)`;
      optinFileInfo.classList.remove('hidden');
      optinUploadZone.classList.add('hidden');
      btnOptinContinue.classList.remove('hidden');
    });
  }

  // ============================================================
  //  REPORT: Prefix Filter
  // ============================================================
  function buildPrefixFilter() {
    prefixList.innerHTML = '';
    selectedPrefixes = new Set();
    includeGenericCards = false;

    const checkSvg = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>`;

    Object.keys(prefixMap).sort().forEach((prefix) => {
      const count = prefixMap[prefix].size;
      selectedPrefixes.add(prefix);

      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'prefix-chip selected';
      chip.innerHTML = `<span class="prefix-chip__check">${checkSvg}</span><span class="prefix-chip__label">${prefix}</span><span class="prefix-chip__count">(${count})</span>`;

      chip.addEventListener('click', () => {
        if (selectedPrefixes.has(prefix)) { selectedPrefixes.delete(prefix); chip.classList.remove('selected'); }
        else { selectedPrefixes.add(prefix); chip.classList.add('selected'); }
        btnGenerate.disabled = selectedPrefixes.size === 0 && !includeGenericCards;
      });
      prefixList.appendChild(chip);
    });

    if (genericCardNames.size > 0) {
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'prefix-chip';
      chip.innerHTML = `<span class="prefix-chip__check">${checkSvg}</span><span class="prefix-chip__label">Guest & Spare Cards</span><span class="prefix-chip__count">(${genericCardNames.size})</span>`;
      chip.addEventListener('click', () => {
        includeGenericCards = !includeGenericCards;
        chip.classList.toggle('selected', includeGenericCards);
        btnGenerate.disabled = selectedPrefixes.size === 0 && !includeGenericCards;
      });
      prefixList.appendChild(chip);
    }
  }

  btnBack.addEventListener('click', () => showSection(optinPromptSection));
  btnGenerate.addEventListener('click', generateReport);
  btnNewReport.addEventListener('click', () => { resetReportState(); showSection(homeSection); });

  // ============================================================
  //  REPORT: Generate
  // ============================================================
  function generateReport() {
    loadingOverlay.classList.remove('hidden');

    setTimeout(() => {
      const filtered = parsedRows.filter(r => {
        if (r.isGeneric) return includeGenericCards;
        return selectedPrefixes.has(r.prefix);
      });

      // Deduplicate: one swipe per employee per day
      const dayMap = {};
      filtered.forEach(({ name, date }) => {
        if (!dayMap[name]) dayMap[name] = new Set();
        dayMap[name].add(date);
      });

      // Build report, optionally filtered by opt-in ref
      reportData = Object.entries(dayMap)
        .map(([name, dates]) => {
          let company = '';

          if (optinRef) {
            const match = optinRef.get(normalizeName(name));
            if (!match) return null; // not on opt-in list — exclude
            company = match.company || '';
          }

          return { name, company, days: dates.size };
        })
        .filter(Boolean)
        .sort((a, b) => {
          const aLast = a.name.split(' ').pop().toLowerCase();
          const bLast = b.name.split(' ').pop().toLowerCase();
          if (aLast !== bLast) return aLast.localeCompare(bLast);
          return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
        });

      renderResults();
      loadingOverlay.classList.add('hidden');
      showSection(resultsSection);
    }, 50);
  }

  function renderResults() {
    const totalEmployees = reportData.length;
    const totalDays = reportData.reduce((sum, r) => sum + r.days, 0);
    const groups = [...selectedPrefixes].sort().join(', ');
    const optinNote = optinRef ? ` — Filtered by opt-in reference (${optinRef.size})` : '';

    resultsMeta.textContent = `${reportMonth} — ${totalEmployees} employees, ${totalDays} total in-office days — Groups: ${groups}${optinNote}`;

    resultsBody.innerHTML = '';
    reportData.forEach((row, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="col-num">${idx + 1}</td>
        <td>${escapeHtml(row.name)}</td>
        <td>${escapeHtml(toLastFirst(row.name))}</td>
        <td>${escapeHtml(row.company)}</td>
        <td class="col-days">${row.days}</td>
      `;
      resultsBody.appendChild(tr);
    });
  }

  // ============================================================
  //  REPORT: Export
  // ============================================================
  btnExport.addEventListener('click', () => {
    if (reportData.length === 0) return;

    const wsData = [['Employee Name', 'Last, First', 'Company', 'Days in Office']];
    reportData.forEach((row) => {
      wsData.push([row.name, toLastFirst(row.name), row.company, row.days]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [{ wch: 28 }, { wch: 28 }, { wch: 24 }, { wch: 15 }];

    const wb = XLSX.utils.book_new();
    const safeMonth = reportMonth.replace(/\s+/g, '-');
    XLSX.utils.book_append_sheet(wb, ws, 'Report');
    XLSX.writeFile(wb, `Commuter-Report-${safeMonth}.xlsx`);
  });

  // ============================================================
  //  OPT-IN REFERENCE MANAGER
  // ============================================================
  setupDropZone(optinMgrUploadZone, optinMgrFileInput, handleOptinMgrImport);

  btnOptinHome.addEventListener('click', () => showSection(homeSection));

  btnOptinClear.addEventListener('click', () => {
    if (optinManagerList.length === 0) {
      showModal('Nothing to Clear', 'The list is already empty.');
      return;
    }
    showConfirm(
      'Clear All Data',
      'Are you sure you want to remove all employees from the opt-in reference list? This cannot be undone.',
      () => {
        optinManagerList = [];
        renderOptinManager();
      }
    );
  });

  function handleOptinMgrImport(file) {
    readWorkbook(file, (rows) => {
      // Detect if this is a raw access log (has "Description #2" header)
      let isAccessLog = false;
      let descCol = -1;
      let headerIdx = -1;

      for (let i = 0; i < Math.min(rows.length, 20); i++) {
        for (let j = 0; j < (rows[i] || []).length; j++) {
          const cell = String(rows[i][j]).trim();
          if (cell === 'Description #2') { descCol = j; headerIdx = i; isAccessLog = true; break; }
        }
        if (isAccessLog) break;
      }

      const existing = new Set(optinManagerList.map(e => normalizeName(e.name)));
      let added = 0;

      if (isAccessLog) {
        // Parse names from Description #2 using same logic as report flow
        const uniqueNames = new Set();
        for (let i = headerIdx + 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length <= descCol) continue;
          const rawName = String(row[descCol] || '').trim();
          if (!rawName) continue;

          const { name, isGeneric } = parseName(rawName);
          if (!name || isGeneric) continue;

          const normalized = normalizeName(name);
          if (uniqueNames.has(normalized) || existing.has(normalized)) continue;
          uniqueNames.add(normalized);

          optinManagerList.push({ name, company: '' });
          existing.add(normalized);
          added++;
        }
      } else {
        // Treat as a formatted opt-in reference (Employee Name + Company columns)
        let nameCol = -1;
        let companyCol = -1;
        let startRow = 0;

        for (let i = 0; i < Math.min(rows.length, 10); i++) {
          for (let j = 0; j < (rows[i] || []).length; j++) {
            const cell = String(rows[i][j]).trim().toLowerCase();
            if (cell === 'employee name' || cell === 'name') nameCol = j;
            if (cell === 'company') companyCol = j;
          }
          if (nameCol >= 0) { startRow = i + 1; break; }
        }

        if (nameCol === -1) { nameCol = 0; companyCol = 1; startRow = 1; }

        for (let i = startRow; i < rows.length; i++) {
          const row = rows[i];
          if (!row) continue;
          const empName = String(row[nameCol] || '').trim();
          if (!empName) continue;
          const normalized = normalizeName(empName);
          if (existing.has(normalized)) continue;

          const company = companyCol >= 0 ? String(row[companyCol] || '').trim() : '';
          optinManagerList.push({ name: empName, company });
          existing.add(normalized);
          added++;
        }
      }

      sortOptinManagerList();
      renderOptinManager();

      if (added === 0) {
        showModal('No New Names', 'All names in the file already exist in the list.');
      }
    });
  }

  // Add employee manually
  btnAddEmployee.addEventListener('click', addEmployeeManually);
  addNameInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') addEmployeeManually(); });
  addCompanyInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') addEmployeeManually(); });

  function addEmployeeManually() {
    const name = addNameInput.value.trim();
    const company = addCompanyInput.value.trim();
    if (!name) { showModal('Missing Name', 'Please enter an employee name.'); return; }

    const existing = optinManagerList.some(e => normalizeName(e.name) === normalizeName(name));
    if (existing) { showModal('Duplicate', 'This employee is already in the list.'); return; }

    optinManagerList.push({ name, company });
    sortOptinManagerList();
    renderOptinManager();

    addNameInput.value = '';
    addCompanyInput.value = '';
    addNameInput.focus();
  }

  function sortOptinManagerList() {
    optinManagerList.sort((a, b) => {
      const aLast = a.name.split(' ').pop().toLowerCase();
      const bLast = b.name.split(' ').pop().toLowerCase();
      if (aLast !== bLast) return aLast.localeCompare(bLast);
      return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
    });
  }

  function renderOptinManager() {
    if (optinManagerList.length === 0) {
      optinTableWrapper.classList.add('hidden');
      optinEmpty.classList.remove('hidden');
      return;
    }

    optinTableWrapper.classList.remove('hidden');
    optinEmpty.classList.add('hidden');
    optinBody.innerHTML = '';

    optinManagerList.forEach((entry, idx) => {
      const tr = document.createElement('tr');

      // Number cell
      const tdNum = document.createElement('td');
      tdNum.className = 'col-num';
      tdNum.textContent = idx + 1;

      // Name cell
      const tdName = document.createElement('td');
      tdName.textContent = entry.name;

      // Company cell with inline input
      const tdCompany = document.createElement('td');
      const companyInput = document.createElement('input');
      companyInput.type = 'text';
      companyInput.className = 'inline-input';
      companyInput.value = entry.company;
      companyInput.placeholder = 'Assign company...';
      companyInput.addEventListener('change', () => {
        entry.company = companyInput.value.trim();
      });
      tdCompany.appendChild(companyInput);

      // Remove button cell
      const tdAction = document.createElement('td');
      tdAction.className = 'col-action';
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'btn btn--danger btn--sm';
      removeBtn.textContent = 'Remove';
      removeBtn.addEventListener('click', () => {
        optinManagerList.splice(idx, 1);
        renderOptinManager();
      });
      tdAction.appendChild(removeBtn);

      tr.append(tdNum, tdName, tdCompany, tdAction);
      optinBody.appendChild(tr);
    });
  }

  // Export opt-in reference
  btnOptinExport.addEventListener('click', () => {
    if (optinManagerList.length === 0) {
      showModal('No Data', 'Add at least one employee before exporting.');
      return;
    }

    const wsData = [['Employee Name', 'Company']];
    optinManagerList.forEach((e) => wsData.push([e.name, e.company]));

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [{ wch: 30 }, { wch: 24 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Opt-in Reference');
    XLSX.writeFile(wb, 'Opt-in-Reference.xlsx');
  });
})();
