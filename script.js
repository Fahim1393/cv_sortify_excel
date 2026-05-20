// Initialize Lucide icons
lucide.createIcons();

// DOM Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const historyCard = document.getElementById('history-card');
const historyList = document.getElementById('history-list');
const clearHistoryBtn = document.getElementById('clear-history-btn');
const uploadContent = document.getElementById('upload-content');
const fileInfo = document.getElementById('file-info');
const fileName = document.getElementById('file-name');
const fileSize = document.getElementById('file-size');
const clearBtn = document.getElementById('clear-btn');
const errorBanner = document.getElementById('error-banner');
const errorMessage = document.getElementById('error-message');
const processingOverlay = document.getElementById('processing-overlay');

const uploadSection = document.getElementById('upload-section');
const resultsSection = document.getElementById('results-section');
const backBtn = document.getElementById('back-btn');
const tableBody = document.getElementById('table-body');
const recordCount = document.getElementById('record-count');
const generateReportBtn = document.getElementById('generate-report-btn');

let currentFile = null;
let currentData = [];

// Drag and Drop Events
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, () => dropZone.classList.add('dragging'), false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragging'), false);
});

dropZone.addEventListener('drop', handleDrop, false);
dropZone.addEventListener('click', () => {
    if (!currentFile) {
        fileInput.click();
    }
});

fileInput.addEventListener('change', handleFileSelect, false);
clearBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    clearFile();
});
backBtn.addEventListener('click', resetView);
if (generateReportBtn) {
    generateReportBtn.addEventListener('click', generateReport);
}

// Functions
function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length > 0) {
        validateAndProcessFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        validateAndProcessFile(files[0]);
    }
}

function validateAndProcessFile(file) {
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    
    if (validTypes.includes(file.type) || file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        currentFile = file;
        showFileUI(file);
        hideError();
        processExcel(file);
    } else {
        showError('Please upload a valid Excel file (.xlsx or .xls)');
        clearFile();
    }
}

function showFileUI(file) {
    uploadContent.classList.add('hidden');
    fileInfo.classList.remove('hidden');
    dropZone.classList.add('has-file');
    fileName.textContent = file.name;
    fileSize.textContent = (file.size / 1024).toFixed(2) + ' KB';
}

function clearFile() {
    currentFile = null;
    fileInput.value = '';
    uploadContent.classList.remove('hidden');
    fileInfo.classList.add('hidden');
    dropZone.classList.remove('has-file');
}

function showError(msg) {
    errorBanner.classList.remove('hidden');
    errorMessage.textContent = msg;
}

function hideError() {
    errorBanner.classList.add('hidden');
}

function resetView() {
    clearFile();
    hideError();
    resultsSection.classList.add('hidden');
    uploadSection.classList.remove('hidden');
}

// Excel Processing Logic
function processExcel(file) {
    processingOverlay.classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
            
            // Format and sort data
            const formattedData = formatAndSortData(json);
            
            // Artificial delay for better UX (to show the processing state)
            setTimeout(() => {
                displayResults(formattedData);
                saveToHistory(file.name, formattedData);
                processingOverlay.classList.add('hidden');
            }, 800);
            
        } catch (error) {
            console.error("Error processing Excel:", error);
            showError("Failed to process the Excel file. Ensure it contains the correct data.");
            processingOverlay.classList.add('hidden');
            clearFile();
        }
    };
    
    reader.onerror = function() {
        showError("Error reading file.");
        processingOverlay.classList.add('hidden');
        clearFile();
    };
    
    reader.readAsArrayBuffer(file);
}

function formatAndSortData(json) {
    const formatted = json.map(row => {
        const newRow = {};
        for (const key in row) {
            const lowerKey = key.toLowerCase().trim();
            if (lowerKey === 'name') newRow.name = row[key];
            else if (lowerKey === 'university') newRow.university = row[key];
            else if (lowerKey === 'department') newRow.department = row[key];
            else if (lowerKey === 'email') newRow.email = row[key];
            else if (lowerKey === 'contact number' || lowerKey === 'contact' || lowerKey === 'phone' || lowerKey === 'phone number' || lowerKey === 'mobile') newRow.contactNumber = row[key];
            else if (lowerKey.includes('explain your role')) newRow.eca = row[key];
            else if (lowerKey === 'cv link' || lowerKey === 'cv' || lowerKey === 'resume' || lowerKey === 'resume link' || lowerKey === 'link' || lowerKey === 'upload your cv') newRow.cvLink = row[key];
        }
        return {
            name: newRow.name || 'N/A',
            university: newRow.university || 'N/A',
            department: newRow.department || 'N/A',
            email: newRow.email || 'N/A',
            contactNumber: newRow.contactNumber || 'N/A',
            eca: newRow.eca || 'N/A',
            cvLink: newRow.cvLink || 'N/A'
        };
    });

    // Sort by university alphabetically
    return formatted.sort((a, b) => {
        return a.university.toLowerCase().localeCompare(b.university.toLowerCase());
    });
}

function displayResults(data) {
    if (data.length === 0) {
        showError("No valid data found in the file.");
        return;
    }

    recordCount.textContent = `${data.length} Records`;
    tableBody.innerHTML = '';
    currentData = data;

    data.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.style.setProperty('--row-index', index);
        
        tr.innerHTML = `
            <td>${escapeHTML(row.name)}</td>
            <td><span class="university-badge">${escapeHTML(row.university)}</span></td>
            <td>${escapeHTML(row.department)}</td>
            <td><a href="mailto:${escapeHTML(row.email)}" class="email-link">${escapeHTML(row.email)}</a></td>
            <td>${escapeHTML(row.contactNumber)}</td>
            <td>${escapeHTML(row.eca)}</td>
            <td>${row.cvLink !== 'N/A' ? `<a href="${escapeHTML(row.cvLink)}" target="_blank" class="cv-link">View CV <i data-lucide="external-link"></i></a>` : 'N/A'}</td>
            <td>
                <select class="review-select" data-index="${index}">
                    <option value="Pending">Pending</option>
                    <option value="Yes">Yes</option>
                    <option value="No">No</option>
                </select>
            </td>
        `;
        tableBody.appendChild(tr);
    });

    uploadSection.classList.add('hidden');
    resultsSection.classList.remove('hidden');
    lucide.createIcons(); // Re-init icons if any dynamic ones were added (not necessary here but good practice)
}

function escapeHTML(str) {
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

function generateReport() {
    if (currentData.length === 0) return;

    // Collect updated review statuses from the table
    const rows = tableBody.querySelectorAll('tr');
    const exportData = currentData.map((row, index) => {
        const select = rows[index].querySelector('.review-select');
        return {
            'Name': row.name,
            'University': row.university,
            'Department': row.department,
            'Email': row.email,
            'Contact Number': row.contactNumber,
            'ECA': row.eca,
            'CV Link': row.cvLink,
            'Review Status': select ? select.value : 'Pending'
        };
    });

    // Create workbook and worksheet
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sorted Candidates");

    // Generate filename with date
    const date = new Date().toISOString().split('T')[0];
    const filename = `CV_Report_${date}.xlsx`;

    // Download file
    XLSX.writeFile(workbook, filename);
}

// History Functions
function getHistory() {
    try {
        const historyStr = localStorage.getItem('sortify_history');
        return historyStr ? JSON.parse(historyStr) : [];
    } catch (e) {
        return [];
    }
}

function saveToHistory(filename, data) {
    let history = getHistory();
    const newItem = {
        id: Date.now(),
        filename: filename,
        date: new Date().toISOString(),
        recordCount: data.length,
        data: data
    };
    
    // Check if filename already exists, remove it to avoid exact duplicates
    history = history.filter(item => item.filename !== filename || item.recordCount !== data.length);
    
    history.unshift(newItem);
    
    if (history.length > 10) {
        history = history.slice(0, 10);
    }
    
    trySaveHistory(history);
}

function trySaveHistory(history) {
    try {
        localStorage.setItem('sortify_history', JSON.stringify(history));
        renderHistory();
    } catch (e) {
        if (history.length > 1) {
            history.pop(); // Remove oldest
            trySaveHistory(history); // Recursively try again
        } else {
            console.error('File too large to save in history.');
        }
    }
}

function renderHistory() {
    if (!historyCard || !historyList) return;
    
    const history = getHistory();
    
    if (history.length === 0) {
        historyCard.classList.add('hidden');
        return;
    }
    
    historyCard.classList.remove('hidden');
    historyList.innerHTML = '';
    
    history.forEach(item => {
        const dateObj = new Date(item.date);
        const dateStr = dateObj.toLocaleDateString(undefined, { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
        
        const div = document.createElement('div');
        div.className = 'history-item';
        div.innerHTML = `
            <div class="history-item-info">
                <span class="history-item-name"><i data-lucide="file-spreadsheet" style="width: 16px; height: 16px; color: var(--success);"></i> ${escapeHTML(item.filename)}</span>
                <span class="history-item-meta">${dateStr} • ${item.recordCount} Records</span>
            </div>
            <div class="history-item-action">
                View <i data-lucide="chevron-right"></i>
            </div>
        `;
        div.addEventListener('click', () => {
            displayResults(item.data);
            hideError();
        });
        historyList.appendChild(div);
    });
    
    lucide.createIcons();
}

function clearHistory() {
    localStorage.removeItem('sortify_history');
    renderHistory();
}

if (clearHistoryBtn) {
    clearHistoryBtn.addEventListener('click', clearHistory);
}

// Render history on load
document.addEventListener('DOMContentLoaded', renderHistory);
