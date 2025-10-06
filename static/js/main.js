// Drag and Drop functionality
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const processBtn = document.getElementById('processBtn');

// Click to browse
dropZone.addEventListener('click', () => {
    fileInput.click();
});

// Prevent default drag behaviors
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, preventDefaults, false);
    document.body.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

// Highlight drop zone when item is dragged over it
['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, highlight, false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, unhighlight, false);
});

function highlight(e) {
    dropZone.classList.add('drag-over');
}

function unhighlight(e) {
    dropZone.classList.remove('drag-over');
}

// Handle dropped files
dropZone.addEventListener('drop', handleDrop, false);

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    
    if (files.length > 0) {
        fileInput.files = files;
        handleFiles(files);
    }
}

// Handle file selection
fileInput.addEventListener('change', function(e) {
    handleFiles(this.files);
});

function handleFiles(files) {
    if (files.length > 0) {
        const file = files[0];
        const fileName = file.name;
        const fileSize = (file.size / 1024).toFixed(2); // KB
        
        // Check file type
        const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
        if (!validTypes.includes(file.type) && !fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
            fileInfo.innerHTML = '<i class="fas fa-exclamation-triangle"></i> Invalid file type. Please upload an Excel file (.xlsx or .xls)';
            fileInfo.style.background = '#fed7d7';
            fileInfo.style.color = '#c53030';
            fileInfo.classList.add('show');
            processBtn.disabled = true;
            return;
        }
        
        // Display file info
        fileInfo.innerHTML = `
            <i class="fas fa-file-excel"></i>
            <strong>${fileName}</strong> (${fileSize} KB)
        `;
        fileInfo.style.background = '#e6fffa';
        fileInfo.style.color = '#234e52';
        fileInfo.classList.add('show');
        processBtn.disabled = false;
    }
}

// Form submission
const uploadForm = document.getElementById('uploadForm');
if (uploadForm) {
    uploadForm.addEventListener('submit', function(e) {
        processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
        processBtn.disabled = true;
    });
}
