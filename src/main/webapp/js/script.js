"use strict";

document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM fully loaded');

    // --- Global Variables ---
    const dynamicContentArea = document.getElementById('dynamicContent');
    const modal = document.getElementById('modal');
    const modalContent = document.getElementById('modal-content');
    let currentRecentFilesPage = 1;
    const recentFilesLimit = 10;
    let currentPreviewPage = 1;
    const previewRecordsPerPage = 5;
    let cachedPreviewData = null; // To store data for preview pagination
    let selectedTemplateIdForManage = null; // Global for manage category functionality

    // --- Utility Functions ---
    function showMessage(type, message, autoClose = true) {
        if (!modal || !modalContent) {
            console.error('Modal elements not found');
            return;
        }
        modalContent.innerHTML = '';
        const h5 = document.createElement('h5');
        h5.textContent = message;
        h5.className = 'mb-3';

        const button = document.createElement('button');
        button.className = type === 'success' ? 'btn-neon btn-green' : 'btn-neon btn-blue';
        button.textContent = 'OK';
        button.addEventListener('click', closeModal);

        modalContent.appendChild(h5);
        modalContent.appendChild(button);
        modal.style.display = 'flex';

        if (autoClose) {
            setTimeout(() => {
                closeModal();
            }, 3000);
        }
    }

    function closeModal() {
        if (modal && modalContent) {
            modal.style.display = 'none';
            modalContent.innerHTML = '';
        }
    }

    function applyShakeAndErrorText(inputElement, errorMessage, originalPlaceholder) {
        inputElement.classList.add('shake-red-neon');
        inputElement.setAttribute('placeholder', errorMessage);
        setTimeout(() => {
            inputElement.classList.remove('shake-red-neon');
            inputElement.setAttribute('placeholder', originalPlaceholder);
        }, 3000);
    }

    // --- Navbar User Info ---
    function loadUserInfo() {
        fetch('/ExcelUploader/getUserInfo')
            .then(response => response.json())
            .then(data => {
                const userNameElement = document.getElementById('loggedInUserName');
                if (data.loggedIn && userNameElement) {
                    userNameElement.textContent = `Welcome, ${data.username}!`;
                } else {
                    console.warn('User not logged in or info unavailable.');
                    window.location.href = 'login.html'; // Redirect if not logged in
                }
            })
            .catch(error => {
                console.error('Error fetching user info:', error);
                window.location.href = 'login.html'; // Redirect on error
            });
    }

    // --- Dynamic Content Loaders ---
    function renderLoadingState(message = "Loading content...") {
        return `
            <div class="text-center py-5">
                <h5>${message} <span class="spinner-border spinner-border-sm text-info" role="status" aria-hidden="true"></span></h5>
            </div>
        `;
    }

    // --- 1. Recently Uploaded Files ---
    function renderRecentFiles(page = 1) {
        dynamicContentArea.innerHTML = renderLoadingState("Loading recently uploaded files...");
        currentRecentFilesPage = page;

        fetch(`/ExcelUploader/recentFiles?page=${currentRecentFilesPage}&limit=${recentFilesLimit}`)
            .then(response => response.json())
            .then(data => {
                if (data.status === 'success') {
                    const files = data.files;
                    const totalRecords = data.totalRecords;
                    const totalPages = Math.ceil(totalRecords / recentFilesLimit);

                    let tableHtml = `
                        <h2 class="text-center mb-4"><i class="bi bi-clock-history"></i> Recently Uploaded Files</h2>
                        <div class="table-responsive">
                            <table class="table table-bordered table-striped">
                                <thead>
                                    <tr>
                                        <th>No.</th>
                                        <th>File Name</th>
                                        <th>Template Type</th>
                                        <th>Upload Date</th>
                                    </tr>
                                </thead>
                                <tbody>
                    `;

                    if (files.length === 0) {
                        tableHtml += `<tr><td colspan="4" class="text-center">No recently uploaded files found.</td></tr>`;
                    } else {
                        files.forEach((file, index) => {
                            tableHtml += `
                                <tr>
                                    <td>${(currentRecentFilesPage - 1) * recentFilesLimit + index + 1}</td>
                                    <td>${file.FileName}</td>
                                    <td>${file.TemplateCategory}</td>
                                    <td>${file.UploadDate}</td>
                                </tr>
                            `;
                        });
                    }

                    tableHtml += `
                                </tbody>
                            </table>
                        </div>
                    `;
                    dynamicContentArea.innerHTML = tableHtml;

                    renderPaginationControls(dynamicContentArea, totalPages, currentRecentFilesPage, renderRecentFiles);

                } else {
                    dynamicContentArea.innerHTML = `<p class="text-danger text-center">${data.message || 'Failed to load recent files.'}</p>`;
                }
            })
            .catch(error => {
                console.error('Error loading recent files:', error);
                dynamicContentArea.innerHTML = `<p class="text-danger text-center">Error loading recent files: ${error.message}</p>`;
            });
    }

    // --- 2. Upload Files Section ---
    function renderUploadForm() {
        dynamicContentArea.innerHTML = `
            <h2 class="text-center mb-4"><i class="bi bi-cloud-arrow-up"></i> Upload New Excel Files</h2>
            <form id="uploadForm" enctype="multipart/form-data">
                <div class="file-upload-wrapper mb-3">
                    <label class="file-input-custom">
                        <input type="file" id="fileInput" name="files" accept=".xlsx" multiple>
                        <div class="file-input-content">
                            <i class="bi bi-file-earmark-arrow-up upload-icon"></i>
                            <span class="upload-text">Choose Excel File(s)</span>
                        </div>
                    </label>
                </div>
                <button type="submit" class="btn-neon btn-blue w-100"><i class="bi bi-upload"></i> Upload</button>
            </form>
            <div id="uploadFeedback" class="mt-3 text-center"></div>
        `;

        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            fileInput.addEventListener('change', function() {
                const files = this.files;
                const uploadText = document.querySelector('#uploadForm .upload-text');
                const fileInputCustom = this.closest('.file-input-custom');

                if (files.length > 0) {
                    let fileNameDisplay = files.length === 1 ? files[0].name : `${files.length} files selected`;
                    uploadText.textContent = fileNameDisplay;
                    if (fileInputCustom) {
                        fileInputCustom.classList.add('file-selected');
                        setTimeout(() => {
                            fileInputCustom.classList.remove('file-selected');
                        }, 1000);
                    }
                } else {
                    uploadText.textContent = 'Choose Excel File(s)';
                    if (fileInputCustom) {
                        fileInputCustom.classList.remove('file-selected');
                    }
                }
                document.getElementById('uploadFeedback').textContent = ''; // Clear feedback
            });
        }

        const uploadForm = document.getElementById('uploadForm');
        if (uploadForm) {
            uploadForm.addEventListener('submit', handleUploadSubmit);
        }
    }

    function handleUploadSubmit(event) {
        event.preventDefault();

        const fileInput = document.getElementById('fileInput');
        const fileInputWrapper = fileInput.closest('.file-input-custom');
        const uploadText = fileInputWrapper.querySelector('.upload-text');

        const files = fileInput.files;

        if (files.length === 0) {
            fileInputWrapper.classList.add('shake-red-neon');
            uploadText.style.color = '#ff4d4d';
            uploadText.textContent = 'Please select files to upload.';
            setTimeout(() => {
                fileInputWrapper.classList.remove('shake-red-neon');
                uploadText.style.color = '';
                uploadText.textContent = 'Choose Excel File(s)';
            }, 3000);
            return;
        }

        let allFilesValid = true;
        for (let i = 0; i < files.length; i++) {
            if (!files[i].name.toLowerCase().endsWith('.xlsx')) {
                allFilesValid = false;
                break;
            }
        }

        if (!allFilesValid) {
            fileInputWrapper.classList.add('shake-red-neon');
            uploadText.style.color = '#ff4d4d';
            uploadText.textContent = 'Only .xlsx files are allowed.';
            setTimeout(() => {
                fileInputWrapper.classList.remove('shake-red-neon');
                uploadText.style.color = '';
                uploadText.textContent = 'Choose Excel File(s)';
            }, 3000);
            return;
        }

        showUploadConfirmation(files);
    }

    function showUploadConfirmation(files) {
        modalContent.innerHTML = '';
        const h5 = document.createElement('h5');
        h5.textContent = `Upload ${files.length} file(s)?`;

        const fileListUl = document.createElement('ul');
        fileListUl.className = 'list-unstyled text-start mb-4 scrollable-list';
        const filesToDisplay = files.length > 5 ? Array.from(files).slice(0, 5) : Array.from(files);
        filesToDisplay.forEach(file => {
            const li = document.createElement('li');
            li.textContent = '• ' + file.name;
            fileListUl.appendChild(li);
        });
        if (files.length > 5) {
            const li = document.createElement('li');
            li.textContent = `...and ${files.length - 5} more.`;
            fileListUl.appendChild(li);
        }

        const buttonContainer = document.createElement('div');
        buttonContainer.className = 'd-flex justify-content-center gap-3';

        const yesButton = document.createElement('button');
        yesButton.className = 'btn-neon btn-blue';
        yesButton.textContent = 'Yes, Upload';
        yesButton.addEventListener('click', function() {
            confirmUpload(files);
        });

        const cancelButton = document.createElement('button');
        cancelButton.className = 'btn-neon btn-red';
        cancelButton.textContent = 'Cancel';
        cancelButton.addEventListener('click', closeModal);

        buttonContainer.appendChild(yesButton);
        buttonContainer.appendChild(cancelButton);
        modalContent.appendChild(h5);
        modalContent.appendChild(fileListUl);
        buttonContainer.appendChild(yesButton);
        buttonContainer.appendChild(cancelButton);
        modalContent.appendChild(buttonContainer);
        modal.style.display = 'flex';
    }

    function showUploadSuccessModal(uploadedFilesCount, uploadedFilesInfo) {
        modalContent.innerHTML = ''; // Clear existing content

        const isSingleFile = uploadedFilesCount === 1;
        const headingText = isSingleFile ? "Selected File Uploaded Successfully" : "Selected Files Uploaded Successfully";

        const h5 = document.createElement('h5');
        h5.className = 'mb-3';
        h5.textContent = headingText;
        modalContent.appendChild(h5);

        if (uploadedFilesInfo && uploadedFilesInfo.length > 0) {
            const fileListDiv = document.createElement('div');
            fileListDiv.className = 'scrollable-list mb-3';
            const ul = document.createElement('ul');
            ul.className = 'list-unstyled text-start mb-0';

            uploadedFilesInfo.forEach(fileInfo => {
                const li = document.createElement('li');
                li.textContent = `• ${fileInfo.fileName} (Type: ${fileInfo.templateCategory})`;
                ul.appendChild(li);
            });
            fileListDiv.appendChild(ul);
            modalContent.appendChild(fileListDiv);
        }

        const okButton = document.createElement('button');
        okButton.className = 'btn-neon btn-green';
        okButton.textContent = 'OK';
        okButton.addEventListener('click', () => {
            closeModal();
            renderUploadForm(); // Call renderUploadForm here
            const uploadButton = document.querySelector('.nav-link-btn[data-target="upload"]');
            if (uploadButton) {
                document.querySelectorAll('.nav-link-btn').forEach(btn => btn.classList.remove('active'));
                uploadButton.classList.add('active');
            }
        });
        modalContent.appendChild(okButton);

        modal.style.display = 'flex'; // Show the modal
    }

    function confirmUpload(files) {
        modalContent.innerHTML = `<h5>Uploading... <span class="spinner-border spinner-border-sm text-info" role="status" aria-hidden="true"></span></h5><p>Please wait. Do not close this window.</p>`;
        modal.style.display = 'flex';

        const totalFiles = files.length;
        let simulatedProcessedCount = 0;
        let intervalId;
        const timePerFileMs = 700;

        function updateProgressMessage() {
            let message;
            if (simulatedProcessedCount < totalFiles) {
                message = `Processing file ${simulatedProcessedCount + 1} of ${totalFiles}...`;
            } else {
                message = `Finalizing upload...`;
            }
            modalContent.innerHTML = `<h5>${message} <span class="spinner-border spinner-border-sm text-info" role="status" aria-hidden="true"></span></h5><p>Please wait. Do not close this window.</p>`;
            modal.style.display = 'flex';
        }

        updateProgressMessage();

        intervalId = setInterval(() => {
            simulatedProcessedCount++;
            if (simulatedProcessedCount <= totalFiles) {
                updateProgressMessage();
            }
        }, timePerFileMs);

        const formData = new FormData();
        for (let i = 0; i < files.length; i++) {
            formData.append('files', files[i]);
        }

        fetch('/ExcelUploader/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            clearInterval(intervalId);
            if (!response.ok) {
                return response.json().then(errorData => {
                    throw new Error(errorData.message || `Upload failed: HTTP ${response.status}`);
                });
            }
            return response.json();
        })
        .then(data => {
            if (data.status === 'success') {
                const uploadForm = document.getElementById('uploadForm');
                if (uploadForm) {
                    uploadForm.reset();
                    const uploadText = document.querySelector('#uploadForm .upload-text');
                    if (uploadText) uploadText.textContent = 'Choose Excel File(s)';
                    const fileInputCustom = document.querySelector('#uploadForm .file-input-custom');
                    if (fileInputCustom) fileInputCustom.classList.remove('file-selected');
                }
                showUploadSuccessModal(data.uploadedFilesCount, data.uploadedFilesInfo);
            } else {
                throw new Error(data.message || 'Upload failed.');
            }
        })
        .catch(error => {
            clearInterval(intervalId);
            console.error('Upload error:', error);
            showMessage('error', error.message || 'An unexpected error occurred during upload.', false);
        });
    }


    // --- 3. View Data Section ---
    function renderViewForm() {
        dynamicContentArea.innerHTML = `
            <h2 class="text-center mb-2"><i class="bi bi-file-earmark-spreadsheet"></i> View Excel Data</h2>

            <div class="row align-items-end">
                <div class="col-md-6 d-flex flex-column">
                    <label for="templateSelectView" class="form-label">Select Template Category</label>
                    <select id="templateSelectView" class="form-select"></select>
                </div>
                <div class="col-md-6 d-flex flex-column">
                    <label for="sheetSelectPreview" class="form-label">Select Sheet</label>
                    <select id="sheetSelectPreview" class="form-select" disabled>
                        <option value="">Select a Category first</option>
                    </select>
                </div>
            </div>

            <div class="view-scroll-area">
                <div id="previewTableContainer">
                    <p class="text-center text-muted">Select a template category to view data.</p>
                </div>
            </div>
            <div id="viewPaginationContainer"></div>
        `;
        loadTemplatesForView();
    }

    function loadTemplatesForView() {
        const select = document.getElementById('templateSelectView');
        const sheetSelectPreview = document.getElementById('sheetSelectPreview');
        if (!select || !sheetSelectPreview) return;

        select.innerHTML = '<option value="">Select a Category</option>';
        sheetSelectPreview.innerHTML = '<option value="">Select a Category first</option>';
        sheetSelectPreview.disabled = true;

        fetch('/ExcelUploader/templates', { cache: 'no-cache' })
            .then(response => {
                if (!response.ok) return response.json().then(errorData => { throw new Error(errorData.message); });
                return response.json();
            })
            .then(data => {
                if (!Array.isArray(data)) {
                    console.error('API response for templates is not an array:', data);
                    showMessage('error', 'Failed to load templates: Invalid server response format.', true);
                    return;
                }

                console.log("Templates for View loaded:", data);
                data.forEach(template => {
                    const option = document.createElement('option');
                    if (template && template.TemplateCategory && template.SampleTemplateID !== undefined) {
                        option.value = template.TemplateCategory;
                        option.textContent = template.TemplateCategory;
                        option.setAttribute('data-template-id', template.SampleTemplateID);
                        select.appendChild(option);
                    } else {
                        console.warn('Skipping malformed template data:', template);
                    }
                });
                if (!select.hasAttribute('data-listener-attached')) {
                    select.addEventListener('change', function() {
                        console.log("Template category selected:", this.value);
                        if (this.value) {
                            loadPreviewData(this.value, null);
                        } else {
                            document.getElementById('previewTableContainer').innerHTML = '<p class="text-center text-muted">Select a template category to view data.</p>';
                            sheetSelectPreview.innerHTML = '<option value="">Select a Category first</option>';
                            sheetSelectPreview.disabled = true;
                            document.getElementById('viewPaginationContainer').innerHTML = '';
                        }
                    });
                    select.setAttribute('data-listener-attached', 'true');
                }
            })
            .catch(error => {
                console.error('Error loading templates for view:', error);
                showMessage('error', 'Failed to load template categories: ' + error.message, true);
            });
    }

    function loadPreviewData(templateCategory, sheetName = null) {
        const previewTableContainer = document.getElementById('previewTableContainer');
        previewTableContainer.innerHTML = renderLoadingState("Loading preview data...");
        document.getElementById('viewPaginationContainer').innerHTML = '';
        currentPreviewPage = 1;
        cachedPreviewData = null;

        let url = `/ExcelUploader/preview?templateCategory=${encodeURIComponent(templateCategory)}`;
        if (sheetName) {
            url += `&sheetName=${encodeURIComponent(sheetName)}`;
        }

        console.log('Fetching preview data from:', url);

        fetch(url)
            .then(response => response.text().then(text => {
                console.log("Raw response text for preview:", text);
                return { status: response.status, ok: response.ok, text };
            }))
            .then(({ status, ok, text }) => {
                const sheetSelectPreview = document.getElementById('sheetSelectPreview');
                sheetSelectPreview.innerHTML = '';

                if (!ok) {
                    console.error(`HTTP ${status}: ${text || 'No response body'}`);
                    let errorMsg = 'Unable to fetch data.';
                    try {
                        const errorJson = JSON.parse(text);
                        errorMsg = errorJson.message || errorMsg;
                    } catch (e) {
                        // Not JSON, use raw text or default
                    }
                    previewTableContainer.innerHTML = `<p class="text-danger text-center">Server error: ${errorMsg}</p>`;
                    sheetSelectPreview.disabled = true;
                    return;
                }
                if (!text.trim()) {
                    console.error('Empty response received');
                    previewTableContainer.innerHTML = '<p class="text-danger text-center">Server returned empty response.</p>';
                    sheetSelectPreview.disabled = true;
                    return;
                }
                try {
                    cachedPreviewData = JSON.parse(text);
                    console.log("Parsed preview data:", cachedPreviewData);
                    if (cachedPreviewData.status === 'success') {
                        if (cachedPreviewData.sheetNames && cachedPreviewData.sheetNames.length > 0) {
                            cachedPreviewData.sheetNames.forEach(name => {
                                const option = document.createElement('option');
                                option.value = name;
                                option.textContent = name;
                                if (name === (sheetName || cachedPreviewData.sheetNames[0])) {
                                    option.selected = true;
                                }
                                sheetSelectPreview.appendChild(option);
                            });
                            sheetSelectPreview.disabled = false;
                            if (!sheetSelectPreview.hasAttribute('data-listener-attached')) {
                                sheetSelectPreview.addEventListener('change', function() {
                                    loadPreviewData(templateCategory, this.value);
                                });
                                sheetSelectPreview.setAttribute('data-listener-attached', 'true');
                            }
                        } else {
                            sheetSelectPreview.innerHTML = '<option value="">No sheets found</option>';
                            sheetSelectPreview.disabled = true;
                        }

                        renderPreviewTable(cachedPreviewData.sheetNames, cachedPreviewData.data, templateCategory);

                    } else {
                        previewTableContainer.innerHTML = `<p class="text-danger text-center">${cachedPreviewData.message || 'No data available for this template category.'}</p нему>`;
                        sheetSelectPreview.disabled = true;
                    }
                } catch (e) {
                    console.error('Failed to parse JSON response:', text, e);
                    previewTableContainer.innerHTML = '<p class="text-danger text-center">Invalid server response format.</p>';
                    sheetSelectPreview.disabled = true;
                }
            })
            .catch(error => {
                console.error('Error loading preview:', error.message);
                previewTableContainer.innerHTML = `<p class="text-danger text-center">Error loading preview data: ${error.message}</p>`;
                document.getElementById('sheetSelectPreview').disabled = true;
            });
    }

    function renderPreviewTable(sheetNames, sheetData, templateCategory) {
        const previewTableContainer = document.getElementById('previewTableContainer');
        const viewPaginationContainer = document.getElementById('viewPaginationContainer');

        if (!previewTableContainer || !viewPaginationContainer) {
            console.error("One or more required containers not found for rendering preview table.");
            return;
        }

        previewTableContainer.innerHTML = '';

        if (!sheetData || !Array.isArray(sheetData.Headers) || !Array.isArray(sheetData.Rows)) {
            console.warn('Invalid sheet data for rendering:', sheetData);
            previewTableContainer.innerHTML = '<p class="text-center text-muted">No data available for this sheet.</p>';
            viewPaginationContainer.innerHTML = '';
            return;
        }

        let totalCols = 0;
        if (sheetData.Headers.length > 0) {
            const lastHeaderRow = sheetData.Headers[sheetData.Headers.length - 1];
            lastHeaderRow.forEach(header => {
                totalCols += (header.colspan || 1);
            });
        }
        if (totalCols === 0) totalCols = 5;

        const tableWrapper = document.createElement('div');
        tableWrapper.className = 'table-responsive';

        const table = document.createElement('table');
        table.className = 'table table-bordered table-striped';

        const thead = document.createElement('thead');
        sheetData.Headers.forEach(headerRow => {
            const tr = document.createElement('tr');
            headerRow.forEach(header => {
                const th = document.createElement('th');
                th.textContent = header.label || '';
                if (header.colspan > 1) th.setAttribute('colspan', header.colspan);
                if (header.rowspan) th.setAttribute('rowspan', header.rowspan);
                th.style.whiteSpace = 'nowrap';
                tr.appendChild(th);
            });
            thead.appendChild(tr);
        });
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        table.appendChild(tbody);
        tableWrapper.appendChild(table);
        previewTableContainer.appendChild(tableWrapper);

        const previewPageChangeCallback = (page) => {
            currentPreviewPage = page;
            renderTableRowsForPreview(tbody, sheetData.Rows, currentPreviewPage, previewRecordsPerPage, totalCols);
            renderPaginationControls(viewPaginationContainer, Math.ceil(sheetData.Rows.length / previewRecordsPerPage), currentPreviewPage, previewPageChangeCallback);
        };

        renderTableRowsForPreview(tbody, sheetData.Rows, currentPreviewPage, previewRecordsPerPage, totalCols);
        renderPaginationControls(viewPaginationContainer, Math.ceil(sheetData.Rows.length / previewRecordsPerPage), currentPreviewPage, previewPageChangeCallback);
    }

    function renderTableRowsForPreview(tbody, rows, page, limit, totalCols) {
        tbody.innerHTML = '';
        const start = (page - 1) * limit;
        const end = Math.min(start + limit, rows.length);
        const paginatedRows = rows.slice(start, end);

        if (paginatedRows.length === 0) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = totalCols;
            td.textContent = 'No rows available for this page.';
            td.style.textAlign = 'center';
            tr.appendChild(td);
            tbody.appendChild(tr);
        } else {
            paginatedRows.forEach(row => {
                const tr = document.createElement('tr');
                row.forEach(value => {
                    const td = document.createElement('td');
                    td.textContent = value !== undefined && value !== null ? String(value) : '';
                    td.setAttribute('title', value !== undefined && value !== null ? String(value) : '');
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }
    }


    // --- 4. Download Files Section ---
    function renderDownloadForm() {
        dynamicContentArea.innerHTML = `
            <h2 class="text-center mb-4"><i class="bi bi-download"></i> Download Excel Files</h2>
            <div class="mb-3">
                <label for="templateSelectDownload" class="form-label">Select Template Category to Download</label>
                <select id="templateSelectDownload" class="form-select"></select>
            </div>
            <div class="d-flex justify-content-center">
                <button type="button" class="btn-neon btn-green" id="downloadBtn"><i class="bi bi-download"></i> Download</button>
            </div>
        `;
        loadTemplatesForDownload();
    }

    function loadTemplatesForDownload() {
        const select = document.getElementById('templateSelectDownload');
        if (!select) return;

        select.innerHTML = '<option value="">Select a Category</option>';
        fetch('/ExcelUploader/templates', { cache: 'no-cache' })
            .then(response => {
                if (!response.ok) return response.json().then(errorData => { throw new Error(errorData.message); });
                return response.json();
            })
            .then(data => {
                if (!Array.isArray(data)) {
                    console.error('API response for templates is not an array:', data);
                    showMessage('error', 'Failed to load templates for download: Invalid server response format.', true);
                    return;
                }

                data.forEach(template => {
                    const option = document.createElement('option');
                    if (template && template.TemplateCategory) {
                        option.value = template.TemplateCategory;
                        option.textContent = template.TemplateCategory;
                        select.appendChild(option);
                    } else {
                        console.warn('Skipping malformed template data for download:', template);
                    }
                });
                if (!document.getElementById('downloadBtn').hasAttribute('data-listener-attached')) {
                    document.getElementById('downloadBtn').addEventListener('click', handleDownload);
                    document.getElementById('downloadBtn').setAttribute('data-listener-attached', 'true');
                }
            })
            .catch(error => {
                console.error('Error loading templates for download:', error);
                showMessage('error', 'Failed to load template categories: ' + error.message, true);
            });
    }

    function showDownloadProgressModal() {
        modalContent.innerHTML = `<h5>Preparing your download... <span class="spinner-border spinner-border-sm text-info" role="status" aria-hidden="true"></span></h5><p>Please wait. Do not close this window.</p>`;
        modal.style.display = 'flex';
    }

    function handleDownload() {
        const select = document.getElementById('templateSelectDownload');
        const templateCategory = select.value;

        if (!templateCategory || templateCategory === 'null') {
            select.classList.add('shake-red-neon');
            select.style.color = '#ff4d4d';
            select.options[0].textContent = 'Please select a template category.';
            setTimeout(() => {
                select.classList.remove('shake-red-neon');
                select.style.color = '';
                select.options[0].textContent = 'Select a Template Category';
            }, 3000);
            return;
        }

        showDownloadProgressModal();

        fetch('/ExcelUploader/export?templateCategory=' + encodeURIComponent(templateCategory), { cache: 'no-cache' })
            .then(response => {
                if (!response.ok) {
                    const contentType = response.headers.get("content-type");
                    if (contentType && contentType.indexOf("application/json") !== -1) {
                        return response.json().then(errorData => {
                            throw new Error(errorData.message || `HTTP ${response.status}: Failed to initiate download`);
                        });
                    } else {
                        return response.text().then(text => {
                            throw new Error(`HTTP ${response.status}: Failed to initiate download. Server response: ${text.substring(0, Math.min(text.length, 200))}...`);
                        });
                    }
                }
                window.location.href = '/ExcelUploader/export?templateCategory=' + encodeURIComponent(templateCategory);

                closeModal();
                showMessage('success', 'Downloaded successfully!', false);

                return Promise.resolve();
            })
            .catch(error => {
                console.error('Download error:', error);
                showMessage('error', error.message || 'Error initiating download.', false);
            });
    }


    // --- 5. Manage Categories Section ---
    function renderManageForm() {
        dynamicContentArea.innerHTML = `
            <h2 class="text-center mb-4"><i class="bi bi-pencil-square"></i> Manage Template Categories</h2>
            <div class="mb-3">
                <label for="existingTemplateSelectManage" class="form-label">Select Template Category to Edit</label>
                <select id="existingTemplateSelectManage" class="form-select"></select>
            </div>
            <div class="mb-3">
                <label for="newCategoryName" class="form-label">New Template Category Name</label>
                <input type="text" class="form-control" id="newCategoryName" placeholder="Enter new category name">
            </div>
            <div class="d-flex justify-content-center">
                <button type="button" class="btn-neon btn-green" id="saveCategoryBtn"><i class="bi bi-floppy"></i> Save Changes</button>
            </div>
        `;
        loadTemplatesForManage();
    }

    function loadTemplatesForManage() {
        const existingTemplateSelect = document.getElementById('existingTemplateSelectManage');
        const newCategoryNameInput = document.getElementById('newCategoryName');
        if (!existingTemplateSelect || !newCategoryNameInput) return;

        selectedTemplateIdForManage = null;
        existingTemplateSelect.innerHTML = '<option value="">Select a Category</option>';

        fetch('/ExcelUploader/templates', { cache: 'no-cache' })
            .then(response => {
                if (!response.ok) {
                    return response.json().then(errorData => {
                        throw new Error(errorData.message || 'Network response was not ok: ' + response.statusText);
                    });
                }
                return response.json();
            })
            .then(data => {
                if (!Array.isArray(data)) {
                    console.error('API response for templates is not an array:', data);
                    showMessage('error', 'Failed to load templates for management: Invalid server response format.', true);
                    return;
                }

                data.forEach(template => {
                    const option = document.createElement('option');
                    if (template && template.TemplateCategory && template.SampleTemplateID !== undefined) {
                        option.value = template.TemplateCategory;
                        option.textContent = template.TemplateCategory;
                        option.setAttribute('data-template-id', template.SampleTemplateID);
                        existingTemplateSelect.appendChild(option);
                    } else {
                        console.warn('Skipping malformed template data for management:', template);
                    }
                });

                if (!existingTemplateSelect.hasAttribute('data-listener-attached')) {
                    existingTemplateSelect.addEventListener('change', function() {
                        this.classList.remove('shake-red-neon');
                        this.style.color = '';

                        const selectedOption = this.options[this.selectedIndex];
                        selectedTemplateIdForManage = selectedOption ? selectedOption.getAttribute('data-template-id') : null;
                        const currentCategoryName = selectedOption ? selectedOption.value : "";

                        newCategoryNameInput.value = currentCategoryName !== "" ? currentCategoryName : "";
                        newCategoryNameInput.classList.remove('shake-red-neon');
                        newCategoryNameInput.style.color = '';
                        newCategoryNameInput.setAttribute('placeholder', 'Enter new category name');
                    });
                    existingTemplateSelect.setAttribute('data-listener-attached', 'true');
                }

                const saveButton = document.getElementById('saveCategoryBtn');
                if (saveButton && !saveButton.hasAttribute('data-listener-attached')) {
                    saveButton.addEventListener('click', handleSaveCategory);
                    saveButton.setAttribute('data-listener-attached', 'true');
                }

            })
            .catch(error => {
                console.error('Error loading templates for management:', error);
                showMessage('error', 'Failed to load templates for management: ' + error.message, true);
            });
    }

    function handleSaveCategory() {
        const existingTemplateSelect = document.getElementById('existingTemplateSelectManage');
        const newCategoryNameInput = document.getElementById('newCategoryName');

        existingTemplateSelect.classList.remove('shake-red-neon');
        existingTemplateSelect.style.color = '';
        newCategoryNameInput.classList.remove('shake-red-neon');
        newCategoryNameInput.style.color = '';
        newCategoryNameInput.setAttribute('placeholder', 'Enter new category name');

        const selectedCategoryValue = existingTemplateSelect.value;
        const newName = newCategoryNameInput.value.trim();

        if (!selectedTemplateIdForManage || selectedCategoryValue === "") {
            existingTemplateSelect.classList.add('shake-red-neon');
            existingTemplateSelect.style.color = '#ff4d4d';
            showMessage('error', 'Please select a template category to edit.', true);
            return;
        }

        if (newName === "") {
            newCategoryNameInput.classList.add('shake-red-neon');
            newCategoryNameInput.style.color = '#ff4d4d';
            newCategoryNameInput.setAttribute('placeholder', 'New category name cannot be empty.');
            showMessage('error', 'New category name cannot be empty.', true);
            return;
        }

        showConfirmSaveDialog(selectedCategoryValue, newName);
    }

    function showConfirmSaveDialog(oldName, newName) {
        modalContent.innerHTML = '';
        const h5 = document.createElement('h5');
        h5.className = 'mb-3';
        h5.textContent = `Change category "${oldName}" to "${newName}"?`;

        const buttonContainer = document.createElement('div');
        buttonContainer.className = 'd-flex justify-content-center gap-3';

        const yesButton = document.createElement('button');
        yesButton.className = 'btn-neon btn-green';
        yesButton.textContent = 'Yes, Save';
        yesButton.addEventListener('click', function() {
            closeModal();
            saveTemplateCategory(selectedTemplateIdForManage, newName);
        });

        const cancelButton = document.createElement('button');
        cancelButton.className = 'btn-neon btn-blue';
        cancelButton.textContent = 'Cancel';
        cancelButton.addEventListener('click', closeModal);

        buttonContainer.appendChild(yesButton);
        buttonContainer.appendChild(cancelButton);
        modalContent.appendChild(h5);
        modalContent.appendChild(buttonContainer);
        modal.style.display = 'flex';
    }

    function saveTemplateCategory(templateId, newCategoryName) {
        modalContent.innerHTML = `<h5>Saving changes... <span class="spinner-border spinner-border-sm text-info" role="status" aria-hidden="true"></span></h5><p>Please wait. Do not close this window.</p>`;
        modal.style.display = 'flex';

        fetch('/ExcelUploader/updateTemplateCategory', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ templateId: templateId, newCategoryName: newCategoryName })
        })
        .then(response => {
            if (!response.ok) return response.json().then(errorData => { throw new Error(errorData.message); });
            return response.json();
        })
        .then(data => {
            if (data.status === 'success') {
                showMessage('success', data.message || 'Template category updated successfully!', false);
                renderManageForm();
            } else {
                showMessage('error', data.message || 'Error updating template category.', false);
            }
        })
        .catch(error => {
            console.error('Error updating template category:', error);
            showMessage('error', error.message || 'An unexpected error occurred during update.', false);
        });
    }

    // --- 6. Delete Entries Section ---
    function renderDeleteForm() {
        dynamicContentArea.innerHTML = `
            <h2 class="text-center mb-4"><i class="bi bi-trash"></i> Delete Excel Files</h2>
            <div class="mb-3">
                <label for="templateCategorySelectDelete" class="form-label">Select Template Category</label>
                <select id="templateCategorySelectDelete" class="form-select">
                    <option value="">Select a Category</option>
                </select>
            </div>
            <div class="mb-3 d-flex justify-content-between align-items-center">
                <label for="templateListDelete" class="form-label mb-0">Sheets in Selected Category</label>
                <div class="form-check d-flex align-items-center">
                    <label class="form-check-label" for="selectAllTemplates">Select All</label>
                    <input type="checkbox" class="form-check-input ms-2" id="selectAllTemplates">
                    <div id="templateCountDisplay" class="template-count-display ms-2">
                        <span id="selectedCount">0</span> / <span id="totalCount">0</span>
                    </div>
                </div>
            </div>
            <div id="templateListDelete" class="mb-3 scrollable-list" style="max-height: 200px;">
                <p class="text-muted text-center">Select a category to view sheets.</p>
            </div>
            <div class="d-flex justify-content-center">
                <button type="button" class="btn-neon btn-red" id="deleteBtn"><i class="bi bi-trash"></i> Delete </button>
            </div>
        `;
        loadTemplatesForDelete();
    }

    function loadTemplatesForDelete() {
        const categorySelect = document.getElementById('templateCategorySelectDelete');
        if (!categorySelect) {
            console.error('Template category select element for delete not found');
            return Promise.reject(new Error('Template category select element for delete not found'));
        }

        categorySelect.innerHTML = '<option value="">Select a Category</option>';

        return fetch('/ExcelUploader/templates', { cache: 'no-cache' })
            .then(response => {
                if (!response.ok) {
                    return response.json().then(errorData => {
                        throw new Error(errorData.message || 'Network response was not ok: ' + response.statusText);
                    });
                }
                return response.json();
            })
            .then(data => {
                if (!Array.isArray(data)) {
                    console.error('API response for templates is not an array:', data);
                    showMessage('error', 'Failed to load templates for deletion: Invalid server response format.', true);
                    return;
                }

                data.forEach(category => {
                    if (category && category.TemplateCategory && category.TemplateCategory !== 'null') {
                        const option = document.createElement('option');
                        option.value = category.TemplateCategory;
                        option.textContent = category.TemplateCategory;
                        categorySelect.appendChild(option);
                    } else {
                        console.warn('Skipping malformed category data for deletion:', category);
                    }
                });
                if (!categorySelect.hasAttribute('data-listener-attached')) {
                    categorySelect.addEventListener('change', function() {
                        const selectedCategory = this.value;
                        this.classList.remove('shake-red-neon');
                        this.style.color = '';
                        if (selectedCategory) {
                            loadTemplatesBySelectedCategoryForDelete(selectedCategory);
                        } else {
                            document.getElementById('templateListDelete').innerHTML = '<p class="text-muted text-center">Select a category to view sheets.</p>';
                            document.getElementById('selectAllTemplates').checked = false;
                            updateTemplateCounters();
                        }
                    });
                    categorySelect.setAttribute('data-listener-attached', 'true');
                }

                const selectAllTemplatesCheckbox = document.getElementById('selectAllTemplates');
                const templateListDeleteDiv = document.getElementById('templateListDelete');
                const deleteButton = document.getElementById('deleteBtn');

                if (!selectAllTemplatesCheckbox.hasAttribute('data-listener-attached')) {
                    selectAllTemplatesCheckbox.addEventListener('change', handleSelectAllTemplates);
                    selectAllTemplatesCheckbox.setAttribute('data-listener-attached', 'true');
                }
                if (!templateListDeleteDiv.hasAttribute('data-listener-attached')) {
                    templateListDeleteDiv.addEventListener('change', handleIndividualTemplateChange);
                    templateListDeleteDiv.setAttribute('data-listener-attached', 'true');
                }
                if (!deleteButton.hasAttribute('data-listener-attached')) {
                    deleteButton.addEventListener('click', handleDeleteSelected);
                    deleteButton.setAttribute('data-listener-attached', 'true');
                }


                updateTemplateCounters();
                return Promise.resolve();
            })
            .catch(error => {
                console.error('Error loading template categories for delete:', error);
                showMessage('error', 'Failed to load template categories: ' + error.message, true);
                return Promise.reject(error);
            });
    }

    function loadTemplatesBySelectedCategoryForDelete(templateCategory) {
        const templateList = document.getElementById('templateListDelete');
        if (!templateList) {
            console.error('Template list element for delete not found');
            return Promise.reject(new Error('Template list element for delete not found'));
        }

        templateList.classList.remove('shake-red-neon');
        templateList.innerHTML = '<p class="text-center text-muted">Loading sheets...</p>';
        document.getElementById('selectAllTemplates').checked = false;

        return fetch(`/ExcelUploader/templatesByCategory?templateCategory=${encodeURIComponent(templateCategory)}`)
            .then(response => {
                if (!response.ok) {
                    return response.json().then(errorData => {
                        throw new Error(errorData.message || 'Network response was not ok: ' + response.statusText);
                    });
                }
                return response.json();
            })
            .then(data => {
                templateList.innerHTML = '';
                if (data.length === 0) {
                    templateList.innerHTML = '<p class="text-muted text-center">No sheets found for this category.</p>';
                } else {
                    data.forEach(responseItem => {
                        const responseID = responseItem.ResponseID;
                        const originalFileName = responseItem.OriginalFileName || 'Unnamed File';
                        const sheetName = responseItem.SheetName || 'Unnamed Sheet';
                        const displayName = `${originalFileName} (Sheet: ${sheetName})`;

                        const div = document.createElement('div');
                        div.className = 'form-check';
                        div.style.marginBottom = '8px';

                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox';
                        checkbox.className = 'form-check-input template-checkbox';
                        checkbox.value = responseID;
                        checkbox.id = `response-${responseID}`;
                        checkbox.setAttribute('data-display-name', displayName);

                        const label = document.createElement('label');
                        label.className = 'form-check-label';
                        label.htmlFor = `response-${responseID}`;
                        label.textContent = displayName;

                        div.appendChild(checkbox);
                        div.appendChild(label);
                        templateList.appendChild(div);
                    });
                }
                updateTemplateCounters();
                return Promise.resolve();
            })
            .catch(error => {
                console.error('Error loading templates by category for delete:', error);
                templateList.innerHTML = `<p class="text-danger text-center">Error loading sheets: ${error.message}</p>`;
                updateTemplateCounters();
                return Promise.reject(error);
            });
    }

    function updateTemplateCounters() {
        const allCheckboxes = document.querySelectorAll('#templateListDelete input[type="checkbox"].template-checkbox');
        const selectedCheckboxes = document.querySelectorAll('#templateListDelete input[type="checkbox"].template-checkbox:checked');

        const totalCountElement = document.getElementById('totalCount');
        const selectedCountElement = document.getElementById('selectedCount');

        if (totalCountElement) totalCountElement.textContent = allCheckboxes.length;
        if (selectedCountElement) selectedCountElement.textContent = selectedCheckboxes.length;
    }

    function handleSelectAllTemplates() {
        const selectAllCheckbox = document.getElementById('selectAllTemplates');
        const checkboxes = document.querySelectorAll('#templateListDelete input[type="checkbox"].template-checkbox');
        checkboxes.forEach(cb => {
            cb.checked = selectAllCheckbox.checked;
        });
        updateTemplateCounters();
    }

    function handleIndividualTemplateChange(event) {
        if (event.target.type === 'checkbox' && event.target.classList.contains('template-checkbox')) {
            document.getElementById('templateListDelete').classList.remove('shake-red-neon');
            updateTemplateCounters();
            const selectAll = document.getElementById('selectAllTemplates');
            if (selectAll) {
                const allCheckboxes = document.querySelectorAll('#templateListDelete input[type="checkbox"].template-checkbox');
                const checkedCheckboxes = document.querySelectorAll('#templateListDelete input[type="checkbox"].template-checkbox:checked');
                selectAll.checked = (allCheckboxes.length > 0 && checkedCheckboxes.length === allCheckboxes.length);
            }
        }
    }

    function handleDeleteSelected() {
        const templateList = document.getElementById('templateListDelete');
        const categorySelect = document.getElementById('templateCategorySelectDelete');

        categorySelect.classList.remove('shake-red-neon');
        categorySelect.style.color = '';
        templateList.classList.remove('shake-red-neon');

        if (!categorySelect || !categorySelect.value || categorySelect.value === 'null' || categorySelect.value === '') {
            categorySelect.classList.add('shake-red-neon');
            categorySelect.style.color = '#ff4d4d';
            showMessage('error', 'Please select a category to delete from.', true);
            return;
        }

        const checkboxes = templateList.querySelectorAll('input[type="checkbox"].template-checkbox:checked');
        if (checkboxes.length === 0) {
            templateList.classList.add('shake-red-neon');
            showMessage('error', 'Please select at least one sheet to delete.', true);
            return;
        }

        const responseIdsToDelete = Array.from(checkboxes).map(cb => parseInt(cb.value, 10));
        const displayNames = Array.from(checkboxes).map(cb => cb.getAttribute('data-display-name') || 'Unnamed Entry');

        showDeleteWarning(responseIdsToDelete, displayNames);
    }

    function showDeleteWarning(responseIds, displayNames) {
        modalContent.innerHTML = '';
        const h5 = document.createElement('h5');
        h5.className = 'mb-3';
        h5.textContent = 'Delete the following uploaded files/sheets?';

        const scrollableDiv = document.createElement('div');
        scrollableDiv.className = 'scrollable-list';
        const ul = document.createElement('ul');
        ul.className = 'list-unstyled text-start mb-0';
        displayNames.forEach(name => {
            const li = document.createElement('li');
            li.textContent = '• ' + name;
            ul.appendChild(li);
        });
        scrollableDiv.appendChild(ul);

        const buttonContainer = document.createElement('div');
        buttonContainer.className = 'd-flex justify-content-center gap-3 mt-3';

        const yesButton = document.createElement('button');
        yesButton.className = 'btn-neon btn-red';
        yesButton.textContent = 'Yes, Delete';
        yesButton.addEventListener('click', function() {
            confirmDelete(responseIds);
        });

        const cancelButton = document.createElement('button');
        cancelButton.className = 'btn-neon btn-blue';
        cancelButton.textContent = 'Cancel';
        cancelButton.addEventListener('click', closeModal);

        buttonContainer.appendChild(yesButton);
        buttonContainer.appendChild(cancelButton);
        modalContent.appendChild(h5);
        modalContent.appendChild(scrollableDiv);
        modalContent.appendChild(buttonContainer);
        modal.style.display = 'flex';
    }

    function showDeleteSuccessModal(dynamicMessage) {
        modalContent.innerHTML = ''; // Clear existing content

        const heading = document.createElement('h5');
        heading.textContent = 'Deletion Successful!'; // Fixed heading
        heading.className = 'mb-2'; 

        const messageParagraph = document.createElement('p');
        messageParagraph.textContent = dynamicMessage; // Dynamic message
        messageParagraph.className = 'mb-3';

        const okButton = document.createElement('button');
        okButton.className = 'btn-neon btn-green';
        okButton.textContent = 'OK';
        okButton.addEventListener('click', () => {
            closeModal();
            renderDeleteForm(); // Re-render the delete form to show updated lists
        });

        modalContent.appendChild(heading);
        modalContent.appendChild(messageParagraph);
        modalContent.appendChild(okButton);
        modal.style.display = 'flex'; // Show the modal
    }


    function confirmDelete(responseIds) {
        modalContent.innerHTML = `<h5>Deleting... <span class="spinner-border spinner-border-sm text-info" role="status" aria-hidden="true"></span></h5><p>Please wait. Do not close this window.</p>`;
        modal.style.display = 'flex';

        const categorySelect = document.getElementById('templateCategorySelectDelete');
        const currentCategory = categorySelect ? categorySelect.value : null;

        fetch('/ExcelUploader/delete', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ responseIds: responseIds })
        })
        .then(response => {
            if (!response.ok) {
                return response.text().then(text => {
                    console.error("Delete failed. Raw response text:", text);
                    try {
                        const errorData = JSON.parse(text);
                        throw new Error(errorData.message || `Delete failed: HTTP ${response.status}`);
                    } catch (e) {
                        throw new Error(`Delete failed: HTTP ${response.status} - Server response was not JSON. Content: ${text.substring(0, Math.min(text.length, 200))}`);
                    }
                });
            }
            return response.json();
        })
        .then(data => {
            const message = data.message || `Successfully deleted ${data.deletedCount} file(s)/sheet(s).`;
            showDeleteSuccessModal(message);

            loadTemplatesForDelete().then(() => {
                const categorySelectRefreshed = document.getElementById('templateCategorySelectDelete');
                const templateList = document.getElementById('templateListDelete');

                const categoryStillExists = currentCategory && Array.from(categorySelectRefreshed.options).some(option => option.value === currentCategory);

                if (categoryStillExists) {
                    categorySelectRefreshed.value = currentCategory;
                    loadTemplatesBySelectedCategoryForDelete(currentCategory).catch(sheetLoadError => {
                        console.error("Error reloading sheets after category re-selection:", sheetLoadError);
                        templateList.innerHTML = `<p class="text-danger text-center">Error reloading sheets: ${sheetLoadError.message}</p>`;
                        updateTemplateCounters();
                    });
                } else {
                    categorySelectRefreshed.value = '';
                    templateList.innerHTML = '<p class="text-muted text-center">Select a category to view sheets.</p>';
                    document.getElementById('selectAllTemplates').checked = false;
                    updateTemplateCounters();
                }
            }).catch(error => {
                console.error("Error refreshing categories after delete:", error);
                showMessage('error', 'Error refreshing category list after deletion: ' + error.message, false);
                const categorySelectRefreshed = document.getElementById('templateCategorySelectDelete');
                const templateList = document.getElementById('templateListDelete');
                if (categorySelectRefreshed) categorySelectRefreshed.value = '';
                if (templateList) templateList.innerHTML = '<p class="text-muted text-center">Select a category to view sheets.</p>';
                document.getElementById('selectAllTemplates').checked = false;
                updateTemplateCounters();
            });
        })
        .catch(error => {
            console.error('Error during batch deletion (outer catch):', error);
            showMessage('error', error.message || 'An unexpected network error occurred during deletion.', false);
            updateTemplateCounters();
        });
    }

    // --- Pagination Helper ---
    function renderPaginationControls(containerElement, totalPages, currentPage, pageChangeCallback) {
        let paginationDiv = containerElement.querySelector('.pagination-controls');
        if (!paginationDiv) {
            paginationDiv = document.createElement('div');
            paginationDiv.className = 'pagination-controls d-flex justify-content-center align-items-center mt-2';
            containerElement.appendChild(paginationDiv);
        }
        paginationDiv.innerHTML = '';

        const ul = document.createElement('ul');
        ul.className = 'pagination mb-0';

        const createPageItem = (text, pageNum, isActive = false, isDisabled = false) => {
            const li = document.createElement('li');
            li.className = `page-item ${isActive ? 'active' : ''} ${isDisabled ? 'disabled' : ''}`;
            const link = document.createElement('a');
            link.className = 'page-link';
            link.href = '#';
            link.textContent = text;
            if (!isDisabled) {
                link.addEventListener('click', (e) => {
                    e.preventDefault();
                    pageChangeCallback(pageNum);
                });
            }
            li.appendChild(link);
            return li;
        };

        ul.appendChild(createPageItem('Previous', currentPage - 1, false, currentPage === 1));

        if (totalPages <= 3) {
            for (let i = 1; i <= totalPages; i++) {
                ul.appendChild(createPageItem(String(i), i, i === currentPage));
            }
        } else {
            ul.appendChild(createPageItem('1', 1, 1 === currentPage));

            let startPage = Math.max(2, currentPage - 1);
            let endPage = Math.min(totalPages - 1, currentPage + 1);

            if (currentPage <= 3) {
                startPage = 2;
                endPage = 3;
            } else if (currentPage >= totalPages - 2) {
                startPage = totalPages - 3;
                endPage = totalPages - 1;
            }

            if (startPage > 2) {
                const ellipsisLi = document.createElement('li');
                ellipsisLi.className = 'page-item disabled';
                const ellipsisSpan = document.createElement('span');
                ellipsisSpan.className = 'page-link';
                ellipsisSpan.textContent = '...';
                ellipsisLi.appendChild(ellipsisSpan);
                ul.appendChild(ellipsisLi);
            }

            for (let i = startPage; i <= endPage; i++) {
                ul.appendChild(createPageItem(String(i), i, i === currentPage));
            }

            if (endPage < totalPages - 1) {
                const ellipsisLi = document.createElement('li');
                ellipsisLi.className = 'page-item disabled';
                const ellipsisSpan = document.createElement('span');
                ellipsisSpan.className = 'page-link';
                ellipsisSpan.textContent = '...';
                ellipsisLi.appendChild(ellipsisSpan);
                ul.appendChild(ellipsisLi);
            }

            if (totalPages > 1) {
                ul.appendChild(createPageItem(String(totalPages), totalPages, totalPages === currentPage));
            }
        }

        ul.appendChild(createPageItem('Next', currentPage + 1, false, currentPage === totalPages));

        paginationDiv.appendChild(ul);

        if (totalPages > 3) {
            const inputGroup = document.createElement('div');
            inputGroup.className = 'input-group ms-2';
            inputGroup.style.maxWidth = '180px';

            const inputSpan = document.createElement('span');
            inputSpan.className = 'input-group-text';
            inputSpan.textContent = 'Go to Page:';
            inputSpan.style.padding = '0.375rem 0.75rem';
            inputSpan.style.fontSize = '0.875rem';
            inputSpan.style.borderTopLeftRadius = '0.25rem';
            inputSpan.style.borderBottomLeftRadius = '0.25rem';

            const pageInput = document.createElement('input');
            pageInput.type = 'number';
            pageInput.className = 'form-control';
            pageInput.min = 1;
            pageInput.max = totalPages;
            pageInput.value = currentPage;
            pageInput.style.textAlign = 'center';
            pageInput.style.padding = '0.375rem 0.75rem';
            pageInput.style.fontSize = '0.875rem';
            pageInput.style.width = '70px';
            pageInput.style.borderTopRightRadius = '0.25rem';
            pageInput.style.borderBottomRightRadius = '0.25rem';

            pageInput.addEventListener('change', (e) => {
                let page = parseInt(e.target.value, 10);
                if (isNaN(page) || page < 1) {
                    page = 1;
                } else if (page > totalPages) {
                    page = totalPages;
                }
                e.target.value = page;
                pageChangeCallback(page);
            });

            pageInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    pageInput.dispatchEvent(new Event('change'));
                }
            });

            inputGroup.appendChild(inputSpan);
            inputGroup.appendChild(pageInput);
            paginationDiv.appendChild(inputGroup);
        }
    }

    // --- Event Listeners for Left Sidebar Navigation ---
    const navButtons = document.querySelectorAll('.nav-link-btn');
    navButtons.forEach(button => {
        button.addEventListener('click', function() {
            navButtons.forEach(btn => btn.classList.remove('active'));
            this.classList.add('active');
            const target = this.getAttribute('data-target');
            switch (target) {
                case 'recentFiles':
                    renderRecentFiles(1);
                    break;
                case 'upload':
                    renderUploadForm();
                    break;
                case 'view':
                    renderViewForm();
                    break;
                case 'download':
                    renderDownloadForm();
                    break;
                case 'manage':
                    renderManageForm();
                    break;
                case 'delete':
                    renderDeleteForm();
                    break;
                default:
                    dynamicContentArea.innerHTML = `<p class="text-center text-muted">Select an option from the sidebar.</p>`;
            }
        });
    });

    // --- Initial Setup ---
    loadUserInfo();
    // Load the default content (Recently Uploaded Files) when the page loads
    renderRecentFiles(1);
    const initialRecentFilesButton = document.querySelector('.nav-link-btn[data-target="recentFiles"]');
    if (initialRecentFilesButton) {
        navButtons.forEach(btn => btn.classList.remove('active'));
        initialRecentFilesButton.classList.add('active');
    }
});