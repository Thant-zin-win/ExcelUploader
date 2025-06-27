// File: res.js
"use strict";

document.addEventListener('DOMContentLoaded', function() {
    const urlParams = new URLSearchParams(window.location.search);
    const templateCategory = urlParams.get('templateCategory'); // Changed from templateId
    const templateCategoryDisplay = templateCategory || 'Unnamed Category';

    const templateNameElement = document.getElementById('templateName');
    if (templateNameElement) {
        templateNameElement.textContent = 'Template Category: ' + decodeURIComponent(templateCategoryDisplay);
    }

    if (templateCategory && templateCategory !== 'null') {
        console.log('Initializing preview for templateCategory:', templateCategory);
        loadPreviewData(templateCategory); // Pass category
    } else {
        console.error('Invalid or missing templateCategory:', templateCategory);
        showMessage('error', 'No valid template category selected.');
    }
});

function loadPreviewData(templateCategory, sheetName = null) { // Changed to templateCategory
    const previewContent = document.getElementById('previewContent');
    if (!previewContent) {
        console.error('Preview content element not found');
        return;
    }

    let url = `/ExcelUploader/preview?templateCategory=${encodeURIComponent(templateCategory)}`; // Pass category
    if (sheetName) {
        url += `&sheetName=${encodeURIComponent(sheetName)}`;
    }

    console.log('Fetching preview data from:', url);

    fetch(url)
        .then(response => response.text().then(text => ({ status: response.status, ok: response.ok, text })))
        .then(({ status, ok, text }) => {
            if (!ok) {
                console.error(`HTTP ${status}: ${text || 'No response body'}`);
                showMessage('error', `Server error: ${text || 'Unable to fetch data.'}`);
                return;
            }
            if (!text.trim()) {
                console.error('Empty response received');
                showMessage('error', 'Server returned empty response.');
                return;
            }
            let data;
            try {
                data = JSON.parse(text);
            } catch (e) {
                console.error('Failed to parse JSON response:', text);
                showMessage('error', 'Invalid server response: ' + text);
                return;
            }
            console.log('Received data:', data);
            if (data.status === 'success') {
                renderPreview(data.sheetNames, data.data, templateCategory); // Pass category to renderPreview
            } else {
                console.error('Server error response:', data);
                showMessage('error', data.message || 'No data available for this template category.');
            }
        })
        .catch(error => {
            console.error('Error loading preview:', error.message);
            showMessage('error', `Error loading preview data: ${error.message}`);
        });
}

function renderPreview(sheetNames, sheetData, templateCategory) {
    const sheetSelectorContainer = document.getElementById('sheetSelectorContainer');
    const previewContent = document.getElementById('previewContent');
    const paginationContainer = document.getElementById('paginationContainer');

    if (!sheetSelectorContainer || !previewContent || !paginationContainer) {
        console.error('One or more layout containers not found');
        return;
    }

    console.log('Rendering preview for sheet:', sheetData ? sheetData.SheetName : 'null');
    sheetSelectorContainer.innerHTML = '';
    previewContent.innerHTML = '';
    paginationContainer.innerHTML = '';


    // Create sheet selection dropdown
    if (sheetNames && sheetNames.length > 1) {
        console.log('Creating dropdown for sheets:', sheetNames);
        const selectDiv = document.createElement('div');
        selectDiv.className = 'mb-2 text-left';
        const sheetSelect = document.createElement('select');
        sheetSelect.id = 'sheetSelect';
        sheetSelect.className = 'form-select';

        sheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            if (sheetData && name === sheetData.SheetName) {
                option.selected = true;
            }
            sheetSelect.appendChild(option);
        });

        sheetSelect.addEventListener('change', function() {
            if (templateCategory && sheetSelect.value) {
                console.log('Sheet changed to:', sheetSelect.value);
                loadPreviewData(templateCategory, sheetSelect.value);
            } else {
                previewContent.innerHTML = '<p>Please select a sheet.</p>';
            }
        });

        selectDiv.appendChild(sheetSelect);
        sheetSelectorContainer.appendChild(selectDiv);
    }

    // Validate sheet data
    if (!sheetData || !Array.isArray(sheetData.Headers) || !Array.isArray(sheetData.Rows)) {
        console.warn('Invalid sheet data:', sheetData);
        previewContent.innerHTML = '<p>No data available for this sheet.</p>';
        return;
    }

    console.log('Sheet data valid, headers:', sheetData.Headers, 'rows:', sheetData.Rows.length);

    // Pagination settings
    const recordsPerPage = 5;
    let currentPage = 1;
    const totalRecords = sheetData.Rows.length;
    const totalPages = Math.ceil(totalRecords / recordsPerPage);

    // Create scrollable table wrapper
    const tableWrapper = document.createElement('div');
    tableWrapper.className = 'table-responsive';

    // Create table
    const table = document.createElement('table');
    table.className = 'table table-bordered';

    // Create header rows
    const thead = document.createElement('thead');
    sheetData.Headers.forEach(headerRow => {
        const tr = document.createElement('tr');
        headerRow.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header.label || '';
            if (header.colspan > 1) {
                th.setAttribute('colspan', header.colspan);
            }
            if (header.rowspan) {
                th.setAttribute('rowspan', header.rowspan);
            }
            th.style.whiteSpace = 'nowrap';
            tr.appendChild(th);
        });
        thead.appendChild(tr);
    });
    table.appendChild(thead);

    // Create table body container
    const tbody = document.createElement('tbody');
    table.appendChild(tbody);

    const paginationDiv = document.createElement('div');
    paginationDiv.className = 'd-flex justify-content-center';

    // Function to render table rows for the current page
    function renderTableRows(page) {
        tbody.innerHTML = ''; // Clear existing rows
        const start = (page - 1) * recordsPerPage;
        const end = Math.min(start + recordsPerPage, totalRecords);
        const paginatedRows = sheetData.Rows.slice(start, end);

        if (paginatedRows.length === 0) {
            console.log('No rows to display; rendering empty table');
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = sheetData.Headers.length > 2 ? sheetData.Headers[2].length : sheetData.Headers[0].length; 
            td.textContent = 'No rows available';
            td.style.textAlign = 'center';
            tr.appendChild(td);
            tbody.appendChild(tr);
        } else {
            paginatedRows.forEach((row, rowIndex) => {
                console.log('Rendering row:', rowIndex + start, row);
                const tr = document.createElement('tr');
                row.forEach(value => {
                    const td = document.createElement('td');
                    td.textContent = value !== undefined && value !== null ? String(value) : '';
                    td.setAttribute('title', value !== undefined && value !== null ? String(value) : ''); // Add title attribute
                    // REMOVED: td.style.whiteSpace = 'nowrap'; // Rely on CSS for truncation and wrapping
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }
    }

    // Function to render pagination controls
    function renderPagination() {
        paginationDiv.innerHTML = ''; // Clear existing pagination

        if (totalPages <= 1) return; // No pagination needed for single page

        const ul = document.createElement('ul');
        ul.className = 'pagination';

        // Previous button
        const prevLi = document.createElement('li');
        prevLi.className = `page-item ${currentPage === 1 ? 'disabled' : ''}`;
        const prevLink = document.createElement('a');
        prevLink.className = 'page-link';
        prevLink.href = '#';
        prevLink.textContent = 'Previous';
        prevLink.addEventListener('click', (e) => {
            e.preventDefault();
            if (currentPage > 1) {
                currentPage--;
                renderTableRows(currentPage);
                renderPagination();
            }
        });
        prevLi.appendChild(prevLink);
        ul.appendChild(prevLi);

        // Page numbers
        const maxPagesToShow = 5;
        let startPage = Math.max(1, currentPage - Math.floor(maxPagesToShow / 2));
        let endPage = Math.min(totalPages, startPage + maxPagesToShow - 1);

        if (endPage - startPage + 1 < maxPagesToShow) {
            startPage = Math.max(1, endPage - maxPagesToShow + 1);
        }

        if (startPage > 1) {
            const firstLi = document.createElement('li');
            firstLi.className = 'page-item';
            const firstLink = document.createElement('a');
            firstLink.className = 'page-link';
            firstLink.href = '#';
            firstLink.textContent = '1';
            firstLink.addEventListener('click', (e) => {
                e.preventDefault();
                currentPage = 1;
                renderTableRows(currentPage);
                renderPagination();
            });
            firstLi.appendChild(firstLink);
            ul.appendChild(firstLi);

            if (startPage > 2) {
                const ellipsisLi = document.createElement('li');
                ellipsisLi.className = 'page-item disabled';
                const ellipsisSpan = document.createElement('span');
                ellipsisSpan.className = 'page-link';
                ellipsisSpan.textContent = '...';
                ellipsisLi.appendChild(ellipsisSpan);
                ul.appendChild(ellipsisLi);
            }
        }

        for (let i = startPage; i <= endPage; i++) {
            const li = document.createElement('li');
            li.className = `page-item ${i === currentPage ? 'active' : ''}`;
            const link = document.createElement('a');
            link.className = 'page-link';
            link.href = '#';
            link.textContent = i;
            link.addEventListener('click', (e) => {
                e.preventDefault();
                currentPage = i;
                renderTableRows(currentPage);
                renderPagination();
            });
            li.appendChild(link);
            ul.appendChild(li);
        }

        if (endPage < totalPages) {
            if (endPage < totalPages - 1) {
                const ellipsisLi = document.createElement('li');
                ellipsisLi.className = 'page-item disabled';
                const ellipsisSpan = document.createElement('span');
                ellipsisSpan.className = 'page-link';
                ellipsisSpan.textContent = '...';
                ellipsisLi.appendChild(ellipsisSpan);
                ul.appendChild(ellipsisLi);
            }

            const lastLi = document.createElement('li');
            lastLi.className = 'page-item';
            const lastLink = document.createElement('a');
            lastLink.className = 'page-link';
            lastLink.href = '#';
            lastLink.textContent = totalPages;
            lastLink.addEventListener('click', (e) => {
                e.preventDefault();
                currentPage = totalPages;
                renderTableRows(currentPage);
                renderPagination();
            });
            lastLi.appendChild(lastLink);
            ul.appendChild(lastLi);
        }

        // Next button
        const nextLi = document.createElement('li');
        nextLi.className = `page-item ${currentPage === totalPages ? 'disabled' : ''}`;
        const nextLink = document.createElement('a');
        nextLink.className = 'page-link';
        nextLink.href = '#';
        nextLink.textContent = 'Next';
        nextLink.addEventListener('click', (e) => {
            e.preventDefault();
            if (currentPage < totalPages) {
                currentPage++;
                renderTableRows(currentPage);
                renderPagination();
            }
        });
        nextLi.appendChild(nextLink);
        ul.appendChild(nextLi);

        paginationDiv.appendChild(ul);
    }

    // Initial render
    renderTableRows(currentPage);
    tableWrapper.appendChild(table);
    previewContent.appendChild(tableWrapper);

    renderPagination();
    paginationContainer.appendChild(paginationDiv);
}

function goBack() {
    window.location.href = '/ExcelUploader/index.html';
}

function showMessage(type, message) {
    const modal = document.getElementById('modal');
    const modalContent = document.getElementById('modal-content');
    if (!modal || !modalContent) {
        console.error('Modal elements not found');
        return;
    }

    modalContent.innerHTML = '';
    const h5 = document.createElement('h5');
    h5.textContent = message;

    const button = document.createElement('button');
    button.className = type === 'error' ? 'btn btn-secondary' : 'btn btn-success';
    button.textContent = 'OK';
    button.addEventListener('click', closeModal);

    modalContent.appendChild(h5);
    modalContent.appendChild(button);
    modal.style.display = 'flex';
}

function closeModal() {
    const modal = document.getElementById('modal');
    const modalContent = document.getElementById('modal-content');
    if (modal && modalContent) {
        modal.style.display = 'none';
        modalContent.innerHTML = '';
    }
}