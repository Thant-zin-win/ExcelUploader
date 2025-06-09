"use strict";

document.addEventListener('DOMContentLoaded', function() {
    // Load templates on page load
    loadTemplates();

    // Attach event listeners
    var uploadForm = document.getElementById('uploadForm');
    if (uploadForm) {
        uploadForm.addEventListener('submit', initiateUpload);
    } else {
        console.error('Upload form element not found');
    }

    var downloadButton = document.querySelector('.btn-green');
    if (downloadButton) {
        downloadButton.addEventListener('click', downloadTemplateData);
    } else {
        console.error('Download button not found');
    }

    var deleteButton = document.querySelector('.btn-red');
    if (deleteButton) {
        deleteButton.addEventListener('click', deleteTemplate);
    } else {
        console.error('Delete button not found');
    }
});

// Load template options from the server
function loadTemplates() {
    var select = document.getElementById('templateSelect');
    if (!select) {
        console.error('Template select element not found');
        return;
    }

    // Store the currently selected value (if any) to restore it after refresh
    var selectedValue = select.value;

    // Clear the dropdown completely before adding new options
    while (select.firstChild) {
        select.removeChild(select.firstChild);
    }

    // Add the default "Select a Template" option
    var defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Select a Template';
    select.appendChild(defaultOption);

    fetch('/ExcelUploader/templates')
        .then(function(response) {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(function(data) {
            if (!Array.isArray(data)) {
                throw new Error('Invalid template data format');
            }

            // Optional: Deduplicate templates client-side based on TemplateID
            var seenTemplateIDs = new Set();
            data.forEach(function(template) {
                var templateID = template.TemplateID || '';
                if (seenTemplateIDs.has(templateID)) {
                    return; // Skip duplicates
                }
                seenTemplateIDs.add(templateID);

                var option = document.createElement('option');
                option.value = templateID;
                option.textContent = template.TemplateName || 'Unnamed Template';
                option.setAttribute('data-template-name', template.TemplateName || '');
                select.appendChild(option);
            });

            // Restore the previously selected value if it still exists
            if (selectedValue) {
                var optionExists = Array.from(select.options).some(function(option) {
                    return option.value === selectedValue;
                });
                if (optionExists) {
                    select.value = selectedValue;
                }
            }
        })
        .catch(function(error) {
            console.error('Error loading templates:', error);
        });
}

// Handle file upload initiation
function initiateUpload(event) {
    event.preventDefault();
    var fileInput = document.getElementById('file');
    if (!fileInput) {
        console.error('File input element not found');
        showMessage('error', 'File input not found.');
        return;
    }

    var file = fileInput.files[0];
    if (file && file.name.endsWith('.xlsx')) {
        showWarningUpload(file);
    } else {
        showMessage('warning', 'Please select a valid .xlsx file.');
    }
}

// Show upload confirmation modal
function showWarningUpload(file) {
    var modal = document.getElementById('modal');
    var modalContent = document.getElementById('modal-content');
    if (!modal || !modalContent) {
        console.error('Modal elements not found');
        return;
    }

    modalContent.innerHTML = '';
    var h5 = document.createElement('h5');
    h5.textContent = 'Upload file "' + file.name + '"?';

    var yesButton = document.createElement('button');
    yesButton.className = 'btn btn-primary';
    yesButton.textContent = 'Yes';
    yesButton.addEventListener('click', function() {
        confirmUpload(file.name);
    });

    var cancelButton = document.createElement('button');
    cancelButton.className = 'btn btn-secondary';
    cancelButton.textContent = 'Cancel';
    cancelButton.addEventListener('click', closeModal);

    modalContent.appendChild(h5);
    modalContent.appendChild(yesButton);
    modalContent.appendChild(cancelButton);
    modal.style.display = 'flex';
}

// Confirm and process file upload
function confirmUpload(fileName) {
    var modalContent = document.getElementById('modal-content');
    if (!modalContent) {
        console.error('Modal content element not found');
        return;
    }

    modalContent.innerHTML = '<h5>Uploading...</h5><p>Please wait.</p>';

    var formData = new FormData();
    var fileInput = document.getElementById('file');
    if (fileInput && fileInput.files[0]) {
        formData.append('file', fileInput.files[0]);
    } else {
        showMessage('error', 'No file selected for upload.');
        return;
    }

    fetch('/ExcelUploader/upload', {
        method: 'POST',
        body: formData
    })
        .then(function(response) {
            if (!response.ok) {
                throw new Error('Upload failed');
            }
            return response.json();
        })
        .then(function(data) {
            if (data.status === 'success') {
                var form = document.getElementById('uploadForm');
                if (form) form.reset();
                showMessage('success', 'File "' + fileName + '" uploaded successfully.');
            } else {
                throw new Error(data.message || 'Upload failed.');
            }
        })
        .catch(function(error) {
            console.error('Upload error:', error);
            showMessage('error', error.message || 'Error uploading file.');
        });
}

// Handle template download
function downloadTemplateData() {
    var select = document.getElementById('templateSelect');
    if (!select) {
        console.error('Template select element not found');
        showMessage('error', 'Template dropdown not found.');
        return;
    }

    var templateId = select.value;
    if (templateId) {
        window.location.href = '/ExcelUploader/export?templateId=' + templateId;
        // Reload the page after a delay to ensure download initiates
        setTimeout(closeModalAndRefresh, 2000); // 2-second delay
    } else {
        showMessage('warning', 'Please select a template.');
    }
}

// Show delete confirmation modal
function deleteTemplate() {
    var select = document.getElementById('templateSelect');
    if (!select) {
        console.error('Template select element not found');
        showMessage('error', 'Template dropdown not found.');
        return;
    }

    var id = select.value;
    var name = select.selectedOptions[0] ? select.selectedOptions[0].getAttribute('data-template-name') : 'Unknown Template';
    if (id) {
        showDeleteWarning(id, name);
    } else {
        showMessage('warning', 'Please select a template.');
    }
}

// Display delete confirmation
function showDeleteWarning(id, name) {
    var modal = document.getElementById('modal');
    var modalContent = document.getElementById('modal-content');
    if (!modal || !modalContent) {
        console.error('Modal elements not found');
        return;
    }

    modalContent.innerHTML = '';
    var h5 = document.createElement('h5');
    h5.textContent = 'Delete template "' + name + '"?';

    var yesButton = document.createElement('button');
    yesButton.className = 'btn btn-danger';
    yesButton.textContent = 'Yes';
    yesButton.addEventListener('click', function() {
        confirmDelete(id, name);
    });

    var cancelButton = document.createElement('button');
    cancelButton.className = 'btn btn-secondary';
    cancelButton.textContent = 'Cancel';
    cancelButton.addEventListener('click', closeModal);

    modalContent.appendChild(h5);
    modalContent.appendChild(yesButton);
    modalContent.appendChild(cancelButton);
    modal.style.display = 'flex';
}

// Confirm and process template deletion
function confirmDelete(id, name) {
    fetch('/ExcelUploader/delete?templateId=' + id, { method: 'DELETE' })
        .then(function(response) {
            if (!response.ok) {
                throw new Error('Delete failed');
            }
            return response.json();
        })
        .then(function(data) {
            if (data.status === 'success') {
                showMessage('success', 'Template "' + name + '" deleted successfully.');
            } else {
                throw new Error(data.message || 'Delete failed.');
            }
        })
        .catch(function(error) {
            console.error('Delete error:', error);
            showMessage('error', error.message || 'Error deleting template.');
        });
}

// Display messages in modal
function showMessage(type, message) {
    var modal = document.getElementById('modal');
    var modalContent = document.getElementById('modal-content');
    if (!modal || !modalContent) {
        console.error('Modal elements not found');
        return;
    }

    modalContent.innerHTML = '';
    var h5 = document.createElement('h5');
    h5.textContent = message;

    var button = document.createElement(type === 'success' ? 'a' : 'button');
    button.className = type === 'success' ? 'btn btn-success' : 'btn btn-secondary';
    button.textContent = 'OK';

    if (type === 'success') {
        button.href = '#';
        button.addEventListener('click', function(event) {
            event.preventDefault();
            closeModalAndRefresh();
        });
    } else {
        button.addEventListener('click', closeModal);
    }

    modalContent.appendChild(h5);
    modalContent.appendChild(button);
    modal.style.display = 'flex';
}

// Close the modal
function closeModal() {
    var modal = document.getElementById('modal');
    var modalContent = document.getElementById('modal-content');
    if (modal && modalContent) {
        modal.style.display = 'none';
        modalContent.innerHTML = '';
    }
}

// Reload the page
function closeModalAndRefresh() {
    location.reload();
}
