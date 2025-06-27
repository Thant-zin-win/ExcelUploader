// File: manage.js
"use strict";

document.addEventListener('DOMContentLoaded', function() {
    const existingTemplateSelect = document.getElementById('existingTemplateSelect');
    const newCategoryNameInput = document.getElementById('newCategoryName');
    const saveCategoryBtn = document.getElementById('saveCategoryBtn');

    let selectedTemplateId = null; // Store the TemplateID of the currently selected template

    // --- Initial Load: Populate Template Categories ---
    loadExistingTemplates().then(() => {
        // Auto-select template category if provided in URL
        const urlParams = new URLSearchParams(window.location.search);
        const autoSelectCategory = urlParams.get('templateCategory');
        if (autoSelectCategory && autoSelectCategory !== 'null') {
            const optionExists = Array.from(existingTemplateSelect.options).some(option => option.value === autoSelectCategory);
            if (optionExists) {
                existingTemplateSelect.value = autoSelectCategory;
                // Manually trigger change to populate newCategoryNameInput
                const event = new Event('change');
                existingTemplateSelect.dispatchEvent(event);
            } else {
                console.warn(`Template category "${autoSelectCategory}" from URL not found in dropdown.`);
            }
        }
    }).catch(error => {
        console.error('Initial loading of template categories failed:', error);
        showMessage('error', 'Initial setup failed: ' + error.message);
    });


    // --- Event Listener: Existing Template Category Selection ---
    if (existingTemplateSelect) {
        existingTemplateSelect.addEventListener('change', function() {
            // Clear any previous error highlighting
            this.classList.remove('shake-red-neon');
            this.style.color = '';
            
            const selectedOption = this.options[this.selectedIndex];
            // Ensure selectedTemplateId is updated to the data-template-id, not just the value
            selectedTemplateId = selectedOption.getAttribute('data-template-id'); 
            const currentCategoryName = selectedOption.value;

            newCategoryNameInput.value = currentCategoryName !== "" ? currentCategoryName : "";

            // Clear shake on newCategoryName input if user selects a different category
            newCategoryNameInput.classList.remove('shake-red-neon');
            newCategoryNameInput.style.color = '';
            newCategoryNameInput.setAttribute('placeholder', 'Enter new category name');
        });
    }

    // --- Event Listener: Save Changes Button ---
    if (saveCategoryBtn) {
        saveCategoryBtn.addEventListener('click', function() {
            // Reset error highlights
            existingTemplateSelect.classList.remove('shake-red-neon');
            existingTemplateSelect.style.color = '';
            newCategoryNameInput.classList.remove('shake-red-neon');
            newCategoryNameInput.style.color = '';
            newCategoryNameInput.setAttribute('placeholder', 'Enter new category name');

            const selectedCategoryValue = existingTemplateSelect.value;
            const newName = newCategoryNameInput.value.trim();

            if (!selectedTemplateId || selectedCategoryValue === "") {
                existingTemplateSelect.classList.add('shake-red-neon');
                existingTemplateSelect.style.color = '#ff4d4d';
                showMessage('error', 'Please select a template category to edit.');
                return;
            }

            if (newName === "") {
                newCategoryNameInput.classList.add('shake-red-neon');
                newCategoryNameInput.style.color = '#ff4d4d';
                newCategoryNameInput.setAttribute('placeholder', 'New category name cannot be empty.');
                showMessage('error', 'New category name cannot be empty.');
                return;
            }

            // Confirm with user before saving
            showConfirmSaveDialog(selectedCategoryValue, newName);
        });
    }
});

/**
 * Fetches existing template categories from the server and populates the dropdown.
 * Also stores the TemplateID in a data attribute for later use.
 * Returns a Promise to allow chaining.
 */
function loadExistingTemplates() {
    return new Promise((resolve, reject) => {
        const existingTemplateSelect = document.getElementById('existingTemplateSelect');
        if (!existingTemplateSelect) {
            console.error('Existing template select element not found');
            return reject(new Error('Existing template select element not found'));
        }

        // Clear existing options, keeping the first placeholder
        while (existingTemplateSelect.options.length > 1) {
            existingTemplateSelect.remove(1);
        }

        fetch('/ExcelUploader/templates') // This servlet already returns TemplateCategory and SampleTemplateID (which is TemplateID)
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
                    throw new Error('Invalid template categories data format');
                }

                data.forEach(template => {
                    if (template.TemplateCategory && template.TemplateCategory !== 'null') {
                        const option = document.createElement('option');
                        option.value = template.TemplateCategory;
                        option.textContent = template.TemplateCategory;
                        option.setAttribute('data-template-id', template.SampleTemplateID); // Store TemplateID
                        existingTemplateSelect.appendChild(option);
                    }
                });
                resolve(); // Resolve the promise once categories are loaded
            })
            .catch(error => {
                console.error('Error loading template categories for management:', error);
                showMessage('error', 'Error loading templates: ' + error.message);
                reject(error); // Reject the promise on error
            });
    });
}

/**
 * Displays a confirmation modal before saving changes.
 * @param {string} oldName The current name of the template category.
 * @param {string} newName The proposed new name for the template category.
 */
function showConfirmSaveDialog(oldName, newName) {
    const modal = document.getElementById('modal');
    const modalContent = document.getElementById('modal-content');
    if (!modal || !modalContent) {
        console.error('Modal elements not found');
        return;
    }

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
        closeModal(); // Close confirmation modal immediately
        saveTemplateCategory(); // Proceed to save
    });

    const cancelButton = document.createElement('button');
    cancelButton.className = 'btn-neon btn-blue';
    cancelButton.textContent = 'Cancel';
    cancelButton.addEventListener('click', closeModal);

    modalContent.appendChild(h5);
    buttonContainer.appendChild(yesButton);
    buttonContainer.appendChild(cancelButton);
    modalContent.appendChild(buttonContainer);
    modal.style.display = 'flex';
}

/**
 * Sends the update request to the server.
 */
function saveTemplateCategory() {
    const existingTemplateSelect = document.getElementById('existingTemplateSelect');
    const newCategoryNameInput = document.getElementById('newCategoryName');

    // Make sure we get the currently active selectedTemplateId, not relying on a global variable
    const selectedOption = existingTemplateSelect.options[existingTemplateSelect.selectedIndex];
    const templateId = selectedOption.getAttribute('data-template-id');
    const newCategoryName = newCategoryNameInput.value.trim();

    if (!templateId || !newCategoryName) {
        showMessage('error', 'Missing template ID or new category name.');
        return;
    }

    // Show "Saving..." message
    const modalContent = document.getElementById('modal-content');
    if (modalContent) {
        modalContent.innerHTML = '<h5>Saving changes...</h5><p>Please wait. Do not close this window.</p>';
        document.getElementById('modal').style.display = 'flex';
    }


    fetch('/ExcelUploader/updateTemplateCategory', {
        method: 'PUT', // Use PUT for updates
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            templateId: templateId,
            newCategoryName: newCategoryName
        })
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(errorData => {
                throw new Error(errorData.message || 'Failed to update template category.');
            });
        }
        return response.json();
    })
    .then(data => {
        if (data.status === 'success') {
            showMessage('success', data.message || 'Template category updated successfully!');
            // Reload templates to reflect the change and select the updated one
            loadExistingTemplates().then(() => {
                // After loading, attempt to re-select the updated category
                existingTemplateSelect.value = newCategoryName;
                newCategoryNameInput.value = ''; // Clear input
                // Manually trigger change to update selectedTemplateId and input field
                const event = new Event('change');
                existingTemplateSelect.dispatchEvent(event);
            });
        } else {
            showMessage('error', data.message || 'Error updating template category.');
        }
    })
    .catch(error => {
        console.error('Error updating template category:', error);
        showMessage('error', error.message || 'An unexpected error occurred during update.');
    });
}

/**
 * Displays a custom message in the modal window.
 */
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
    h5.className = 'mb-3';

    const button = document.createElement('button');
    button.className = type === 'success' ? 'btn-neon btn-green' : 'btn-neon btn-blue';
    button.textContent = 'OK';
    button.addEventListener('click', closeModal);

    modalContent.appendChild(h5);
    modalContent.appendChild(button);
    modal.style.display = 'flex';

    setTimeout(() => {
        closeModal();
    }, 3000); // Auto-close after 3 seconds
}

/**
 * Closes the modal window and clears its content.
 */
function closeModal() {
    const modal = document.getElementById('modal');
    const modalContent = document.getElementById('modal-content');
    if (modal && modalContent) {
        modal.style.display = 'none';
        modalContent.innerHTML = '';
    }
}