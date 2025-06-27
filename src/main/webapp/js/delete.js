// File: delete.js
"use strict";

document.addEventListener('DOMContentLoaded', function() {
    // --- Initial Load: Populate Template Categories ---
    loadTemplateCategories().then(() => {
        const urlParams = new URLSearchParams(window.location.search);
        const autoSelectCategory = urlParams.get('templateCategory');
        const categorySelect = document.getElementById('templateCategorySelect');

        if (categorySelect && autoSelectCategory && autoSelectCategory !== 'null') {
            const optionExists = Array.from(categorySelect.options).some(option => option.value === autoSelectCategory);
            if (optionExists) {
                categorySelect.value = autoSelectCategory;
                loadTemplatesByCategory(autoSelectCategory);
            } else {
                console.warn(`Template category "${autoSelectCategory}" from URL not found in dropdown.`);
                updateTemplateCounters(); 
            }
        } else {
            updateTemplateCounters();
        }
    }).catch(error => {
        console.error('Initial loading of template categories failed:', error);
        showMessage('error', 'Initial setup failed: ' + error.message);
        updateTemplateCounters();
    });


    // --- Event Listener: Template Category Selection ---
    const categorySelect = document.getElementById('templateCategorySelect');
    if (categorySelect) {
        categorySelect.addEventListener('change', function() {
            const selectedCategory = this.value;

            // Clear shake-red-neon from category select and template list when category changes
            this.classList.remove('shake-red-neon');
            // Revert default option text if it was changed
            if (this.options[0] && this.options[0].getAttribute('data-original-text')) {
                this.options[0].textContent = this.options[0].getAttribute('data-original-text');
                this.options[0].style.color = ''; // Reset color
            }
            this.style.color = ''; // Reset select box text color

            const templateList = document.getElementById('templateList');
            if (templateList) {
                templateList.classList.remove('shake-red-neon');
                // Clear any temporary error message in templateList
                const tempMessage = templateList.querySelector('.temp-error-message');
                if (tempMessage) {
                    tempMessage.remove();
                    templateList.innerHTML = '<p>Select a category to view templates.</p>'; // Revert to original message
                }
            }
            
            if (selectedCategory) {
                loadTemplatesByCategory(selectedCategory);
            } else {
                document.getElementById('templateList').innerHTML = '<p>Select a category to view templates.</p>';
                document.getElementById('selectAllTemplates').checked = false;
                updateTemplateCounters();
            }
        });
    }

    // --- Event Listener: "Select All" Checkbox ---
    const selectAllCheckbox = document.getElementById('selectAllTemplates');
    if (selectAllCheckbox) {
        selectAllCheckbox.addEventListener('change', function() {
            const checkboxes = document.querySelectorAll('#templateList input[type="checkbox"].template-checkbox');
            checkboxes.forEach(cb => {
                cb.checked = this.checked;
            });
            updateTemplateCounters();
        });
    }

    // --- Event Listener: Individual Template Checkbox Changes ---
    const templateListDiv = document.getElementById('templateList');
    if (templateListDiv) {
        templateListDiv.addEventListener('change', function(event) {
            if (event.target.type === 'checkbox' && event.target.classList.contains('template-checkbox')) {
                // Clear shake-red-neon from template list when a checkbox is interacted with
                const templateList = document.getElementById('templateList');
                if (templateList) {
                    templateList.classList.remove('shake-red-neon');
                    // Clear any temporary error message in templateList
                    const tempMessage = templateList.querySelector('.temp-error-message');
                    if (tempMessage) {
                        tempMessage.remove();
                    }
                }

                updateTemplateCounters();

                const selectAll = document.getElementById('selectAllTemplates');
                if (selectAll) {
                    const allCheckboxes = document.querySelectorAll('#templateList input[type="checkbox"].template-checkbox');
                    const checkedCheckboxes = document.querySelectorAll('#templateList input[type="checkbox"].template-checkbox:checked');

                    if (allCheckboxes.length > 0 && checkedCheckboxes.length === allCheckboxes.length) {
                        selectAll.checked = true;
                    } else {
                        selectAll.checked = false;
                    }
                }
            }
        });
    }
});

/**
 * Fetches distinct template categories from the server and populates the
 * templateCategorySelect dropdown.
 */
function loadTemplateCategories() {
    return new Promise((resolve, reject) => {
        const categorySelect = document.getElementById('templateCategorySelect');
        if (!categorySelect) {
            console.error('Template category select element not found');
            return reject(new Error('Template category select element not found'));
        }

        if (categorySelect.options[0] && !categorySelect.options[0].getAttribute('data-original-text')) {
            categorySelect.options[0].setAttribute('data-original-text', categorySelect.options[0].textContent);
        }

        categorySelect.classList.remove('shake-red-neon');
        // Revert default option text/color explicitly in case it was changed
        if (categorySelect.options[0] && categorySelect.options[0].getAttribute('data-original-text')) {
            categorySelect.options[0].textContent = categorySelect.options[0].getAttribute('data-original-text');
            categorySelect.options[0].style.color = '';
        }
        categorySelect.style.color = '';


        // Keep the first option if it's the placeholder, remove others
        while (categorySelect.options.length > 1) { 
            categorySelect.remove(1);
        }

        fetch('/ExcelUploader/templates')
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

                data.forEach(category => {
                    if (category.TemplateCategory && category.TemplateCategory !== 'null') {
                        const option = document.createElement('option');
                        option.value = category.TemplateCategory;
                        option.textContent = category.TemplateCategory;
                        categorySelect.appendChild(option);
                    }
                });
                resolve();
            })
            .catch(error => {
                console.error('Error loading template categories:', error);
                showMessage('error', 'Error loading template categories: ' + error.message);
                reject(error);
            });
    });
}

/**
 * Fetches all individual uploaded files/sheets for a specific category and populates the templateList div
 * with checkboxes. Updates counters after loading.
 * Each checkbox now represents a ResponseID (individual sheet from an uploaded file).
 * @param {string} templateCategory The category name to load templates for.
 */
function loadTemplatesByCategory(templateCategory) {
    const templateList = document.getElementById('templateList');
    if (!templateList) {
        console.error('Template list element not found');
        return;
    }

    templateList.classList.remove('shake-red-neon');
    const tempMessage = templateList.querySelector('.temp-error-message');
    if(tempMessage) tempMessage.remove();

    templateList.innerHTML = '<p>Loading templates...</p>';
    const selectAllCheckbox = document.getElementById('selectAllTemplates');
    if (selectAllCheckbox) {
        selectAllCheckbox.checked = false;
    }

    // Now fetches responses (individual files/sheets) by template category
    fetch(`/ExcelUploader/templatesByCategory?templateCategory=${encodeURIComponent(templateCategory)}`)
        .then(response => {
            if (!response.ok) {
                return response.json().then(errorData => {
                    throw new Error(errorData.message || 'Network response was not ok');
                });
            }
            return response.json();
        })
        .then(data => {
            templateList.innerHTML = ''; // Clear loading message.

            if (!Array.isArray(data)) {
                throw new Error('Invalid template data format for category');
            }

            if (data.length === 0) {
                templateList.innerHTML = '<p>No templates found for this category.</p>';
            } else {
                data.forEach(responseItem => {
                    const responseID = responseItem.ResponseID;
                    const originalFileName = responseItem.OriginalFileName || 'Unnamed File';
                    const sheetName = responseItem.SheetName || 'Unnamed Sheet';
                    
                    // Display name: "OriginalFileName (SheetName)"
                    const displayName = `${originalFileName} (Sheet: ${sheetName})`;

                    const div = document.createElement('div');
                    div.className = 'form-check';
                    div.style.marginBottom = '8px';

                    const checkbox = document.createElement('input');
                    checkbox.type = 'checkbox';
                    checkbox.className = 'form-check-input template-checkbox';
                    checkbox.value = responseID; // Checkbox value is ResponseID
                    checkbox.id = `response-${responseID}`;
                    checkbox.setAttribute('data-template-name', displayName); // Use display name for modal

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
            const selectAll = document.getElementById('selectAllTemplates');
            if (selectAll) {
                const allCheckboxes = document.querySelectorAll('#templateList input[type="checkbox"].template-checkbox');
                if (allCheckboxes.length > 0 && allCheckboxes.length === document.querySelectorAll('#templateList input[type="checkbox"].template-checkbox:checked').length) {
                    selectAll.checked = true;
                } else {
                    selectAll.checked = false;
                }
            }
        })
        .catch(error => {
            console.error('Error loading templates by category:', error);
            templateList.innerHTML = `<p style="color: red;">Error loading templates: ${error.message}</p>`;
            updateTemplateCounters();
        });
}

/**
 * Updates the display counters for total and selected templates.
 */
function updateTemplateCounters() {
    const allCheckboxes = document.querySelectorAll('#templateList input[type="checkbox"].template-checkbox');
    const selectedCheckboxes = document.querySelectorAll('#templateList input[type="checkbox"].template-checkbox:checked');

    const totalCountElement = document.getElementById('totalCount');
    const selectedCountElement = document.getElementById('selectedCount');

    if (totalCountElement) {
        totalCountElement.textContent = allCheckboxes.length;
    }
    if (selectedCountElement) {
        selectedCountElement.textContent = selectedCheckboxes.length;
    }
}

/**
 * Initiates the template deletion process by gathering selected templates
 * and showing a confirmation warning.
 */
function deleteTemplates() {
    const templateList = document.getElementById('templateList');
    const categorySelect = document.getElementById('templateCategorySelect');

    // Reset any previous error states from both elements before validation
    if (categorySelect) {
        categorySelect.classList.remove('shake-red-neon');
        if (categorySelect.options[0] && categorySelect.options[0].getAttribute('data-original-text')) {
            categorySelect.options[0].textContent = categorySelect.options[0].getAttribute('data-original-text');
            categorySelect.options[0].style.color = '';
        }
        categorySelect.style.color = '';
    }
    if (templateList) {
        templateList.classList.remove('shake-red-neon');
        const tempMessage = templateList.querySelector('.temp-error-message');
        if(tempMessage) tempMessage.remove();
    }

    // Validation Step 1: Check if a category is selected
    if (!categorySelect || !categorySelect.value || categorySelect.value === 'null' || categorySelect.value === '') {
        if (categorySelect) {
            categorySelect.classList.add('shake-red-neon');
            if (categorySelect.options[0]) {
                categorySelect.options[0].textContent = 'Please select a category.';
                categorySelect.options[0].style.color = '#ff4d4d';
            }
            categorySelect.style.color = '#ff4d4d';
            
            setTimeout(() => {
                categorySelect.classList.remove('shake-red-neon');
                if (categorySelect.options[0] && categorySelect.options[0].getAttribute('data-original-text')) {
                    categorySelect.options[0].textContent = categorySelect.options[0].getAttribute('data-original-text');
                    categorySelect.options[0].style.color = '';
                }
                categorySelect.style.color = '';
            }, 3000);
        }
        return;
    }

    // Validation Step 2: Check if any templates (responses) are selected within the chosen category
    const checkboxes = templateList.querySelectorAll('input[type="checkbox"].template-checkbox:checked');
    if (checkboxes.length === 0) {
        if (templateList) {
            templateList.classList.add('shake-red-neon');
            let errorMessageDiv = templateList.querySelector('.temp-error-message');
            if (!errorMessageDiv) {
                errorMessageDiv = document.createElement('p');
                errorMessageDiv.className = 'temp-error-message';
                errorMessageDiv.style.color = '#ff4d4d';
                templateList.prepend(errorMessageDiv);
            }
            errorMessageDiv.textContent = 'Please select at least one sheet.';
            
            setTimeout(() => {
                templateList.classList.remove('shake-red-neon');
                if(errorMessageDiv) errorMessageDiv.remove();
            }, 3000);
        }
        return;
    }

    // If both category and templates are selected, proceed to confirmation
    // Ensure IDs are numbers for JSON serialization
    const responseIdsToDelete = Array.from(checkboxes).map(cb => parseInt(cb.value, 10)); 
    const templateNamesForDisplay = Array.from(checkboxes).map(cb => cb.getAttribute('data-template-name') || 'Unnamed Template/Sheet');
    
    showDeleteWarning(responseIdsToDelete, templateNamesForDisplay);
}

/**
 * Displays a confirmation modal with the list of templates (responses) to be deleted.
 * @param {Array<number>} responseIds An array of ResponseIDs to delete.
 * @param {Array<string>} displayNames An array of display names corresponding to the ResponseIDs.
 */
function showDeleteWarning(responseIds, displayNames) {
    const modal = document.getElementById('modal');
    const modalContent = document.getElementById('modal-content');
    if (!modal || !modalContent) {
        console.error('Modal elements not found');
        return;
    }

    modalContent.innerHTML = '';
    const h5 = document.createElement('h5');
    h5.className = 'mb-3';
    h5.textContent = 'Delete the following uploaded files/sheets?'; // More specific message
    
    // NEW: Create a scrollable div for the list of items
    const scrollableDiv = document.createElement('div');
    scrollableDiv.className = 'scrollable-list'; // Apply new CSS class for scrollability
    scrollableDiv.style.maxHeight = '150px'; // Limit height for scrollability
    scrollableDiv.style.overflowY = 'auto'; // Enable vertical scrolling
    scrollableDiv.style.paddingRight = '10px'; // Add some padding for scrollbar

    const ul = document.createElement('ul');
    ul.className = 'list-unstyled text-start mb-0'; 
    displayNames.forEach(name => {
        const li = document.createElement('li');
        li.textContent = '• ' + name;
        ul.appendChild(li);
    });
    scrollableDiv.appendChild(ul); 

    const buttonContainer = document.createElement('div');
    buttonContainer.className = 'd-flex justify-content-center gap-3 mt-3'; // Added mt-3 for spacing

    const yesButton = document.createElement('button');
    yesButton.className = 'btn-neon btn-red'; 
    yesButton.textContent = 'Yes, Delete';
    yesButton.addEventListener('click', function() {
        confirmDelete(responseIds); // Only pass IDs to backend, displayNames are for frontend UI
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

/**
 * Confirms and executes the deletion of selected responses.
 * Now sends a single batch DELETE request.
 * @param {Array<number>} responseIds An array of ResponseIDs to delete.
 */
function confirmDelete(responseIds) { // Removed displayNames as it's not needed by backend
    const modalContent = document.getElementById('modal-content');
    if (!modalContent) {
        console.error('Modal content element not found');
        return;
    }

    // CHANGED: Show Bootstrap spinner and message for deletion
    modalContent.innerHTML = `
        <h5>Deleting... <span class="spinner-border spinner-border-sm text-info" role="status" aria-hidden="true"></span></h5>
        <p>Please wait. Do not close this window.</p>
    `;
    // No need for a custom 'loading' class if using Bootstrap's spinner directly

    const categorySelect = document.getElementById('templateCategorySelect');
    const currentCategory = categorySelect ? categorySelect.value : null;

    // CHANGED: Send all responseIds in the request body as JSON
    fetch('/ExcelUploader/delete', { 
        method: 'DELETE',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ responseIds: responseIds }) // Send array of IDs
    })
    .then(response => {
        // No need to remove loading class in .then or .catch, showMessage will override
        if (!response.ok) {
            return response.json().then(errorData => {
                throw new Error(errorData.message || `Delete failed: HTTP ${response.status}`);
            });
        }
        return response.json();
    })
    .then(data => {
        let message;
        if (data.status === 'success') {
            message = data.message || `Successfully deleted ${data.deletedCount} files/sheets.`;
        } else {
            message = data.message || `Failed to delete files/sheets.`;
        }

        // Changed showMessage to not auto-close for delete confirmation
        showMessage(data.status, message, null, false); 

        // MODIFIED: After showing the message, refresh categories and templates
        loadTemplateCategories().then(() => {
            const categorySelectRefreshed = document.getElementById('templateCategorySelect');
            const templateList = document.getElementById('templateList');
            
            // Check if the previously selected category still exists in the refreshed dropdown
            const categoryStillExists = currentCategory && Array.from(categorySelectRefreshed.options).some(option => option.value === currentCategory);

            if (categoryStillExists) {
                // If the category still exists, re-select it and load its templates
                categorySelectRefreshed.value = currentCategory;
                loadTemplatesByCategory(currentCategory);
            } else {
                // The previously selected category was likely deleted, or none was selected initially.
                // Reset the dropdown to default and clear the template list.
                categorySelectRefreshed.value = ''; // Set dropdown to "Select a Category"
                templateList.innerHTML = '<p>Select a category to view templates.</p>';
                document.getElementById('selectAllTemplates').checked = false; // Uncheck "Select All"
                updateTemplateCounters(); // Reset counters to 0/0
            }
        }).catch(error => {
            console.error("Error refreshing categories after delete:", error);
            // Changed showMessage to not auto-close for errors
            showMessage('error', 'Error refreshing category list after deletion: ' + error.message, null, false); 
            // Fallback: Ensure UI is cleared even if refresh fails
            const categorySelectRefreshed = document.getElementById('templateCategorySelect');
            const templateList = document.getElementById('templateList');
            if (categorySelectRefreshed) categorySelectRefreshed.value = '';
            if (templateList) templateList.innerHTML = '<p>Select a category to view templates.</p>';
            document.getElementById('selectAllTemplates').checked = false;
            updateTemplateCounters();
        });
    })
    .catch(error => {
        console.error('Error during batch deletion:', error);
        // Changed showMessage to not auto-close for errors
        showMessage('error', 'An unexpected error occurred during deletion: ' + error.message, null, false); 
        updateTemplateCounters();
    });
}

/**
 * Displays a custom message in the modal window.
 */
// MODIFIED: Added autoClose parameter
function showMessage(type, message, context = null, autoClose = true) {
    const modal = document.getElementById('modal');
    const modalContent = document.getElementById('modal-content');
    if (!modal || !modalContent) {
        console.error('Modal elements not found');
        return;
    }

    // Ensure no loading spinner is showing if this message is for completion
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

    // Only auto-close if autoClose is true
    if (autoClose) {
        setTimeout(() => {
            closeModal();
        }, 3000);
    }
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