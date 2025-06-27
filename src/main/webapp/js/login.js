document.addEventListener('DOMContentLoaded', function() {
	const loginForm = document.getElementById('loginForm');
	const usernameOrEmailInput = document.getElementById('usernameOrEmail');
	const passwordInput = document.getElementById('password');
	const errorDiv = document.getElementById('error');
	const messageDiv = document.getElementById('message'); // Get message div for registration success

	// Store original placeholders to revert to them
	const originalUsernameOrEmailPlaceholder = usernameOrEmailInput.placeholder;
	const originalPasswordPlaceholder = passwordInput.placeholder;

	let errorTimeoutIds = {}; // Store timeout IDs per input for clearing error messages

	function applyShakeAndErrorText(inputElement, errorMessage, originalPlaceholder) {
		// Clear any existing timeout for this input
		if (errorTimeoutIds[inputElement.id]) {
			clearTimeout(errorTimeoutIds[inputElement.id]);
		}

		inputElement.classList.add('shake-red-neon');
		inputElement.setAttribute('placeholder', errorMessage); // Change placeholder text

		// Set a timeout to revert
		errorTimeoutIds[inputElement.id] = setTimeout(() => {
			inputElement.classList.remove('shake-red-neon');
			inputElement.setAttribute('placeholder', originalPlaceholder); // Revert placeholder text
			// Clear the general error message div only if it matches this specific error (prevents clearing server errors prematurely)
			if (errorDiv.textContent === errorMessage) {
				errorDiv.textContent = '';
			}
		}, 3000); // Clear after 3 seconds
	}

	loginForm.addEventListener('submit', function(event) {
		let isValid = true;
		errorDiv.textContent = ''; // Clear general error message on new attempt

		// Immediately remove shake classes from previous attempts
		usernameOrEmailInput.classList.remove('shake-red-neon');
		passwordInput.classList.remove('shake-red-neon');

		// Client-side validation
		if (usernameOrEmailInput.value.trim() === '') {
			applyShakeAndErrorText(usernameOrEmailInput, 'Please insert username or email.', originalUsernameOrEmailPlaceholder);
			isValid = false;
		}

		if (passwordInput.value.trim() === '') {
			applyShakeAndErrorText(passwordInput, 'Please insert password.', originalPasswordPlaceholder);
			isValid = false;
		}

		if (!isValid) {
			event.preventDefault(); // Prevent form submission if client-side validation fails
		}
		// If isValid is true, the form will submit normally to the servlet for server-side validation.
	});

	// Handle server-side errors from URL parameters (existing logic, remains in errorDiv)
	const urlParams = new URLSearchParams(window.location.search);
	const serverError = urlParams.get('error');
	const registeredSuccess = urlParams.get('registered');

	if (serverError === '1') {
		errorDiv.textContent = 'Incorrect username/email or password';
		// Clear this server-side message after 3 seconds
		setTimeout(() => errorDiv.textContent = '', 3000);
	} else if (registeredSuccess === '1') {
		messageDiv.textContent = 'Registration successful, please log in';
		// Clear success message too
		setTimeout(() => messageDiv.textContent = '', 3000);
	}
});