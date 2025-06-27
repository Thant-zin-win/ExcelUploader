document.addEventListener('DOMContentLoaded', function() {
	const registerForm = document.getElementById('registerForm');
	const usernameInput = document.getElementById('username');
	const emailInput = document.getElementById('email');
	const passwordInput = document.getElementById('password');
	const confirmPasswordInput = document.getElementById('confirmPassword');
	const errorDiv = document.getElementById('error');

	// Store original placeholders
	const originalUsernamePlaceholder = usernameInput.placeholder;
	const originalEmailPlaceholder = emailInput.placeholder;
	const originalPasswordPlaceholder = passwordInput.placeholder;
	const originalConfirmPasswordPlaceholder = confirmPasswordInput.placeholder;

	let errorTimeoutIds = {}; // Store timeout IDs per input

	function applyShakeAndErrorText(inputElement, errorMessage, originalPlaceholder) {
		if (errorTimeoutIds[inputElement.id]) {
			clearTimeout(errorTimeoutIds[inputElement.id]);
		}

		inputElement.classList.add('shake-red-neon');
		inputElement.setAttribute('placeholder', errorMessage);

		errorTimeoutIds[inputElement.id] = setTimeout(() => {
			inputElement.classList.remove('shake-red-neon');
			inputElement.setAttribute('placeholder', originalPlaceholder);
		}, 3000);
	}

	registerForm.addEventListener('submit', function(event) {
		let isValid = true;
		let primaryErrorMessage = ''; // To hold the first error message for errorDiv (if needed)

		// Clear general error message and shake classes from previous attempts
		errorDiv.textContent = '';
		[usernameInput, emailInput, passwordInput, confirmPasswordInput].forEach(input => {
			input.classList.remove('shake-red-neon');
		});

		// Client-side validation checks
		if (usernameInput.value.trim() === '') {
			applyShakeAndErrorText(usernameInput, 'Please insert a username.', originalUsernamePlaceholder);
			primaryErrorMessage = primaryErrorMessage || 'Please fill in all required fields.';
			isValid = false;
		}

		if (emailInput.value.trim() === '') {
			applyShakeAndErrorText(emailInput, 'Please insert an email address.', originalEmailPlaceholder);
			primaryErrorMessage = primaryErrorMessage || 'Please fill in all required fields.';
			isValid = false;
		} else if (!emailInput.value.trim().match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/)) {
			applyShakeAndErrorText(emailInput, 'Please enter a valid email address.', originalEmailPlaceholder);
			primaryErrorMessage = primaryErrorMessage || 'Please enter a valid email address.';
			isValid = false;
		}

		if (passwordInput.value.trim() === '') {
			applyShakeAndErrorText(passwordInput, 'Please insert a password.', originalPasswordPlaceholder);
			primaryErrorMessage = primaryErrorMessage || 'Please fill in all required fields.';
			isValid = false;
		} else if (passwordInput.value.length < 6 || passwordInput.value.length > 15) {
			applyShakeAndErrorText(passwordInput, 'Password must be 6-15 characters.', originalPasswordPlaceholder);
			primaryErrorMessage = primaryErrorMessage || 'Password must be between 6 and 15 characters.';
			isValid = false;
		}

		if (confirmPasswordInput.value.trim() === '') {
			applyShakeAndErrorText(confirmPasswordInput, 'Please confirm your password.', originalConfirmPasswordPlaceholder);
			primaryErrorMessage = primaryErrorMessage || 'Please fill in all required fields.';
			isValid = false;
		}

		// Password match check (only if both are not empty)
		if (passwordInput.value.trim() !== '' && confirmPasswordInput.value.trim() !== '' && passwordInput.value !== confirmPasswordInput.value) {
			applyShakeAndErrorText(passwordInput, 'Passwords do not match.', originalPasswordPlaceholder);
			applyShakeAndErrorText(confirmPasswordInput, 'Passwords do not match.', originalConfirmPasswordPlaceholder);
			primaryErrorMessage = primaryErrorMessage || 'Passwords do not match.';
			isValid = false;
		}

		if (!isValid) {
			event.preventDefault(); // Prevent form submission
			// If there's a primary error message, display it in the div and clear it after 3s
			if (primaryErrorMessage) {
				errorDiv.textContent = primaryErrorMessage;
				setTimeout(() => errorDiv.textContent = '', 3000);
			}
		}
		// If isValid is true, the form will submit normally.
	});

	// Handle server-side errors from URL parameters
	const urlParams = new URLSearchParams(window.location.search);
	const serverError = urlParams.get('error');
	if (serverError) {
		let message;
		switch (serverError) {
			case 'passwords_do_not_match':
				message = 'Passwords do not match';
				break;
			case 'invalid_email':
				message = 'Invalid email address';
				break;
			case 'user_exists':
				message = 'Username or email already exists';
				break;
			case 'database_error':
				message = 'Database error, please try again later';
				break;
			default:
				message = 'An error occurred';
		}
		errorDiv.textContent = message;
		setTimeout(() => errorDiv.textContent = '', 3000); // Clear server error messages too
	}
});