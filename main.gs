/**
 * Google Workspace User Onboarding - Backend Script
 * Version 1.30
 * 
 * This Google Apps Script handles the backend functionality for creating
 * new Google Workspace users through a web interface.
 * 
 * Required APIs:
 * - Admin SDK Directory API
 * - Admin SDK License Manager API (optional, for license checking)
 * 
 * Required Permissions:
 * - Super Admin access to Google Workspace
 */

/**
 * Serves the HTML web interface
 * Called when the web app URL is accessed
 * 
 * @returns {HtmlOutput} The HTML page for the onboarding form
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Google Workspace Onboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Retrieves all organizational units from Google Workspace
 * Used to populate the OU dropdown in the form
 * 
 * @returns {Array<Object>} Array of OU objects with name and orgUnitPath
 * @throws {Error} If fetching OUs fails
 */
function getOUs() {
  try {
    // Fetch all organizational units from the domain
    const admin = AdminDirectory.Orgunits.list('my_customer', {
      type: 'all'
    });
    
    // Return empty array if no OUs found
    if (!admin.organizationUnits) {
      return [];
    }

    // Log all OUs for debugging purposes
    admin.organizationUnits.forEach(function(ou) {
      Logger.log('OU: ' + ou.name + ' | Path: ' + ou.orgUnitPath);
    });

    // Sort OUs alphabetically by their full path
    const sortedOUs = admin.organizationUnits.slice().sort(function(a, b) {
      return a.orgUnitPath.localeCompare(b.orgUnitPath);
    });

    // Format OUs for display, removing leading slash from path
    return sortedOUs.map(function(ou) {
      let displayPath = ou.orgUnitPath.startsWith('/') ? ou.orgUnitPath.substring(1) : ou.orgUnitPath;
      // If path is empty after removing '/', use the OU name (typically root OU)
      if (!displayPath) displayPath = ou.name;
      return {
        orgUnitPath: ou.orgUnitPath,
        name: displayPath
      };
    });
  } catch (error) {
    console.error('Error fetching OUs:', error);
    throw new Error('Failed to fetch organizational units');
  }
}

/**
 * Checks if an email address already exists in the domain
 * Searches both primary emails and aliases to prevent conflicts
 * 
 * @param {string} email - The email address to check
 * @returns {boolean} True if email exists, false otherwise
 */
function emailExistsAnywhere(email) {
  // Check if email exists as a primary email
  try {
    const user = AdminDirectory.Users.get(email);
    if (user && user.primaryEmail && user.primaryEmail.toLowerCase() === email.toLowerCase()) {
      return true;
    }
  } catch (e) {
    // User not found as primary email, continue checking aliases
  }
  
  // Check if email exists as an alias
  let pageToken;
  do {
    const response = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 100,
      pageToken: pageToken,
      fields: 'users(primaryEmail,aliases),nextPageToken'
    });
    
    // Search through all users' aliases
    if (response.users && response.users.length > 0) {
      for (let u of response.users) {
        if (u.aliases && u.aliases.map(a => a.toLowerCase()).includes(email.toLowerCase())) {
          return true;
        }
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return false;
}

/**
 * Checks if there are available Business Standard licenses
 * Optional function - can be removed if license checking is not needed
 * 
 * @returns {boolean} True if licenses are available, false otherwise
 */
function hasAvailableBusinessStandardLicense() {
  try {
    const skuId = '1010020020'; // Business Standard SKU ID
    const productId = 'Google-Apps';
    const info = AdminLicenseManager.LicenseCounts.get(productId, skuId);
    // Check if total licenses minus assigned licenses is greater than 0
    return (info.total - info.assigned) > 0;
  } catch (e) {
    Logger.log('Failed to check license availability: ' + e.toString());
    return false;
  }
}

/**
 * Verifies if the current user has Super Admin privileges
 * Attempts to list users as a permission check
 * 
 * @returns {boolean} True if user is Super Admin, false otherwise
 */
function isSuperAdmin() {
  try {
    // Try to list users (requires admin privileges)
    AdminDirectory.Users.list({customer: 'my_customer', maxResults: 1});
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Creates a new Google Workspace user with provided form data
 * Main function called when the form is submitted
 * 
 * @param {Object} formData - Object containing all user information from the form
 * @param {string} formData.firstName - User's first name
 * @param {string} formData.lastName - User's last name
 * @param {string} formData.email - Primary email address
 * @param {string} formData.title - Job title
 * @param {string} formData.department - Department name
 * @param {string} formData.secondaryEmail - Recovery/secondary email
 * @param {string} formData.phoneNumber - Phone number with country code
 * @param {string} formData.ou - Organizational unit path
 * @param {string} formData.manager - Manager's email address (optional)
 * @param {string} formData.managerName - Manager's full name (optional)
 * @returns {Object} Success response with created user details
 * @throws {Error} If user creation fails or validation fails
 */
function createUser(formData) {
  try {
    // Verify the current user has Super Admin permissions
    if (!isSuperAdmin()) {
      throw new Error('You must be a Google Workspace super admin to use this tool.');
    }

    // Check if the email already exists as a primary email or alias
    if (emailExistsAnywhere(formData.email)) {
      throw new Error('A user or alias with this email already exists');
    }

    // Define all required fields for validation
    const requiredFields = {
      firstName: 'First Name',
      lastName: 'Last Name',
      email: 'Primary Email',
      title: 'Title',
      department: 'Department',
      secondaryEmail: 'Secondary Email',
      phoneNumber: 'Phone Number',
      ou: 'Organizational Unit'
    };

    // Validate that all required fields are present
    for (const [field, label] of Object.entries(requiredFields)) {
      if (!formData[field]) {
        throw new Error(`${label} is required`);
      }
    }

    // Get the organization's domain from the admin's email
    const adminEmail = Session.getActiveUser().getEmail();
    const domain = adminEmail.split('@')[1];

    // Validate primary email format
    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    if (!emailRegex.test(formData.email)) {
      throw new Error('Invalid primary email format');
    }

    // Validate secondary email format
    if (!emailRegex.test(formData.secondaryEmail)) {
      throw new Error('Invalid secondary email format');
    }

    // Ensure primary email uses the organization's domain
    if (!formData.email.endsWith('@' + domain)) {
      throw new Error('Primary email must use the organization domain: @' + domain);
    }

    // Validate phone number format: +<countrycode><number>, 8-15 digits total
    const phoneRegex = /^\+\d{8,15}$/;
    if (!phoneRegex.test(formData.phoneNumber)) {
      throw new Error('Phone number must include country code and be 8-15 digits, e.g., +14165551234');
    }

    // Generate a random secure password for the new user
    const password = generatePassword();
    
    // Build the user object for creation
    const user = {
      primaryEmail: formData.email,
      name: {
        givenName: formData.firstName,
        familyName: formData.lastName
      },
      password: password,
      changePasswordAtNextLogin: true, // Force password change on first login
      organizations: [{
        title: formData.title,
        department: formData.department,
        primary: true
      }],
      orgUnitPath: formData.ou,
      emails: [{
        address: formData.secondaryEmail,
        type: 'work'
      }],
      phones: [{
        value: formData.phoneNumber,
        type: 'work'
      }]
    };

    // Add manager relationship if provided
    if (formData.manager) {
      user.relations = [{
        type: 'manager',
        value: formData.manager
      }];
    }

    // Create the user via Admin SDK
    const createdUser = AdminDirectory.Users.insert(user);

    // Log successful creation
    Logger.log('User created successfully: ' + formData.email);

    // Return success response with user details
    return { 
      success: true, 
      message: 'User created successfully',
      user: {
        email: createdUser.primaryEmail,
        name: formData.firstName + ' ' + formData.lastName,
        department: formData.department,
        title: formData.title,
        manager: formData.manager,
        managerName: formData.managerName
      }
    };
  } catch (error) {
    Logger.log('Error creating user: ' + error.toString());
    throw new Error('Failed to create user: ' + error.message);
  }
}

/**
 * Generates a secure random password
 * Creates a password with at least one uppercase, lowercase, number, and special character
 * 
 * @returns {string} A randomly generated 12-character password
 */
function generatePassword() {
  const length = 12;
  const charset = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*';
  let password = '';
  
  // Ensure at least one of each required character type for password strength
  password += 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[Math.floor(Math.random() * 26)]; // Uppercase letter
  password += 'abcdefghijklmnopqrstuvwxyz'[Math.floor(Math.random() * 26)]; // Lowercase letter
  password += '0123456789'[Math.floor(Math.random() * 10)]; // Number
  password += '!@#$%^&*'[Math.floor(Math.random() * 8)]; // Special character
  
  // Fill the remaining characters randomly from the full charset
  for (let i = password.length; i < length; i++) {
    password += charset[Math.floor(Math.random() * charset.length)];
  }
  
  // Shuffle the password to randomize the position of guaranteed characters
  return password.split('').sort(() => Math.random() - 0.5).join('');
}

/**
 * Retrieves all active users from Google Workspace
 * Used for the manager autocomplete functionality in the form
 * 
 * @returns {Array<Object>} Array of user objects with email and name
 * @throws {Error} If fetching users fails
 */
function getAllUsers() {
  try {
    const users = [];
    let pageToken;
    
    // Paginate through all users (max 100 per request)
    do {
      const response = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 100,
        orderBy: 'givenName', // Sort alphabetically by first name
        pageToken: pageToken,
        query: 'isSuspended=false', // Only include active (non-suspended) users
        fields: 'users(primaryEmail,name/fullName),nextPageToken' // Limit returned fields for efficiency
      });
      
      // Add users to the results array
      if (response.users && response.users.length > 0) {
        response.users.forEach(function(user) {
          users.push({
            email: user.primaryEmail,
            name: user.name.fullName
          });
        });
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    return users;
  } catch (error) {
    console.error('Error fetching users:', error);
    throw new Error('Failed to fetch users');
  }
}
