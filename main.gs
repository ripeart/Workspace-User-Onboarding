// Version 1.30
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Google Workspace Onboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOUs() {
  try {
    const admin = AdminDirectory.Orgunits.list('my_customer', {
      type: 'all'
    });
    
    if (!admin.organizationUnits) {
      return [];
    }

    // Log all OUs for debugging
    admin.organizationUnits.forEach(function(ou) {
      Logger.log('OU: ' + ou.name + ' | Path: ' + ou.orgUnitPath);
    });

    // Sort OUs alphabetically by orgUnitPath
    const sortedOUs = admin.organizationUnits.slice().sort(function(a, b) {
      return a.orgUnitPath.localeCompare(b.orgUnitPath);
    });

    // Return OUs with full path as name, but remove leading slash
    return sortedOUs.map(function(ou) {
      let displayPath = ou.orgUnitPath.startsWith('/') ? ou.orgUnitPath.substring(1) : ou.orgUnitPath;
      // If the path is empty after removing '/', use the OU name (root OU)
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

function emailExistsAnywhere(email) {
  // Check for primary email
  try {
    const user = AdminDirectory.Users.get(email);
    if (user && user.primaryEmail && user.primaryEmail.toLowerCase() === email.toLowerCase()) {
      return true;
    }
  } catch (e) {}
  // Check for alias
  let pageToken;
  do {
    const response = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 100,
      pageToken: pageToken,
      fields: 'users(primaryEmail,aliases),nextPageToken'
    });
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

function hasAvailableBusinessStandardLicense() {
  try {
    const skuId = '1010020020';
    const productId = 'Google-Apps';
    const info = AdminLicenseManager.LicenseCounts.get(productId, skuId);
    // info.licenses: { assigned, total, inUse }
    return (info.total - info.assigned) > 0;
  } catch (e) {
    Logger.log('Failed to check license availability: ' + e.toString());
    return false;
  }
}

function isSuperAdmin() {
  try {
    // Try to list users (requires admin privileges)
    AdminDirectory.Users.list({customer: 'my_customer', maxResults: 1});
    return true;
  } catch (e) {
    return false;
  }
}

function createUser(formData) {
  try {
    // Check if user is a super admin
    if (!isSuperAdmin()) {
      throw new Error('You must be a Google Workspace super admin to use this tool.');
    }

    // Check if email exists as primary or alias
    if (emailExistsAnywhere(formData.email)) {
      throw new Error('A user or alias with this email already exists');
    }

    // Validate all required fields
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

    // Check for missing required fields
    for (const [field, label] of Object.entries(requiredFields)) {
      if (!formData[field]) {
        throw new Error(`${label} is required`);
      }
    }

    // Get the domain from the admin email
    const adminEmail = Session.getActiveUser().getEmail();
    const domain = adminEmail.split('@')[1];

    // Validate email format
    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    if (!emailRegex.test(formData.email)) {
      throw new Error('Invalid primary email format');
    }

    // Validate secondary email format
    if (!emailRegex.test(formData.secondaryEmail)) {
      throw new Error('Invalid secondary email format');
    }

    // Ensure primary email is using the correct domain
    if (!formData.email.endsWith('@' + domain)) {
      throw new Error('Primary email must use the organization domain: @' + domain);
    }

    // Validate phone number format: +<countrycode><number>, 8-15 digits, no spaces/dashes
    const phoneRegex = /^\+\d{8,15}$/;
    if (!phoneRegex.test(formData.phoneNumber)) {
      throw new Error('Phone number must include country code and be 8-15 digits, e.g., +14165551234');
    }

    // Create the user object
    const password = generatePassword();
    const user = {
      primaryEmail: formData.email,
      name: {
        givenName: formData.firstName,
        familyName: formData.lastName
      },
      password: password,
      changePasswordAtNextLogin: true,
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

    // Add manager relation if provided
    if (formData.manager) {
      user.relations = [{
        type: 'manager',
        value: formData.manager
      }];
    }

    // Create the user
    const createdUser = AdminDirectory.Users.insert(user);

    // Log the creation
    Logger.log('User created successfully: ' + formData.email);

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

function generatePassword() {
  const length = 12;
  const charset = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*';
  let password = '';
  
  // Ensure at least one of each required character type
  password += 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[Math.floor(Math.random() * 26)]; // Uppercase
  password += 'abcdefghijklmnopqrstuvwxyz'[Math.floor(Math.random() * 26)]; // Lowercase
  password += '0123456789'[Math.floor(Math.random() * 10)]; // Number
  password += '!@#$%^&*'[Math.floor(Math.random() * 8)]; // Special character
  
  // Fill the rest randomly
  for (let i = password.length; i < length; i++) {
    password += charset[Math.floor(Math.random() * charset.length)];
  }
  
  // Shuffle the password
  return password.split('').sort(() => Math.random() - 0.5).join('');
}

function getAllUsers() {
  try {
    const users = [];
    let pageToken;
    do {
      const response = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 100,
        orderBy: 'givenName',
        pageToken: pageToken,
        query: 'isSuspended=false',
        fields: 'users(primaryEmail,name/fullName),nextPageToken'
      });
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
