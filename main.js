
const SHEETS = {
  Users: ['id', 'name', 'email', 'passwordHash', 'salt', 'role', 'organization', 'createdAt'],
  Donations: ['id', 'donorId', 'foodName', 'category', 'quantity', 'unit', 'preparedAt', 'pickupDeadline', 'location', 'storageDetails', 'allergens', 'status', 'createdAt'],
  Requests: ['id', 'donationId', 'ngoId', 'volunteerId', 'message', 'status', 'createdAt']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FoodBridge')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupFoodBridge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEETS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(SHEETS[name]);
      sheet.setFrozenRows(1);
    }
  });

  return 'FoodBridge database is ready!';
}

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Database not initialized. Run setupFoodBridge first.');
  return sheet;
}


function getRows(name) {
  const sheet = getSheet(name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  return values.slice(1)
    .filter(row => row[0] !== '')
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        const value = row[i];
        obj[header] = value instanceof Date
          ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
          : value;
      });
      return obj;
    });
}


function addRow(name, obj) {
  const headers = SHEETS[name];
  getSheet(name).appendRow(headers.map(h => obj[h] ?? ''));
}

function updateRow(name, id, changes) {
  const sheet = getSheet(name);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idColumn = headers.indexOf('id');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idColumn]) === String(id)) {
      Object.keys(changes).forEach(key => {
        const col = headers.indexOf(key);
        if (col >= 0) sheet.getRange(r + 1, col + 1).setValue(changes[key]);
      });
      return true;
    }
  }
  throw new Error('Record not found.');
}

function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + password,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function registerUser(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const name = String(data.name || '').trim();
    const email = String(data.email || '').trim().toLowerCase();
    const password = String(data.password || '');
    const role = String(data.role || '');
    const organization = String(data.organization || '').trim();

    if (!name || !email || !password) throw new Error('Please fill in all required fields.');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email.');
    if (password.length < 8) throw new Error('Password must be at least 8 characters.');
    if (!['Donor', 'NGO', 'Volunteer'].includes(role)) throw new Error('Invalid account role.');

    if (getRows('Users').some(u => String(u.email).toLowerCase() === email)) {
      throw new Error('An account with this email already exists.');
    }

    const salt = Utilities.getUuid();
    addRow('Users', {
      id: Utilities.getUuid(),
      name, email,
      passwordHash: hashPassword(password, salt),
      salt, role, organization,
      createdAt: new Date()
    });

    return { success: true, message: 'Account created. Please log in.' };
  } finally {
    lock.releaseLock();
  }
}

function loginUser(data) {
  const email = String(data.email || '').trim().toLowerCase();
  const password = String(data.password || '');

  const user = getRows('Users').find(u => String(u.email).toLowerCase() === email);
  if (!user || hashPassword(password, user.salt) !== user.passwordHash) {
    throw new Error('Incorrect email or password.');
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  const safeUser = {
    id: user.id,
    name: user.name,
    email: user.email,
    role: user.role,
    organization: user.organization
  };

  CacheService.getScriptCache().put('session_' + token, JSON.stringify(safeUser), 21600);
  return { token, user: safeUser };
}

function getSession(token) {
  if (!token) throw new Error('Please log in.');
  const value = CacheService.getScriptCache().get('session_' + token);
  if (!value) throw new Error('Your session expired. Please log in again.');
  return JSON.parse(value);
}

function logoutUser(token) {
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return true;
}

function getDashboard(token) {
  const user = getSession(token);
  const donations = getRows('Donations');
  const requests = getRows('Requests');

  let visibleDonations = donations;
  let visibleRequests = requests;

  if (user.role === 'Donor') {
    visibleDonations = donations.filter(d => d.donorId === user.id);
    const ids = visibleDonations.map(d => d.id);
    visibleRequests = requests.filter(r => ids.includes(r.donationId));
  } else if (user.role === 'NGO') {
    visibleRequests = requests.filter(r => r.ngoId === user.id);
  } else if (user.role === 'Volunteer') {
    visibleRequests = requests.filter(r =>
  r.volunteerId === user.id ||
  (r.status === 'Accepted' && !r.volunteerId));
  }

  return {
    user,
    donations: visibleDonations,
    requests: visibleRequests,
    available: donations.filter(d => d.status === 'Available')
  };
}

function createDonation(token, data) {
  const user = getSession(token);
  if (user.role !== 'Donor' && user.role !== 'Admin') throw new Error('Only donors can post food.');

  const foodName = String(data.foodName || '').trim();
  const quantity = Number(data.quantity);
  const location = String(data.location || '').trim();

  if (!foodName || !location || !Number.isFinite(quantity) || quantity <= 0) {
    throw new Error('Enter a food name, valid quantity, and pickup location.');
  }
  if (!data.pickupDeadline) throw new Error('Please enter a pickup deadline.');

  addRow('Donations', {
    id: Utilities.getUuid(),
    donorId: user.id,
    foodName,
    category: String(data.category || 'Other'),
    quantity,
    unit: String(data.unit || 'portions'),
    preparedAt: String(data.preparedAt || ''),
    pickupDeadline: String(data.pickupDeadline),
    location,
    storageDetails: String(data.storageDetails || ''),
    allergens: String(data.allergens || ''),
    status: 'Available',
    createdAt: new Date()
  });

  return { success: true };
}

function requestDonation(token, donationId, message) {
  const user = getSession(token);
  if (user.role !== 'NGO') throw new Error('Only NGOs can request donations.');

  const donation = getRows('Donations').find(d => d.id === donationId);
  if (!donation || donation.status !== 'Available') throw new Error('This donation is no longer available.');

  const requests = getRows('Requests');
  if (requests.some(r => r.donationId === donationId && r.ngoId === user.id &&
      ['Pending', 'Accepted'].includes(r.status))) {
    throw new Error('You already have an active request for this donation.');
  }

  addRow('Requests', {
    id: Utilities.getUuid(),
    donationId,
    ngoId: user.id,
    volunteerId: '',
    message: String(message || ''),
    status: 'Pending',
    createdAt: new Date()
  });

  return { success: true };
}

function updateRequest(token, requestId, action) {
  const user = getSession(token);
  const request = getRows('Requests').find(r => r.id === requestId);
  if (!request) throw new Error('Request not found.');

  const donation = getRows('Donations').find(d => d.id === request.donationId);
  if (!donation) throw new Error('Donation not found.');

  if (user.role === 'Donor') {
    if (donation.donorId !== user.id) throw new Error('Not authorized.');
    if (!['Accepted', 'Rejected'].includes(action) || request.status !== 'Pending') {
      throw new Error('Invalid request action.');
    }
    updateRow('Requests', requestId, { status: action });
    if (action === 'Accepted') updateRow('Donations', donation.id, { status: 'Accepted' });
    return { success: true };
  }

  if (user.role === 'Volunteer') {
    if (action === 'Claim') {
      if (request.status !== 'Accepted' || request.volunteerId) throw new Error('This pickup is not available.');
      updateRow('Requests', requestId, { volunteerId: user.id });
      return { success: true };
    }
    if (request.volunteerId !== user.id) throw new Error('This pickup is assigned to someone else.');
    if (action === 'Collected' || action === 'Delivered') {
      if (action === 'Delivered' && donation.status !== 'Collected') {
        throw new Error('Mark the food as collected first.');
      }
      updateRow('Donations', donation.id, { status: action });
      return { success: true };
    }
  }

  if (user.role === 'Admin') {
    if (!['Accepted', 'Rejected', 'Collected', 'Delivered', 'Cancelled'].includes(action)) {
      throw new Error('Invalid status.');
    }
    if (['Accepted', 'Rejected'].includes(action)) updateRow('Requests', requestId, { status: action });
    else updateRow('Donations', donation.id, { status: action });
    return { success: true };
  }

  throw new Error('You are not authorized to perform this action.');
}

function getAdminData(token) {
  const user = getSession(token);
  if (user.role !== 'Admin') throw new Error('Administrator access only.');
  return {
    users: getRows('Users').map(u => ({
      id: u.id, name: u.name, email: u.email, role: u.role
    })),
    donations: getRows('Donations'),
    requests: getRows('Requests')
  };
}

