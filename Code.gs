/** Baby Rafael — Meals & Gifts for Kristina and Daniel.
 * Add an HTML file named index and paste the HTML below into it.
 * Deploy as a Web app, executing as Me.
 */
const CONFIG = Object.freeze({
  spreadsheetId: '1U97EEnzbXx9OKWKxq6J5zx5LRP3Tpll4zBbstrSQg30',
  sheetName: 'SIGNUPS',
  timeZone: 'America/Chicago',
  types: ['Breakfast', 'Lunch', 'Dinner', 'Snacks', 'Gift Card', 'Baby Gift', 'Other'],
  mealTypes: ['Breakfast', 'Lunch', 'Dinner']
});

const HEADERS = [
  'Name',
  'Date',
  'Support Type',
  'What You Are Bringing',
  'Delivery Time',
  'Public Note',
  'Gift Card Amount',
  'Submitted At',
  'Private Contact',
  'Submission ID'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('A little love for Rafael')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Public response omits contact details, amounts, timestamps, and IDs.
function getPageData() {
  const today = today_();

  const signups = rows_(sheet_())
    .map(publicRow_)
    .filter(function (r) {
      return r.name && r.date && r.date >= today;
    })
    .sort(function (a, b) {
      return (a.date + a.time + a.type)
        .localeCompare(b.date + b.time + b.type);
    });

  return {
    today: today,
    timeZone: CONFIG.timeZone,
    types: CONFIG.types,
    signups: signups
  };
}

function submitSignup(input) {
  const item = validate_(input);
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    throw new Error(
      'Someone else is signing up. Please try again in a moment.'
    );
  }

  try {
    const sheet = sheet_();
    const rows = rows_(sheet);

    // A retry after a lost response never creates a second row.
    const existing = rows.find(function (r) {
      return String(r[9]) === item.requestId;
    });

    if (existing) {
      return { ok: true, alreadySaved: true };
    }

    const isMeal = CONFIG.mealTypes.indexOf(item.type) !== -1;
    const taken = rows.some(function (r) {
      return dateString_(r[1]) === item.date &&
        String(r[2]) === item.type;
    });

    if (isMeal && taken) {
      throw new Error(
        'That ' + item.type.toLowerCase() +
        ' is already covered. Choose another date or support type.'
      );
    }

    const next = sheet.getLastRow() + 1;

    if (next > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 100);
    }

    // Preserve the supplied table's column types.
    const values = [
      item.name,
      item.date,
      item.type,
      item.item,
      item.time,
      item.note,
      item.amount,
      Utilities.formatDate(
        new Date(),
        CONFIG.timeZone,
        'yyyy-MM-dd HH:mm:ss'
      ),
      item.contact,
      item.requestId
    ];

    sheet.getRange(next, 1, 1, HEADERS.length)
      .setValues([values.map(safeCell_)]);

    SpreadsheetApp.flush();

    return { ok: true, alreadySaved: false };
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const sheet = SpreadsheetApp
    .openById(CONFIG.spreadsheetId)
    .getSheetByName(CONFIG.sheetName);

  if (!sheet) {
    throw new Error(
      'The signup sheet is unavailable. Please contact the organizer.'
    );
  }

  const actual = sheet
    .getRange(1, 1, 1, HEADERS.length)
    .getDisplayValues()[0];

  if (actual.some(function (h, i) {
    return h !== HEADERS[i];
  })) {
    throw new Error(
      'The signup sheet headers have changed. ' +
      'Please ask the organizer to restore them.'
    );
  }

  return sheet;
}

function rows_(sheet) {
  const count = sheet.getLastRow() - 1;

  return count > 0
    ? sheet.getRange(2, 1, count, HEADERS.length).getValues()
    : [];
}

function today_() {
  return Utilities.formatDate(
    new Date(),
    CONFIG.timeZone,
    'yyyy-MM-dd'
  );
}

function dateString_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      CONFIG.timeZone,
      'yyyy-MM-dd'
    );
  }

  const text = String(value || '').trim();
  return validDate_(text) ? text : '';
}

function timeString_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.timeZone, 'HH:mm');
  }

  const text = String(value || '').trim();
  return /^([01]\d|2[0-3]):[0-5]\d$/.test(text) ? text : '';
}

function publicRow_(r) {
  return {
    name: String(r[0] || ''),
    date: dateString_(r[1]),
    type: String(r[2] || ''),
    item: String(r[3] || ''),
    time: timeString_(r[4]),
    note: String(r[5] || '')
  };
}

function text_(value, label, max, required) {
  if (
    value !== undefined &&
    value !== null &&
    typeof value !== 'string'
  ) {
    throw new Error(label + ' must be text.');
  }

  const result = String(value || '').trim();

  if (required && !result) {
    throw new Error('Please enter ' + label.toLowerCase() + '.');
  }

  if (result.length > max) {
    throw new Error(
      label + ' must be ' + max + ' characters or fewer.'
    );
  }

  return result;
}

function validDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return false;
  }

  const d = new Date(s + 'T12:00:00Z');

  return !isNaN(d.getTime()) &&
    d.toISOString().slice(0, 10) === s;
}

function validate_(input) {
  if (
    !input ||
    typeof input !== 'object' ||
    Array.isArray(input)
  ) {
    throw new Error('Please complete the signup form.');
  }

  if (input.website) {
    throw new Error('Unable to accept this signup.');
  }

  const item = {
    name: text_(input.name, 'Name', 80, true),
    date: text_(input.date, 'Date', 10, true),
    type: text_(input.type, 'Support type', 30, true),
    item: text_(
      input.item,
      'What you are bringing',
      220,
      true
    ),
    time: text_(input.time, 'Delivery time', 5, false),
    note: text_(input.note, 'Public note', 400, false),
    contact: text_(
      input.contact,
      'Contact information',
      160,
      false
    ),
    requestId: text_(
      input.requestId,
      'Submission reference',
      80,
      true
    ),
    amount: ''
  };

  if (!/^[a-zA-Z0-9_-]{16,80}$/.test(item.requestId)) {
    throw new Error('Please reload the page and try again.');
  }

  if (!validDate_(item.date) || item.date < today_()) {
    throw new Error('Please choose today or a future date.');
  }

  if (CONFIG.types.indexOf(item.type) === -1) {
    throw new Error('Please select a listed support type.');
  }

  if (
    item.time &&
    !/^([01]\d|2[0-3]):[0-5]\d$/.test(item.time)
  ) {
    throw new Error('Please enter a valid delivery time.');
  }

  if (
    input.amount !== '' &&
    input.amount !== undefined &&
    input.amount !== null
  ) {
    const amount = String(input.amount).trim();

    if (
      item.type !== 'Gift Card' ||
      !/^\d+(\.\d{1,2})?$/.test(amount) ||
      Number(amount) <= 0 ||
      Number(amount) > 10000
    ) {
      throw new Error(
        'Enter a gift-card amount between $0.01 and $10,000, ' +
        'or leave it blank.'
      );
    }

    item.amount = Number(amount);
  }

  return item;
}

function safeCell_(value) {
  // Store formula-like guest input as literal text.
  return typeof value === 'string' && /^[=+@'\-]/.test(value)
    ? "'" + value
    : value;
}