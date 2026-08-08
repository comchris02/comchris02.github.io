const RSVP_CONFIG = Object.freeze({
  siteUrl: 'https://shelbyandchrisgethitched.com/',
  setupSheet: 'Guest Setup',
  householdsSheet: 'Households',
  guestsSheet: 'Guests',
  inviteLinksSheet: 'Invite Links',
  spreadsheetProperty: 'RSVP_SPREADSHEET_ID',
  maxShortTextLength: 200,
  maxLongTextLength: 500
});

const RSVP_EVENTS = Object.freeze([
  { key: 'rehearsal', invitedHeader: 'rehearsal_invited', responseHeader: 'rehearsal_response', collectResponse: false },
  { key: 'rantaReception', invitedHeader: 'ranta_reception_invited', responseHeader: 'ranta_reception_response' },
  { key: 'welcomeCelebration', invitedHeader: 'welcome_celebration_invited', responseHeader: 'welcome_celebration_response' },
  { key: 'bridalPartyPrep', invitedHeader: 'bridal_party_prep_invited', responseHeader: 'bridal_party_prep_response', collectResponse: false },
  { key: 'preWeddingPhotos', invitedHeader: 'pre_wedding_photos_invited', responseHeader: 'pre_wedding_photos_response', collectResponse: false },
  { key: 'wedding', invitedHeader: 'wedding_invited', responseHeader: 'wedding_response' },
  { key: 'afterparty', invitedHeader: 'afterparty_invited', responseHeader: 'afterparty_response' },
  { key: 'cabanaBaySendoff', invitedHeader: 'cabana_bay_sendoff_invited', responseHeader: 'cabana_bay_sendoff_response' },
  { key: 'welcomeTransportation', invitedHeader: 'welcome_celebration_invited', responseHeader: 'welcome_transportation_response' }
]);

const RSVP_INVITED_HEADERS = Object.freeze(RSVP_EVENTS.reduce(function (headers, eventDefinition) {
  if (headers.indexOf(eventDefinition.invitedHeader) === -1) {
    headers.push(eventDefinition.invitedHeader);
  }
  return headers;
}, []));

const RSVP_HOUSEHOLD_FIELDS = Object.freeze([
  { key: 'visitedOrlando', header: 'visited_orlando' },
  { key: 'danceFloorSong', header: 'dance_floor_song' },
  { key: 'karaokeSong', header: 'karaoke_song' },
  { key: 'anticipatedHotel', header: 'anticipated_hotel' },
  { key: 'cabanaBayArrival', header: 'cabana_bay_arrival' },
  { key: 'accessibilityNeeds', header: 'accessibility_needs' },
  { key: 'dietaryRestrictions', header: 'dietary_restrictions' },
  { key: 'contactPhone', header: 'contact_phone' }
]);

const RSVP_HEADERS = Object.freeze({
  setup: [
    'household_key',
    'household_name',
    'invite_code',
    'contact_email',
    'guest_name'
  ].concat(RSVP_INVITED_HEADERS),
  households: [
    'household_key',
    'household_name',
    'household_id',
    'invite_code_hash',
    'contact_email',
    'updated_at'
  ].concat(RSVP_HOUSEHOLD_FIELDS.map(function (fieldDefinition) {
    return fieldDefinition.header;
  })),
  guests: [
    'guest_id',
    'household_id',
    'guest_name'
  ].concat(
    RSVP_INVITED_HEADERS,
    RSVP_EVENTS.map(function (eventDefinition) {
      return eventDefinition.responseHeader;
    })
  ),
  inviteLinks: [
    'household_key',
    'household_name',
    'contact_email',
    'invite_code',
    'personalized_link'
  ]
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RSVP Admin')
    .addItem('1. Set up RSVP sheets', 'setupRsvpSystem')
    .addItem('2. Build invitations from Guest Setup', 'buildInvitationsFromSetup')
    .addSeparator()
    .addItem('Change one household code', 'rotateInvitationCodeFromMenu')
    .addToUi();
}

function setupRsvpSystem() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Open this script from the Google Sheet, then run setupRsvpSystem again.');
  }

  PropertiesService.getScriptProperties().setProperty(
    RSVP_CONFIG.spreadsheetProperty,
    spreadsheet.getId()
  );

  ensureSheet_(spreadsheet, RSVP_CONFIG.setupSheet, RSVP_HEADERS.setup);
  ensureSheet_(spreadsheet, RSVP_CONFIG.householdsSheet, RSVP_HEADERS.households);
  ensureSheet_(spreadsheet, RSVP_CONFIG.guestsSheet, RSVP_HEADERS.guests);
  ensureSheet_(spreadsheet, RSVP_CONFIG.inviteLinksSheet, RSVP_HEADERS.inviteLinks);

  const setupSheet = spreadsheet.getSheetByName(RSVP_CONFIG.setupSheet);
  if (setupSheet.getLastRow() === 1) {
    setupSheet.getRange(2, 1, 2, RSVP_HEADERS.setup.length).setValues([
      [
        'SMITH', 'The Smith Family', 'SmithSunset', 'smith@example.com', 'Alex Smith',
        true, true, true, true, true, true, true, true
      ],
      [
        'SMITH', 'The Smith Family', 'SmithSunset', 'smith@example.com', 'Jordan Smith',
        false, false, true, false, false, true, true, true
      ]
    ]);
    setupSheet.getRange(2, 1, 2, RSVP_HEADERS.setup.length)
      .setBackground('#fff4e8');
    setupSheet.getRange('A4').setNote(
      'Delete the two orange example rows before adding the real guest list.'
    );
  }

  SpreadsheetApp.getUi().alert(
    'RSVP sheets are ready. Replace the orange sample rows in "Guest Setup" with the real guest list, then use RSVP Admin → Build invitations from Guest Setup.'
  );
}

function buildInvitationsFromSetup() {
  const spreadsheet = getSpreadsheet_();
  const setupSheet = spreadsheet.getSheetByName(RSVP_CONFIG.setupSheet);
  const householdsSheet = spreadsheet.getSheetByName(RSVP_CONFIG.householdsSheet);
  const guestsSheet = spreadsheet.getSheetByName(RSVP_CONFIG.guestsSheet);
  const inviteLinksSheet = spreadsheet.getSheetByName(RSVP_CONFIG.inviteLinksSheet);

  if (!setupSheet || !householdsSheet || !guestsSheet || !inviteLinksSheet) {
    throw new Error('Run setupRsvpSystem before building invitations.');
  }

  if (
    householdsSheet.getLastRow() > 1 ||
    guestsSheet.getLastRow() > 1 ||
    inviteLinksSheet.getLastRow() > 1
  ) {
    throw new Error(
      'Invitations have already been built. This safety check prevents existing RSVP responses from being overwritten.'
    );
  }

  const rows = setupSheet.getDataRange().getValues().slice(1);
  const householdsByKey = {};
  const householdKeyByCode = {};
  const householdOrder = [];

  rows.forEach(function (row, index) {
    if (row.every(function (value) { return String(value).trim() === ''; })) return;

    const sheetRow = index + 2;
    const householdKey = sanitizeIdentifier_(row[0], 'household_key', sheetRow);
    const householdName = sanitizeText_(row[1], 120, 'household_name', sheetRow);
    const inviteCode = sanitizeInviteCode_(row[2], sheetRow);
    const normalizedInviteCode = normalizeInviteCode_(inviteCode);
    const contactEmail = sanitizeEmail_(row[3], sheetRow);
    const guestName = sanitizeText_(row[4], 120, 'guest_name', sheetRow);

    if (
      householdKeyByCode[normalizedInviteCode] &&
      householdKeyByCode[normalizedInviteCode] !== householdKey
    ) {
      throw new Error(
        'invite_code "' + inviteCode + '" is assigned to more than one household.'
      );
    }
    householdKeyByCode[normalizedInviteCode] = householdKey;

    if (!householdsByKey[householdKey]) {
      householdsByKey[householdKey] = {
        key: householdKey,
        name: householdName,
        code: inviteCode,
        email: contactEmail,
        id: Utilities.getUuid(),
        guests: []
      };
      householdOrder.push(householdKey);
    }

    const household = householdsByKey[householdKey];
    if (
      household.name !== householdName ||
      normalizeInviteCode_(household.code) !== normalizedInviteCode ||
      household.email !== contactEmail
    ) {
      throw new Error(
        'Rows for household_key "' + householdKey +
        '" must use the same household name, RSVP password, and email.'
      );
    }

    const invitations = RSVP_INVITED_HEADERS.map(function (invitedHeader, eventIndex) {
      return parseBoolean_(row[5 + eventIndex]);
    });

    household.guests.push({
      id: Utilities.getUuid(),
      name: guestName,
      invitations: invitations
    });
  });

  if (householdOrder.length === 0) {
    throw new Error('Guest Setup does not contain any guest rows.');
  }

  const householdRows = [];
  const guestRows = [];
  const inviteRows = [];

  householdOrder.forEach(function (key) {
    const household = householdsByKey[key];
    householdRows.push(
      [
        safeForSheet_(household.key),
        safeForSheet_(household.name),
        household.id,
        hashInviteCode_(household.code),
        safeForSheet_(household.email),
        ''
      ].concat(RSVP_HOUSEHOLD_FIELDS.map(function () { return ''; }))
    );

    household.guests.forEach(function (guest) {
      guestRows.push(
        [
          guest.id,
          household.id,
          safeForSheet_(guest.name)
        ].concat(
          guest.invitations,
          RSVP_EVENTS.map(function () { return ''; })
        )
      );
    });

    inviteRows.push([
      safeForSheet_(household.key),
      safeForSheet_(household.name),
      safeForSheet_(household.email),
      household.code,
      RSVP_CONFIG.siteUrl + '?invite=' + encodeURIComponent(household.code) + '#rsvp'
    ]);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    householdsSheet.getRange(2, 1, householdRows.length, RSVP_HEADERS.households.length)
      .setValues(householdRows);
    guestsSheet.getRange(2, 1, guestRows.length, RSVP_HEADERS.guests.length)
      .setValues(guestRows);
    inviteLinksSheet.getRange(2, 1, inviteRows.length, RSVP_HEADERS.inviteLinks.length)
      .setValues(inviteRows);
    inviteLinksSheet.autoResizeColumns(1, RSVP_HEADERS.inviteLinks.length);
  } finally {
    lock.releaseLock();
  }

  SpreadsheetApp.getUi().alert(
    householdRows.length + ' household invitations and ' +
    guestRows.length + ' guest records were created. Keep the "Invite Links" sheet private.'
  );
}

function rotateInvitationCodeFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const householdPrompt = ui.prompt(
    'Change RSVP password',
    'Enter the household_key whose RSVP password should be replaced. The old password will stop working immediately.',
    ui.ButtonSet.OK_CANCEL
  );
  if (householdPrompt.getSelectedButton() !== ui.Button.OK) return;

  const householdKey = String(householdPrompt.getResponseText() || '').trim().toUpperCase();
  if (!/^[A-Z0-9_-]{2,40}$/.test(householdKey)) {
    ui.alert('That household_key is not valid.');
    return;
  }

  const codePrompt = ui.prompt(
    'Choose the new RSVP password',
    'Enter a unique 2–40 character password using letters, numbers, hyphens, or underscores. Do not use spaces.',
    ui.ButtonSet.OK_CANCEL
  );
  if (codePrompt.getSelectedButton() !== ui.Button.OK) return;

  const newCode = String(codePrompt.getResponseText() || '').trim();
  if (!isValidInviteCodeFormat_(newCode)) {
    ui.alert('Use 2–40 letters, numbers, hyphens, or underscores with no spaces.');
    return;
  }

  const spreadsheet = getSpreadsheet_();
  const householdsSheet = spreadsheet.getSheetByName(RSVP_CONFIG.householdsSheet);
  const inviteLinksSheet = spreadsheet.getSheetByName(RSVP_CONFIG.inviteLinksSheet);
  const householdRows = householdsSheet.getDataRange().getValues();
  const inviteRows = inviteLinksSheet.getDataRange().getValues();
  const newCodeHash = hashInviteCode_(newCode);
  let householdRowNumber = 0;
  let inviteRowNumber = 0;

  for (let index = 1; index < householdRows.length; index += 1) {
    if (String(householdRows[index][0]).toUpperCase() === householdKey) {
      householdRowNumber = index + 1;
      break;
    }
  }

  for (let index = 1; index < inviteRows.length; index += 1) {
    if (String(inviteRows[index][0]).toUpperCase() === householdKey) {
      inviteRowNumber = index + 1;
      break;
    }
  }

  if (!householdRowNumber || !inviteRowNumber) {
    ui.alert('No generated invitation was found for household_key "' + householdKey + '".');
    return;
  }

  if (String(householdRows[householdRowNumber - 1][3]) === newCodeHash) {
    ui.alert('That is already the RSVP password for this household.');
    return;
  }

  for (let index = 1; index < householdRows.length; index += 1) {
    if (
      index + 1 !== householdRowNumber &&
      String(householdRows[index][3]) === newCodeHash
    ) {
      ui.alert('That RSVP password is already assigned to another household.');
      return;
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    householdsSheet.getRange(householdRowNumber, 4).setValue(newCodeHash);
    inviteLinksSheet.getRange(inviteRowNumber, 4, 1, 2).setValues([[
      safeForSheet_(newCode),
      RSVP_CONFIG.siteUrl + '?invite=' + encodeURIComponent(newCode) + '#rsvp'
    ]]);
  } finally {
    lock.releaseLock();
  }

  ui.alert('The RSVP password for ' + householdKey + ' was changed. Use the updated password recorded in "Invite Links".');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Shelby & Chris RSVP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInvitation(inviteCode) {
  try {
    const household = findHouseholdByCode_(inviteCode);
    if (!household) {
      return {
        ok: false,
        launchNotice: true,
        message: 'RSVP system will go live Saturday August 8th'
      };
    }

    const spreadsheet = getSpreadsheet_();
    const guestsSheet = spreadsheet.getSheetByName(RSVP_CONFIG.guestsSheet);
    const guestRows = guestsSheet.getDataRange().getValues().slice(1);
    const responseOffset = 3 + RSVP_INVITED_HEADERS.length;

    const guests = guestRows
      .filter(function (row) { return String(row[1]) === household.id; })
      .map(function (row) {
        const events = {};
        RSVP_EVENTS.forEach(function (eventDefinition, eventIndex) {
          events[eventDefinition.key] = {
            invited: isGuestInvitedToEvent_(row, eventDefinition),
            response: normalizeStoredResponse_(row[responseOffset + eventIndex])
          };
        });

        return {
          id: String(row[0]),
          name: String(row[2]),
          events: events
        };
      });

    if (guests.length === 0) {
      return { ok: false, message: 'This invitation does not have any guests assigned yet.' };
    }

    return {
      ok: true,
      householdName: household.name,
      updatedAt: household.updatedAt || '',
      partyInfo: household.partyInfo,
      guests: guests
    };
  } catch (error) {
    console.error(error);
    return { ok: false, message: 'We could not load this invitation right now. Please try again.' };
  }
}

function saveRsvp(inviteCode, submission) {
  const household = findHouseholdByCode_(inviteCode);
  if (!household) {
    return { ok: false, message: 'We could not find that personal RSVP password.' };
  }

  if (
    !submission ||
    !Array.isArray(submission.guests) ||
    !submission.partyInfo ||
    typeof submission.partyInfo !== 'object'
  ) {
    return { ok: false, message: 'The RSVP submission was incomplete.' };
  }

  const submittedByGuestId = {};
  submission.guests.forEach(function (guest) {
    if (guest && typeof guest.id === 'string') {
      submittedByGuestId[guest.id] = guest;
    }
  });

  const submittedGuestIds = Object.keys(submittedByGuestId);
  if (submittedGuestIds.length !== submission.guests.length) {
    return { ok: false, message: 'The RSVP submission included duplicate or invalid guests.' };
  }

  const partyInfo = submission.partyInfo;
  const contactEmail = validateContactEmail_(partyInfo.contactEmail);
  const contactPhone = validateContactPhone_(partyInfo.contactPhone);
  const householdResponseValues = [
    validateOrlandoHistory_(partyInfo.visitedOrlando),
    safeForSheet_(sanitizeOptionalText_(
      partyInfo.danceFloorSong,
      RSVP_CONFIG.maxShortTextLength
    )),
    safeForSheet_(sanitizeOptionalText_(
      partyInfo.karaokeSong,
      RSVP_CONFIG.maxShortTextLength
    )),
    safeForSheet_(sanitizeOptionalText_(
      partyInfo.anticipatedHotel,
      RSVP_CONFIG.maxShortTextLength
    )),
    validateArrivalMethod_(partyInfo.cabanaBayArrival),
    safeForSheet_(sanitizeOptionalText_(
      partyInfo.accessibilityNeeds,
      RSVP_CONFIG.maxLongTextLength
    )),
    safeForSheet_(sanitizeOptionalText_(
      partyInfo.dietaryRestrictions,
      RSVP_CONFIG.maxLongTextLength
    )),
    safeForSheet_(contactPhone)
  ];

  const spreadsheet = getSpreadsheet_();
  const guestsSheet = spreadsheet.getSheetByName(RSVP_CONFIG.guestsSheet);
  const householdsSheet = spreadsheet.getSheetByName(RSVP_CONFIG.householdsSheet);
  const lock = LockService.getScriptLock();
  const responseOffset = 3 + RSVP_INVITED_HEADERS.length;
  const persistedValueCount = RSVP_EVENTS.length;

  lock.waitLock(15000);
  try {
    const guestValues = guestsSheet.getDataRange().getValues();
    const householdGuestIds = {};
    const pendingUpdates = [];

    for (let rowIndex = 1; rowIndex < guestValues.length; rowIndex += 1) {
      const row = guestValues[rowIndex];
      if (String(row[1]) !== household.id) continue;

      const guestId = String(row[0]);
      householdGuestIds[guestId] = true;
      const submittedGuest = submittedByGuestId[guestId];
      if (!submittedGuest) {
        return { ok: false, message: 'Please answer for every guest listed on this invitation.' };
      }

      const submittedEvents = submittedGuest.events || {};
      const eventResponses = RSVP_EVENTS.map(function (eventDefinition, eventIndex) {
        return validateEventResponse_(
          isGuestInvitedToEvent_(row, eventDefinition) && eventDefinition.collectResponse !== false,
          submittedEvents[eventDefinition.key]
        );
      });

      pendingUpdates.push({
        rowIndex: rowIndex,
        values: eventResponses
      });
    }

    if (pendingUpdates.length === 0) {
      return { ok: false, message: 'No guests were found for this invitation.' };
    }

    const hasUnexpectedGuest = submittedGuestIds.some(function (guestId) {
      return !householdGuestIds[guestId];
    });
    if (hasUnexpectedGuest || submittedGuestIds.length !== pendingUpdates.length) {
      return {
        ok: false,
        message: 'The guest list for this invitation has changed. Please reload and try again.'
      };
    }

    // Validate the complete party before changing any response cells. Then write
    // all guest event responses in one batch so validation failures cannot leave
    // the party partially updated.
    pendingUpdates.forEach(function (update) {
      for (let columnOffset = 0; columnOffset < update.values.length; columnOffset += 1) {
        guestValues[update.rowIndex][responseOffset + columnOffset] = update.values[columnOffset];
      }
    });

    guestsSheet.getRange(
      2,
      responseOffset + 1,
      guestValues.length - 1,
      persistedValueCount
    ).setValues(
      guestValues.slice(1).map(function (row) {
        return row.slice(responseOffset, responseOffset + persistedValueCount);
      })
    );

    householdsSheet.getRange(household.rowNumber, 5).setValue(
      safeForSheet_(contactEmail)
    );

    householdsSheet.getRange(
      household.rowNumber,
      6,
      1,
      1 + RSVP_HOUSEHOLD_FIELDS.length
    ).setValues([[new Date()].concat(householdResponseValues)]);
  } catch (error) {
    console.error(error);
    return {
      ok: false,
      message: error && error.message
        ? error.message
        : 'We could not save the RSVP right now. Please try again.'
    };
  } finally {
    lock.releaseLock();
  }

  return {
    ok: true,
    message: 'Your RSVP has been saved. You can return with the same personal RSVP password to make changes.'
  };
}

function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(RSVP_CONFIG.spreadsheetProperty);

  if (!spreadsheetId) {
    throw new Error('The RSVP spreadsheet has not been configured. Run setupRsvpSystem first.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  ensureHouseholdContactPhoneColumn_(spreadsheet);
  ensureGuestTransportationResponseColumn_(spreadsheet);
  return spreadsheet;
}

function isGuestInvitedToEvent_(row, eventDefinition) {
  const invitationIndex = RSVP_INVITED_HEADERS.indexOf(eventDefinition.invitedHeader);
  return invitationIndex !== -1 && row[3 + invitationIndex] === true;
}

function ensureGuestTransportationResponseColumn_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(RSVP_CONFIG.guestsSheet);
  if (!sheet || sheet.getLastRow() === 0) return;

  const expectedHeader = 'welcome_transportation_response';
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function (header) { return String(header); });
  if (headers.indexOf(expectedHeader) !== -1) return;

  const expectedColumn = RSVP_HEADERS.guests.indexOf(expectedHeader) + 1;
  if (expectedColumn < 1 || lastColumn !== expectedColumn - 1) {
    throw new Error('The Guests sheet could not be upgraded with welcome_transportation_response.');
  }

  sheet.getRange(1, expectedColumn)
    .setValue(expectedHeader)
    .setFontWeight('bold')
    .setBackground('#f5ede0');
}

function ensureHouseholdContactPhoneColumn_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(RSVP_CONFIG.householdsSheet);
  if (!sheet || sheet.getLastRow() === 0) return;

  const expectedHeader = 'contact_phone';
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function (header) { return String(header); });
  const existingIndex = headers.indexOf(expectedHeader);
  if (existingIndex !== -1) return;

  const expectedColumn = RSVP_HEADERS.households.indexOf(expectedHeader) + 1;
  if (expectedColumn < 1 || lastColumn !== expectedColumn - 1) {
    throw new Error('The Households sheet could not be upgraded with contact_phone.');
  }

  sheet.getRange(1, expectedColumn)
    .setValue(expectedHeader)
    .setFontWeight('bold')
    .setBackground('#f5ede0');
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (currentHeaders.join('|') !== headers.join('|')) {
      throw new Error('The "' + name + '" sheet has unexpected column headers.');
    }
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f5ede0');
  return sheet;
}

function findHouseholdByCode_(inviteCode) {
  if (!isValidInviteCodeFormat_(inviteCode)) return null;

  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(RSVP_CONFIG.householdsSheet);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const targetHash = hashInviteCode_(String(inviteCode).trim());
  const rows = sheet.getDataRange().getValues();

  for (let index = 1; index < rows.length; index += 1) {
    if (String(rows[index][3]) === targetHash) {
      return {
        rowNumber: index + 1,
        key: String(rows[index][0]),
        name: String(rows[index][1]),
        id: String(rows[index][2]),
        updatedAt: rows[index][5]
          ? new Date(rows[index][5]).toISOString()
          : '',
        partyInfo: {
          contactEmail: String(rows[index][4] || ''),
          visitedOrlando: normalizeOrlandoHistory_(rows[index][6]),
          danceFloorSong: String(rows[index][7] || ''),
          karaokeSong: String(rows[index][8] || ''),
          anticipatedHotel: String(rows[index][9] || ''),
          cabanaBayArrival: normalizeArrivalMethod_(rows[index][10]),
          accessibilityNeeds: String(rows[index][11] || ''),
          dietaryRestrictions: String(rows[index][12] || ''),
          contactPhone: String(rows[index][13] || '')
        }
      };
    }
  }

  return null;
}

function validateEventResponse_(isInvited, response) {
  if (!isInvited) return '';
  const normalized = String(response || '').toLowerCase();
  if (normalized !== 'yes' && normalized !== 'no') {
    throw new Error('Please select Yes or No for every event.');
  }
  return normalized;
}

function validateOrlandoHistory_(response) {
  const normalized = normalizeOrlandoHistory_(response);
  if (!normalized) {
    throw new Error('Please choose how many people in your party have been to Orlando.');
  }
  return normalized;
}

function normalizeOrlandoHistory_(value) {
  const normalized = String(value || '').toLowerCase();
  return ['all', 'some', 'none'].indexOf(normalized) !== -1 ? normalized : '';
}

function validateArrivalMethod_(value) {
  const normalized = normalizeArrivalMethod_(value);
  if (!normalized) {
    throw new Error('Please select how you plan to arrive at the ceremony on Saturday.');
  }
  return normalized;
}

function normalizeArrivalMethod_(value) {
  const normalized = String(value || '').toLowerCase();
  return ['cabana-bay', 'rideshare', 'drive-park'].indexOf(normalized) !== -1
    ? normalized
    : '';
}

function normalizeStoredResponse_(value) {
  const normalized = String(value || '').toLowerCase();
  return normalized === 'yes' || normalized === 'no' ? normalized : '';
}

function normalizeInviteCode_(value) {
  return String(value || '').trim().toLowerCase();
}

function isValidInviteCodeFormat_(value) {
  return /^[A-Za-z0-9_-]{2,40}$/.test(String(value || '').trim());
}

function sanitizeInviteCode_(value, sheetRow) {
  const code = String(value || '').trim();
  if (!isValidInviteCodeFormat_(code)) {
    throw new Error(
      'invite_code on Guest Setup row ' + sheetRow +
      ' must be 2–40 letters, numbers, hyphens, or underscores with no spaces.'
    );
  }
  return code;
}

function hashInviteCode_(inviteCode) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    normalizeInviteCode_(inviteCode),
    Utilities.Charset.UTF_8
  );
  return digest.map(function (byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function parseBoolean_(value) {
  if (value === true) return true;
  const normalized = String(value || '').trim().toLowerCase();
  return ['true', 'yes', 'y', '1'].indexOf(normalized) !== -1;
}

function validateContactEmail_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized)) {
    throw new Error('Please enter a valid email address for event updates.');
  }
  return normalized;
}

function validateContactPhone_(value) {
  const normalized = String(value || '').trim();
  const digitCount = (normalized.match(/\d/g) || []).length;
  if (normalized.length > 40 || digitCount < 7) {
    throw new Error('Please enter a valid phone number for event updates.');
  }
  return normalized;
}

function sanitizeIdentifier_(value, fieldName, rowNumber) {
  const normalized = String(value || '').trim().toUpperCase();
  if (!/^[A-Z0-9_-]{2,40}$/.test(normalized)) {
    throw new Error(
      fieldName + ' on row ' + rowNumber +
      ' must be 2–40 letters, numbers, underscores, or hyphens.'
    );
  }
  return normalized;
}

function sanitizeText_(value, maxLength, fieldName, rowNumber) {
  const normalized = String(value || '').trim();
  if (!normalized || normalized.length > maxLength) {
    throw new Error(
      fieldName + ' on row ' + rowNumber +
      ' is required and must be no more than ' + maxLength + ' characters.'
    );
  }
  return normalized;
}

function sanitizeEmail_(value, rowNumber) {
  const normalized = String(value || '').trim().toLowerCase();
  if (normalized && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized)) {
    throw new Error('contact_email on row ' + rowNumber + ' is not valid.');
  }
  return normalized;
}

function sanitizeOptionalText_(value, maxLength) {
  const normalized = String(value || '').trim();
  if (normalized.length > maxLength) {
    throw new Error('Dietary notes must be no more than ' + maxLength + ' characters.');
  }
  return normalized;
}

function safeForSheet_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
