/******************************
 * CONFIGURATION
 * Version: 1.3.0
 ******************************/
const CONFIG = {
  calendarId: 'primary',            // 'primary' or your calendar ID

  // Title customization
  useEmoji: true,                    // Add 🎂 emoji to event titles. Only affects default title formats.
  showYearOrAge: false,               // Recurrence on: shows (*YYYY), off: shows (age). Only affects default title formats.
  showAgeOnRecurring: true,         // If true, shows (age) on recurring events instead of (*YYYY)
  
  // Language and localization
  language: 'en',                    // Language code: 'en' (English), 'it' (Italian), 'fr' (French), 'de' (German), 'es' (Spanish)
  titleFormat: '',                   // Title format template (see LANGUAGE_CONFIG for placeholders), e.g. {emoji}{name} ({ageOrYear})

  // Recurrence
  useRecurrence: true,               // Create recurring yearly events
  futureYears: 2,                   // Recurring events end this many years in the future
  pastYears: 1,                      // Recurring events start this many years in the past

  // Reminder settings
  useReminders: true,                // Enable/disable reminders for birthday events
  reminderMinutesBefore: 1440,       // Popup reminder time (in minutes) - only used if useReminders is true
                                     // Common values: 0 = at event time, 60 = 1 hour before, 1440 = 1 day before, 10080 = 1 week before

  // Availability settings
  setTransparency: false,            // Events display you as busy by default (Transparency OPAQUE). Set to true to display you as available (Transparency TRANSPARENT).

  // Cleanup
  cleanupEvents: false,              // ⚠️⚠️⚠️ Deletes all matching birthday events between ±100 years

  // Trigger options
  useTrigger: true,                  // Automatically run on a schedule
  triggerFrequency: 'daily',         // 'daily' or 'hourly'
  triggerHour: 4,                    // If 'daily', the hour of day to run (0–23)

  // Script identification
  scriptKey: 'CREATED_BY_Auto-Birthdays', // Unique identifier for events created by this script; customize if desired

  // Contact label filtering (optional)
  useLabels: false,                  // Enable filtering contacts by labels
  contactLabels: [],                 // Array of contact label IDs to include (e.g. ['abc123'])

  // Month filtering (optional)
  useMonthFilter: false,             // Enable filtering contacts by birth month
  filterMonths: []                   // Array of months to include (1-12), e.g. [1, 2, 4] for Jan, Feb, Apr
};

/**
 * Language configuration for localized event titles and descriptions
 */
const LANGUAGE_CONFIG = {
  en: {
    titleFormats: {
      'default': '{emoji}{name} - {age} {years}'
    },
    descriptionFormats: {
      'default': '{name} turns {age} {years} old!',
      'unknownYear': "{name}'s birthday!",
      'recurring': "{name}'s birthday! (Born: {birthYear})"
    },
    terms: {
      'age': 'age',
      'years': 'years',
      'year': 'year',
      'birthday': 'birthday',
      'happyBirthday': 'Happy Birthday'
    }
  },
  it: {
    titleFormats: {
      'default': '{emoji}Compleanno di {name} - {age} {years}'
    },
    descriptionFormats: {
      'default': '{name} compie {age} {years}!',
      'unknownYear': '{name} compie gli anni!',
      'recurring': '{name} compie gli anni! (Anno di nascita: {birthYear})'
    },
    terms: {
      'age': 'età',
      'years': 'anni',
      'year': 'anno',
      'birthday': 'compleanno',
      'happyBirthday': 'Buon Compleanno'
    }
  },
  fr: {
    titleFormats: {
      'default': '{emoji}Anniversaire de {name} - {age} {years}'
    },
    descriptionFormats: {
      'default': '{name} a {age} {years} !',
      'unknownYear': '{name} fête son anniversaire !',
      'recurring': '{name} fête son anniversaire ! (Année de naissance : {birthYear})'
    },
    terms: {
      'age': 'âge',
      'years': 'ans',
      'year': 'an',
      'birthday': 'anniversaire',
      'happyBirthday': 'Joyeux Anniversaire'
    }
  },
  de: {
    titleFormats: {
      'default': '{emoji}Geburtstag von {name} - {age} {years}'
    },
    descriptionFormats: {
      'default': '{name} wird {age} {years} alt!',
      'unknownYear': '{name} hat Geburtstag!',
      'recurring': '{name} hat Geburtstag! (Jahrgang: {birthYear})'
    },
    terms: {
      'age': 'Alter',
      'years': 'Jahre',
      'year': 'Jahr',
      'birthday': 'Geburtstag',
      'happyBirthday': 'Alles Gute zum Geburtstag'
    }
  },
  es: {
    titleFormats: {
      'default': '{emoji}Cumpleaños de {name} - {age} {years}'
    },
    descriptionFormats: {
      'default': '¡{name} cumple {age} {years}!',
      'unknownYear': '¡Es el cumpleaños de {name}!',
      'recurring': '¡Es el cumpleaños de {name}! (Año de nacimiento: {birthYear})'
    },
    terms: {
      'age': 'edad',
      'years': 'años',
      'year': 'año',
      'birthday': 'cumpleaños',
      'happyBirthday': 'Feliz Cumpleaños'
    }
  }
};


/**
 * Main function to create/update birthday events.
 * Also manages the optional time-based trigger.
 */
function loopThroughContacts() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    Logger.log('Another Auto-Birthdays run is already active; skipping this run.');
    return;
  }

  try {
    return syncBirthdayEvents();
  } finally {
    lock.releaseLock();
  }
}

function syncBirthdayEvents() {
  if (CONFIG.useTrigger) {
    ensureTriggerExists();
  } else {
    removeTriggerIfExists();
  }

  const connections = getAllContacts();
  // Validate reminder configuration
  if (CONFIG.useReminders && (typeof CONFIG.reminderMinutesBefore !== 'number' || CONFIG.reminderMinutesBefore < 0)) {
    Logger.log("⚠️ Invalid reminder configuration: reminderMinutesBefore must be a non-negative number when useReminders is true.");
    return;
  }

  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);

  if (!calendar) {
    Logger.log("⚠️ Calendar not found.");
    return;
  }

  if (CONFIG.cleanupEvents) {
    Logger.log("🧹 Starting cleanup of old birthday events...");
    cleanupOldBirthdayEvents(calendar, connections);
    Logger.log("🎉 Cleanup completed! Script will now exit without recreating events.");
    return;
  }

  Logger.log("📊 Starting birthday event processing...");

  const currentYear = new Date().getFullYear();
  const startDate = new Date(currentYear - CONFIG.pastYears - 1, 0, 1);
  const endDate = new Date(currentYear + CONFIG.futureYears + 1, 11, 31);
  const allEvents = new Set(calendar.getEvents(startDate, endDate));
  const eventIndex  = buildBirthdayIndex(allEvents);

  removeUnindexedEvents(allEvents, eventIndex);

  // Initialize counters for processing report
  let totalContacts = connections.length;
  let contactsWithBirthdays = 0;
  let processedContacts = 0;
  let eventsCreated = 0;
  let eventsUpdated = 0;
  let skippedByLabelFilter = 0;
  let skippedByMonthFilter = 0;
  let skippedInvalidBirthdays = 0;

  for (const person of connections) {
    // Check if label filtering is enabled and if this contact has the required labels
    if (CONFIG.useLabels && !hasRequiredLabel(person, CONFIG.contactLabels)) {
      skippedByLabelFilter++;
      continue; // Skip this contact if it doesn't have the required labels
    }

    // Check if month filtering is enabled and if this contact's birthday month matches
    if (CONFIG.useMonthFilter && !hasMatchingBirthMonth(person, CONFIG.filterMonths)) {
      skippedByMonthFilter++;
      continue; // Skip this contact if its birthday month doesn't match the filter
    }

    const birthdayData = person.birthdays?.find(b => b.date);
    if (birthdayData) {
      contactsWithBirthdays++;
      try {
        const result = updateOrCreateBirthDayEvent(person, birthdayData, calendar, allEvents, eventIndex);
        if (result === 'created') {
          processedContacts++;
          eventsCreated++;
        } else if (result === 'updated') {
          processedContacts++;
          eventsUpdated++;
        } else if (result === 'skipped_existing') {
          processedContacts++;
          // Event already exists and is correct - no action needed
        } else if (result === 'skipped_invalid') {
          skippedInvalidBirthdays++;
        }
      } catch (error) {
        const contactName = getContactName(person);
        Logger.log(`❌ Error processing ${contactName}: ${error}`);
        skippedInvalidBirthdays++;
      }
    }
  }

  // Log processing summary report
  Logger.log("📊 PROCESSING SUMMARY REPORT:");
  Logger.log(`📞 Total contacts retrieved: ${totalContacts}`);
  Logger.log(`🎂 Contacts with birthday data: ${contactsWithBirthdays}`);
  Logger.log(`✅ Successfully processed contacts: ${processedContacts}`);
  Logger.log(`🆕 Events created: ${eventsCreated}`);
  Logger.log(`🔄 Events updated: ${eventsUpdated}`);
  
  const eventsAlreadyCorrect = processedContacts - eventsCreated - eventsUpdated;
  if (eventsAlreadyCorrect > 0) {
    Logger.log(`✓ Events already up-to-date: ${eventsAlreadyCorrect}`);
  }
  if (CONFIG.useLabels) {
    Logger.log(`🏷️ Contacts skipped by label filter: ${skippedByLabelFilter}`);
  }
  
  if (CONFIG.useMonthFilter) {
    Logger.log(`📅 Contacts skipped by month filter: ${skippedByMonthFilter}`);
  }
  
  if (skippedInvalidBirthdays > 0) {
    Logger.log(`⚠️ Contacts with invalid birthday data: ${skippedInvalidBirthdays}`);
  }
  
  const contactsWithoutBirthdays = totalContacts - contactsWithBirthdays - skippedByLabelFilter - skippedByMonthFilter;
  if (contactsWithoutBirthdays > 0) {
    Logger.log(`📝 Contacts without birthday data: ${contactsWithoutBirthdays}`);
  }
  
  // Show reminder configuration info
  if (CONFIG.useReminders) {
    const reminderText = CONFIG.reminderMinutesBefore === 0 ? "at event time" :
                        CONFIG.reminderMinutesBefore === 60 ? "1 hour before" :
                        CONFIG.reminderMinutesBefore === 1440 ? "1 day before" :
                        CONFIG.reminderMinutesBefore === 10080 ? "1 week before" :
                        `${CONFIG.reminderMinutesBefore} minutes before`;
    Logger.log(`⏰ Reminders enabled: ${reminderText}`);
  } else {
    Logger.log(`⏰ Reminders disabled`);
  }
  
  Logger.log("🎉 Processing completed successfully!");
}

/**
 * Language configuration for localized event titles and descriptions.
 * Supports: English (en), Italian (it), French (fr), German (de), Spanish (es)
 * 
 * For configuration examples and full documentation, see README.md
 */

/**
 * Format localized text for a birthday event
 * @param {string} formatTemplate - The format template
 * @param {string} contactName - The name of the person
 * @param {number|null} age - The age of the person (null if unknown)
 * @param {number|null} birthYear - The birth year (null if unknown)
 * @param {boolean} showYear - Whether to show birth year instead of age
 * @param {boolean} isRecurring - Whether this is a recurring event
 * @param {boolean} usingCustomFormat - Whether a custom title format is being used
 * @returns {string} The formatted text
 */
function formatLocalizedText(formatTemplate, contactName, age, birthYear, showYear, isRecurring, usingCustomFormat = false) {
  const langConfig = LANGUAGE_CONFIG[CONFIG.language] || LANGUAGE_CONFIG['en'];
  
  // Build replacement values
  // useEmoji only applies to default formats; custom formats control emoji via {emoji} placeholder
  const emoji = (usingCustomFormat || CONFIG.useEmoji) ? '🎂 ' : '';
  let ageOrYear = '';
  let ageText = '';
  
  // Always populate age-related variables if we have the data (regardless of showYearOrAge setting)
  // The showYearOrAge setting only affects default formats, not custom formats
  if (birthYear) {
    if (showYear) {
      ageOrYear = `${birthYear}`;
      ageText = `${birthYear}`;
    } else if (age !== null) {
      const yearWord = age === 1 ? langConfig.terms.year : langConfig.terms.years;
      ageOrYear = age.toString();
      ageText = `${age} ${yearWord}`;
    }
  }
  
  // Replace placeholders in the format template
  let text = formatTemplate
    .replace(/{emoji}/g, emoji)
    .replace(/{name}/g, contactName)
    .replace(/{ageOrYear}/g, ageOrYear)
    .replace(/{age}/g, age !== null ? age.toString() : '')
    .replace(/{ageText}/g, ageText)
    .replace(/{birthYear}/g, birthYear ? birthYear.toString() : '')
    .replace(/{years}/g, age === 1 ? langConfig.terms.year : (age !== null ? langConfig.terms.years : ''))  // Automatically singular/plural
    .replace(/{year}/g, age === 1 ? langConfig.terms.year : '')  // Singular form only for age 1 (for backwards compatibility)
    .replace(/{birthday}/g, langConfig.terms.birthday);
  
  // Clean up empty age/year information - remove parentheses, dashes, etc. around empty values
  text = text
    .replace(/\s*\(\s*\)\s*/g, '')        // Remove empty parentheses like " () "
    .replace(/\s*\(\s*\*\s*\)\s*/g, '')   // Remove empty year parentheses like " (*) "
    .replace(/\s*-\s*\*?\s*$/g, '')       // Remove trailing " - " or " - *"
    .replace(/\s*-\s*$/g, '')             // Remove trailing " - "
    .replace(/\s+/g, ' ');                // Clean up multiple spaces
    
  return text.trim();
}

/**
 * Generate a localized title for a birthday event
 * @param {string} contactName - The name of the person
 * @param {number|null} age - The age of the person (null if unknown)
 * @param {number|null} birthYear - The birth year (null if unknown)  
 * @param {boolean} showYear - Whether to show birth year instead of age
 * @param {boolean} isRecurring - Whether this is a recurring event
 * @returns {string} The formatted title
 */
function generateLocalizedTitle(contactName, age, birthYear, showYear, isRecurring) {
  const langConfig = LANGUAGE_CONFIG[CONFIG.language] || LANGUAGE_CONFIG['en'];
  
  // Use titleFormat if provided, otherwise fall back to language default
  const formatTemplate = CONFIG.titleFormat || langConfig.titleFormats['default'];
  const usingCustomFormat = CONFIG.titleFormat && CONFIG.titleFormat.trim() !== '';

  return formatLocalizedText(formatTemplate, contactName, age, birthYear, showYear, isRecurring, usingCustomFormat);
}

/**
 * Generate a localized description for a birthday event
 * @param {string} contactName - The name of the person
 * @param {number|null} age - The age of the person (null if unknown)
 * @param {number|null} birthYear - The birth year (null if unknown)
 * @param {boolean} showYear - Whether to show birth year instead of age
 * @param {boolean} isRecurring - Whether this is a recurring event
 * @returns {string} The formatted description
 */
function generateLocalizedDescription(contactName, age, birthYear, showYear, isRecurring) {
  const langConfig = LANGUAGE_CONFIG[CONFIG.language] || LANGUAGE_CONFIG['en'];

  const formatKey = !birthYear ? 'unknownYear' : (isRecurring ? 'recurring' : 'default');
  const formatTemplate = langConfig.descriptionFormats[formatKey];

  return formatLocalizedText(formatTemplate, contactName, age, birthYear, showYear, isRecurring);
}

/**
 * Build a quick‑lookup map keying events by "contactId|month|day|year".
 * Significantly reduces runtime from O(contacts × events) to O(contacts + events).
 */
function buildBirthdayIndex(events) {
  const map = new Map();
  const duplicateKeys = new Set();
  for (const ev of events) {
    if (!ev.isAllDayEvent()) continue;
    const contactId = tryGetTag(ev, 'contactId');
    if (!contactId) continue;
    const d   = ev.getStartTime();
    const key = `${contactId}|${d.getMonth()}|${d.getDate()}|${d.getFullYear()}`;
    if (duplicateKeys.has(key)) continue;
    if (map.has(key)) {
      map.delete(key);
      duplicateKeys.add(key);
      continue;
    }
    map.set(key, ev);
  }
  return map;
}

/**
 * Delete events excluded from the index in this script-owned calendar.
 * Excluding an occurrence excludes its whole series, including indexed dates.
 * Mutates the supplied event Set and index in place.
 */
function removeUnindexedEvents(events, eventIndex) {
  const retainedEvents = new Set(eventIndex.values());
  const seriesById = new Map();
  const seriesIdsToDelete = new Set();

  for (const event of events) {
    if (event.isRecurringEvent()) {
      const series = event.getEventSeries();
      const seriesId = series.getId();
      if (!seriesById.has(seriesId)) seriesById.set(seriesId, { series, events: [] });
      seriesById.get(seriesId).events.push(event);
      if (!retainedEvents.has(event)) seriesIdsToDelete.add(seriesId);
    }
  }

  for (const seriesId of seriesIdsToDelete) {
    const { series, events: seriesEvents } = seriesById.get(seriesId);
    series.deleteEventSeries();
    for (const event of seriesEvents) events.delete(event);
  }

  for (const event of events) {
    if (retainedEvents.has(event)) continue;

    event.deleteEvent();
    events.delete(event);
  }

  for (const [key, event] of eventIndex) {
    if (!events.has(event)) eventIndex.delete(key);
  }
}

/**
 * Check if a contact has any of the specified labels.
 * @param {object} person - The contact person object
 * @param {string[]} labelIds - Array of contact label IDs to check for (just the ID part, not the full resource name)
 * @returns {boolean} - True if contact has any of the specified labels
 */
function hasRequiredLabel(person, labelIds) {
  if (!labelIds || labelIds.length === 0) {
    return true; // No labels specified, so all contacts match
  }

  if (!person.memberships || person.memberships.length === 0) {
    return false; // Contact has no group memberships
  }

  // Check if any membership matches any of the specified label IDs
  for (const membership of person.memberships) {
    if (membership.contactGroupMembership && membership.contactGroupMembership.contactGroupId) {
      const contactGroupId = membership.contactGroupMembership.contactGroupId;
      
      // Check if this contactGroupId contains any of our target labels
      // contactGroupId format is typically "contactGroups/abc122"
      for (const labelId of labelIds) {
        if (contactGroupId.includes(labelId)) {
          return true;
        }
      }
    }
  }

  return false;
}

/**
 * Check if a contact's birthday month matches any of the specified filter months.
 * @param {object} person - The contact person object
 * @param {number[]} filterMonths - Array of months to include (1-12)
 * @returns {boolean} - True if contact's birthday month matches any filter month
 */
function hasMatchingBirthMonth(person, filterMonths) {
  if (!filterMonths || filterMonths.length === 0) {
    return true; // No month filter specified, so all contacts match
  }

  const birthdayData = person.birthdays?.find(b => b.date);
  if (!birthdayData || !birthdayData.date || typeof birthdayData.date.month !== 'number') {
    return false; // No valid birthday data
  }

  const birthMonth = birthdayData.date.month;
  return filterMonths.includes(birthMonth);
}

/**
 * Get all Google Contacts using pagination.
 */
function getAllContacts() {
  const connections = [];
  let nextPageToken;

  do {
    const reqOpts = {
      personFields: 'names,birthdays,memberships',
      sortOrder: 'LAST_NAME_ASCENDING',
      pageSize: 100,
      pageToken: nextPageToken
    };

    const response = People.People.Connections.list('people/me', reqOpts);
    connections.push(...(response.connections || []));
    nextPageToken = response.nextPageToken;
  } while (nextPageToken);

  return connections;
}

/**
 * Create or update a calendar event for a contact's birthday.
 * @returns {string} - Operation result: 'created', 'updated', 'skipped_existing', or 'skipped_invalid'
 */
function updateOrCreateBirthDayEvent(person, birthdayRaw, calendar, allEvents, eventIndex) {
  const contactName = getContactName(person);
  const contactId = person.resourceName;

  // Extract actual date from birthday object
  const birthdayDate = birthdayRaw.date;
  if (!birthdayDate || typeof birthdayDate.day !== 'number' || typeof birthdayDate.month !== 'number') {
    Logger.log(`⚠️ Skipping ${contactName} due to invalid birthdayDate: ${JSON.stringify(birthdayRaw)}`);
    return 'skipped_invalid';
  }

  const currentYear = new Date().getFullYear();
  const startYear = currentYear - CONFIG.pastYears;

  const month = parseInt(birthdayDate.month, 10) - 1;
  const day = parseInt(birthdayDate.day, 10);

  const firstBirthdayYear = birthdayDate.year ? birthdayDate.year + 1 : startYear;
  const birthdayStartDate = new Date(Math.max(startYear, firstBirthdayYear), month, day);
  const birthdayDateThisYear = new Date(currentYear, month, day);

  if (isNaN(birthdayStartDate.getTime()) || isNaN(birthdayDateThisYear.getTime())) {
    Logger.log(`⚠️ Invalid date for ${contactName}: ${birthdayDate.month}-${birthdayDate.day}`);
    return 'skipped_invalid';
  }
  
  // Build expected title using localized function
  const age = birthdayDate.year ? currentYear - birthdayDate.year : null;
  const showYear = CONFIG.useRecurrence && !CONFIG.showAgeOnRecurring;
  const expectedTitle = generateLocalizedTitle(
    contactName,
    age,
    birthdayDate.year,
    showYear,
    CONFIG.useRecurrence
  );

  // Generate expected description using localized function
  const expectedDescription = generateLocalizedDescription(
    contactName,
    age,
    birthdayDate.year,
    showYear,
    CONFIG.useRecurrence
  );

  const usingIndividualEvents = !CONFIG.useRecurrence || (CONFIG.showAgeOnRecurring && birthdayDate.year);

  // Find and update related events
  const relatedEvents = findBirthdayEvents(allEvents, contactId, month, day);
  const deletedSeriesIds = new Set();
  const deletedEventIds = new Set();
  let correctEventExists = false;
  let eventsWereDeleted = false;

  for (const event of relatedEvents) {
    const title = event.getTitle();
    let description;
    try {
      description = event.getDescription() || '';
    } catch (e) {
      description = '';
    }
    
    // For individual events, generate the expected title for the event's specific year
    let eventExpectedTitle = expectedTitle;
    let eventExpectedDescription = expectedDescription;
    if (usingIndividualEvents) {
      const eventYear = event.getStartTime().getFullYear();
      const eventAge = birthdayDate.year ? eventYear - birthdayDate.year : null;
      eventExpectedTitle = generateLocalizedTitle(
        contactName,
        eventAge,
        birthdayDate.year,
        false, // Always show age for individual events
        false  // Individual events are not recurring
      );
      eventExpectedDescription = generateLocalizedDescription(
        contactName,
        eventAge,
        birthdayDate.year,
        false, // Always show age for individual events
        false  // Individual events are not recurring
      );
    }
    
    const isTitleOutdated = title !== eventExpectedTitle;
    const isDescriptionOutdated = description !== eventExpectedDescription;
    const isNotAllDay = !event.isAllDayEvent();
    
    // When using individual events (recurrence disabled or age display), events should NOT be recurring
    // When not using individual events, recurrence should match CONFIG.useRecurrence
    const eventIsRecurring = event.isRecurringEvent && event.isRecurringEvent();
    const isRecurrenceMismatch = usingIndividualEvents 
      ? eventIsRecurring  // Individual events should not be recurring
      : CONFIG.useRecurrence !== eventIsRecurring;  // Otherwise match config
    
    const isNotFromScript = !isEventCreatedByScript(event);
    const hasIncorrectReminders = !hasCorrectReminders(event);

    // When using individual events, we need to delete any existing recurring events
    const needsConversionToIndividual = usingIndividualEvents && eventIsRecurring;

    if (isTitleOutdated || isDescriptionOutdated || isNotAllDay || isRecurrenceMismatch || isNotFromScript || hasIncorrectReminders || needsConversionToIndividual) {
      eventsWereDeleted = true;

      // Log the reason for deletion
      Logger.log(`🗑️ Deleting outdated event: ${title} | Reasons: ${isTitleOutdated ? 'Title ' : ''}${isDescriptionOutdated ? 'Description ' : ''}${isNotAllDay ? 'NotAllDay ' : ''}${isRecurrenceMismatch ? 'RecurrenceMismatch ' : ''}${isNotFromScript ? 'NotFromScript ' : ''}${hasIncorrectReminders ? 'IncorrectReminders ' : ''}${needsConversionToIndividual ? 'NeedsConversionToIndividual ' : ''}`);

      try {
        if (event.isRecurringEvent && event.isRecurringEvent()) {
          const series = event.getEventSeries();
          const seriesId = series.getId();
          if (deletedSeriesIds.has(seriesId)) {
            break;
          }
          deletedSeriesIds.add(seriesId);
          series.deleteEventSeries();
          Logger.log(`🗑️ Deleted outdated recurring series: ${title}`);
        } else {
          const eventId = event.getId();
          deletedEventIds.add(eventId);
          event.deleteEvent();
          Logger.log(`🗑️ Deleted outdated event: ${title}`);
        }
        const eventDate = event.getStartTime();
        const eventKey = `${contactId}|${eventDate.getMonth()}|${eventDate.getDate()}|${eventDate.getFullYear()}`;
        if (eventIndex.get(eventKey) === event) eventIndex.delete(eventKey);
      } catch (e) {
        Logger.log(`❌ Error deleting event: ${title} → ${e}`);
      }
    } else {
      // Only set correctEventExists for non-individual events
      // For individual events, we need to check if ALL years exist, not just one
      if (!usingIndividualEvents) {
        correctEventExists = true;
      }
    }
  }

  // For non-individual events, if a correct event exists, skip
  if (correctEventExists && !usingIndividualEvents) {
    return 'skipped_existing'; // Correct event already exists
  }

  // Create events - individual yearly events if showing age on recurring, or recurrence is disabled, otherwise use recurrence
  if (usingIndividualEvents) {
    // Create individual events for each year and show correct age (if known)
    // Only create years that don't already exist (handles timeout recovery)
    let createdCount = 0;
    for (let year = startYear; year <= currentYear + CONFIG.futureYears; year++) {
      // Do not create events before the first birthday
      if (birthdayDate.year && year - birthdayDate.year < 1) continue;

      const yearAge = birthdayDate.year ? year - birthdayDate.year : null;
      const yearBirthdayDate = new Date(year, month, day);
      
      // Build title with age for this specific year using localized function
      const yearTitle = generateLocalizedTitle(
        contactName,
        yearAge,
        birthdayDate.year,
        false, // Always show age for individual events
        false  // Individual events are not recurring
      );
      const yearDescription = generateLocalizedDescription(
        contactName,
        yearAge,
        birthdayDate.year,
        false, // Always show age for individual events
        false  // Individual events are not recurring
      );
      
      // Check if this specific year already exists and is correct
      const yearKey = `${contactId}|${month}|${day}|${year}`;
      const existingYearEvent = eventIndex.get(yearKey);
      const existingYearEventId = existingYearEvent ? existingYearEvent.getId() : null;
      if (existingYearEvent &&
          (existingYearEventId && !deletedEventIds.has(existingYearEventId)) &&
          isEventCreatedByScript(existingYearEvent) && 
          (existingYearEvent.getTitle() || '') === yearTitle &&
          (existingYearEvent.getDescription() || '') === yearDescription && 
          hasCorrectReminders(existingYearEvent)) {
        // This year already has a correct event, skip it
        continue;
      }
      
      const event = calendar.createAllDayEvent(
        yearTitle,
        yearBirthdayDate,
        { description: yearDescription }
      );
      event.setTag('contactId', contactId);
      event.setTag('scriptKey', CONFIG.scriptKey);
      if (CONFIG.useReminders) {
        event.addPopupReminder(CONFIG.reminderMinutesBefore);
      }
      if (CONFIG.setTransparency) {
        event.setTransparency(CalendarApp.EventTransparency.TRANSPARENT);
      }
      eventIndex.set(yearKey, event);
      createdCount++;
      Logger.log(`🎁 Created individual event: ${yearTitle} [${yearBirthdayDate.toDateString()}]`);
    }
    if (createdCount === 0) {
      return 'skipped_existing'; // All individual events already existed
    }
  } else {
    // Create standard recurring event
    const recurrence = CalendarApp.newRecurrence()
      .addYearlyRule()
      .until(new Date(currentYear + CONFIG.futureYears, 11, 31));

    const eventSeries = calendar.createAllDayEventSeries(
      expectedTitle,
      birthdayStartDate,
      recurrence,
      { description: expectedDescription }
    );
    eventSeries.setTag('contactId', contactId);
    eventSeries.setTag('scriptKey', CONFIG.scriptKey);
    if (CONFIG.useReminders) {
      eventSeries.addPopupReminder(CONFIG.reminderMinutesBefore);
    }
    if (CONFIG.setTransparency) {
      eventSeries.setTransparency(CalendarApp.EventTransparency.TRANSPARENT);
    }
    const seriesKey = `${contactId}|${month}|${day}|${birthdayStartDate.getFullYear()}`;
    eventIndex.set(seriesKey, eventSeries);
    Logger.log(`🎉 Created RECURRING event: ${expectedTitle} [starts ${birthdayStartDate.toDateString()}]`);
  }
  
  return eventsWereDeleted ? 'updated' : 'created'; // Successfully processed the birthday event
}

/**
 * Check if an event was created by this script by looking for its script key tag.
 */
function isEventCreatedByScript(event) {
  return tryGetTag(event, 'scriptKey') === CONFIG.scriptKey;
}

/**
 * Safely get an event tag, returning null when it is missing or unavailable.
 */
function tryGetTag(event, tagName) {
  try {
    return event.getTag(tagName) || null;
  } catch (e) {
    return null;
  }
}

/**
 * Check if an event has the correct reminder settings according to current configuration.
 * @param {CalendarEvent} event - The calendar event to check
 * @returns {boolean} - True if reminders match current CONFIG settings
 */
function hasCorrectReminders(event) {
  try {
    const reminders = event.getPopupReminders();
    
    if (!CONFIG.useReminders) {
      // If reminders are disabled, event should have no reminders
      return reminders.length === 0;
    } else {
      // If reminders are enabled, event should have exactly one reminder with correct timing
      return reminders.length === 1 && reminders[0] === CONFIG.reminderMinutesBefore;
    }
  } catch (e) {
    // If we can't get reminders info, assume it's incorrect to be safe
    return false;
  }
}

/**
 * Extract name from contact.
 */
function getContactName(person) {
  if (person.names && person.names.length > 0) {
    return person.names[0].displayName ||
      `${person.names[0].givenName} ${person.names[0].familyName}`.trim();
  }
  return "Unknown";
}

/**
 * Check if birthday event already exists.
 */
function findBirthdayEvents(allEvents, contactId, month, day) {
  return [...allEvents].filter(ev => {
    const eventContactId = tryGetTag(ev, 'contactId');
    const d     = ev.getStartTime();
    return ev.isAllDayEvent() &&
           (eventContactId === contactId && (d.getMonth() === month && d.getDate() === day));
  });
}

/**
 * Get upcoming birthday (this year or next).
 */
function calculateNextBirthday(birthdayDate) {
  const day = birthdayDate.day;
  const month = birthdayDate.month - 1;
  const today = new Date();
  let year = today.getFullYear();
  const birthdayThisYear = new Date(year, month, day);
  if (today >= birthdayThisYear) {
    year++;
  }
  return new Date(year, month, day);
}

/**
 * Cleans birthdayEvents in the future and past 100 years
 */
function cleanupOldBirthdayEvents(calendar, allContacts) {
  const currentYear = new Date().getFullYear();
  const startYear = currentYear - 100;
  const endYear = currentYear + 100;

  const startDate = new Date(startYear, 0, 1);
  const endDate = new Date(endYear, 11, 31);
  const allEvents = calendar.getEvents(startDate, endDate);
  
  Logger.log(`🧹 Cleanup started between: ${startDate.toDateString()} - ${endDate.toDateString()}`);

  // Initialize cleanup counters
  let totalEventsScanned = allEvents.length;
  let candidateEventsFound = 0;
  let recurringSeriesDeleted = 0;
  let singleEventsDeleted = 0;
  let deleteFailures = 0;
  let retryAttempts = 0;
  let retrySuccesses = 0;

  // Collect valid contacts with birthday info
  const contactBirthdays = allContacts
    .filter(person => {
      // Apply label filtering if enabled
      if (CONFIG.useLabels && !hasRequiredLabel(person, CONFIG.contactLabels)) {
        return false;
      }
      return true;
    })
    .map(person => {
      if (!person.birthdays) return null;
      const bday = person.birthdays.find(b => b.date);
      if (!bday) return null;

      return {
        contactId: person.resourceName,
        day: bday.date.day,
        month: bday.date.month - 1,
      };
    })
    .filter(Boolean); // remove nulls

  const deletedSeriesIds = new Set();
  const eventsToDelete = [];

  // Phase 1: Scan events and mark candidates for deletion
  Logger.log(`🔍 Scanning ${allEvents.length} events to identify birthday events for cleanup...`);

  for (const event of allEvents) {
    const title = event.getTitle();
    const start = event.getStartTime();
    const isAllDay = event.isAllDayEvent();
    const isFromScript = isEventCreatedByScript(event);
    const eventContactId = tryGetTag(event, 'contactId');

    for (const contact of contactBirthdays) {
      const isContactMatch = eventContactId === contact.contactId;
      const isBirthdayDateMatch =
        start.getDate() === contact.day &&
        start.getMonth() === contact.month;

      const shouldDelete =
        isFromScript && // Only delete events created by this script
        (isAllDay && isContactMatch && isBirthdayDateMatch);

      if (!shouldDelete) continue;

      candidateEventsFound++;

      // Check if it's a recurring event and already handled
      if (event.isRecurringEvent && event.isRecurringEvent()) {
        try {
          const series = event.getEventSeries();
          const seriesId = series.getId();
          if (deletedSeriesIds.has(seriesId)) {
            break;
          }
          deletedSeriesIds.add(seriesId);
          eventsToDelete.push({ event: series, isSeries: true });
          Logger.log(`📋 Marked recurring series for deletion: ${title}`);
          break;
        } catch (e) {
          retryAttempts++;
          Logger.log(`⚠️ Failed to resolve recurring series: ${title} — ${e}`);
          Logger.log(`⚠️ Waiting for rate limiter and safely retrying`);
          Utilities.sleep(5000);
          try {
            const series = event.getEventSeries();
            const seriesId = series.getId();
            if (deletedSeriesIds.has(seriesId)) {
              retrySuccesses++;
              Logger.log(`✅ Retry successful - series marked for deletion: ${title}`);
              break;
            }
            deletedSeriesIds.add(seriesId);
            eventsToDelete.push({ event: series, isSeries: true });
            retrySuccesses++;
            Logger.log(`✅ Retry successful - marked recurring series for deletion: ${title}`);
            break;
          } catch (e) {
            Logger.log(`⚠️ Second attempt failed — skipping entry: ${e}`);
          }
        }
      } else {
        eventsToDelete.push({ event, isSeries: false });
        Logger.log(`📋 Marked single event for deletion: ${title} on ${event.getStartTime().toDateString()}`);
        break;
      }
    }
  }

  // Phase 2: Execute deletions for all marked events
  Logger.log(`🗑️ Starting deletion of ${eventsToDelete.length} marked events (${deletedSeriesIds.size} recurring series, ${eventsToDelete.length - deletedSeriesIds.size} single events)`);
  
  for (const { event, isSeries } of eventsToDelete) {
    try {
      if (isSeries) {
        event.deleteEventSeries(); // Deletes the entire series
        recurringSeriesDeleted++;
        Logger.log(`🧹 Deleted recurring series → ${event.getTitle()}`);
      } else {
        event.deleteEvent();
        singleEventsDeleted++;
        Logger.log(`🧹 Deleted single event → ${event.getTitle()} on ${event.getStartTime().toDateString()}`);
      }
    } catch (e) {
      deleteFailures++;
      Logger.log(`❌ Failed to delete event → ${event.getTitle()} | Reason: ${e}`);
    }
  }

  // Log comprehensive cleanup summary report
  Logger.log("🧹 CLEANUP SUMMARY REPORT:");
  Logger.log(`📅 Events scanned in date range: ${totalEventsScanned}`);
  Logger.log(`🎯 Birthday events found for cleanup: ${candidateEventsFound}`);
  Logger.log(`🔄 Recurring series deleted: ${recurringSeriesDeleted}`);
  Logger.log(`📅 Single events deleted: ${singleEventsDeleted}`);
  
  const totalSuccessfulDeletions = recurringSeriesDeleted + singleEventsDeleted;
  Logger.log(`✅ Total successful deletions: ${totalSuccessfulDeletions}`);
  
  if (deleteFailures > 0) {
    Logger.log(`❌ Failed deletions: ${deleteFailures}`);
  }
  
  if (retryAttempts > 0) {
    Logger.log(`🔄 Retry attempts made: ${retryAttempts}`);
    Logger.log(`✅ Successful retries: ${retrySuccesses}`);
  }
  
  Logger.log(`👥 Contacts with birthdays considered: ${contactBirthdays.length}`);
  Logger.log("🎉 Cleanup completed successfully!");
}

/**
 * Ensure a correct time-based trigger exists for loopThroughContacts.
 * Uses PropertiesService to track trigger configuration and detect changes.
 */
function ensureTriggerExists() {
  const triggers = ScriptApp.getProjectTriggers();
  const properties = PropertiesService.getScriptProperties();
  let existingTrigger = null;
  
  // Find existing trigger for loopThroughContacts
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'loopThroughContacts') {
      existingTrigger = trigger;
      break;
    }
  }
  
  // Check if existing trigger matches current configuration
  if (existingTrigger) {
    const storedFrequency = properties.getProperty('triggerFrequency');
    const storedHour = properties.getProperty('triggerHour');
    
    // Compare stored configuration with current CONFIG
    const frequencyMatches = storedFrequency === CONFIG.triggerFrequency;
    const hourMatches = CONFIG.triggerFrequency === 'hourly' || storedHour === String(CONFIG.triggerHour);
    
    if (frequencyMatches && hourMatches) {
      Logger.log(`✓ Trigger already exists and is up-to-date (${CONFIG.triggerFrequency}${CONFIG.triggerFrequency === 'daily' ? ' at ' + CONFIG.triggerHour + ':00' : ''})`);
      return;
    }
    
    // Configuration changed - delete old trigger
    ScriptApp.deleteTrigger(existingTrigger);
    Logger.log(`🗑️ Removed outdated trigger (was: ${storedFrequency}${storedFrequency === 'daily' ? ' at ' + storedHour + ':00' : ''})`);
  }

  // Create new trigger with current configuration
  const builder = ScriptApp.newTrigger('loopThroughContacts').timeBased();
  if (CONFIG.triggerFrequency === 'hourly') {
    builder.everyHours(1);
    Logger.log("✅ Created hourly trigger");
  } else {
    builder.everyDays(1).atHour(CONFIG.triggerHour);
    Logger.log(`✅ Created daily trigger at ${CONFIG.triggerHour}:00`);
  }

  builder.create();
  
  // Store current configuration in PropertiesService for future comparison
  properties.setProperty('triggerFrequency', CONFIG.triggerFrequency);
  properties.setProperty('triggerHour', String(CONFIG.triggerHour));
}

/**
 * Remove time-based trigger if disabled in config.
 */
function removeTriggerIfExists() {
  const triggers = ScriptApp.getProjectTriggers();
  const properties = PropertiesService.getScriptProperties();
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'loopThroughContacts') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("🗑️ Removed existing trigger");
    }
  }
  
  // Clean up stored trigger configuration
  properties.deleteProperty('triggerFrequency');
  properties.deleteProperty('triggerHour');
}
