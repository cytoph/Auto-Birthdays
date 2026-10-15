/******************************
 * CONFIGURATION
 ******************************/
const CONFIG = {
  calendarId: 'primary',             // primary or your calendar ID
  // Title customization
  useEmoji: true,                    // Add 🎂 emoji to event titles
  showYearOrAge: true,               // Recurrence on: shows (*YYYY), off: shows (age)

  // Recurrence
  useRecurrence: true,               // Create recurring yearly events
  futureYears: 10,                   // Recurring events end this many years in the future
  pastYears: 1,                      // Recurring events start this many years in the past

  // Reminder settings
  reminderMinutesBefore: 1440,       // Popup reminder (in minutes); 1440 = 1 day before

  // Cleanup
  cleanupEvents: false,              // ⚠️ Deletes all matching birthday events between ±100 years

  // Trigger options
  useTrigger: true,                  // Automatically run on a schedule
  triggerFrequency: 'daily',         // 'daily' or 'hourly'
  triggerHour: 4                     // If 'daily', the hour of day to run (0–23)
};

/**
 * Main function to create/update birthday events.
 * Also manages the optional time-based trigger.
 */
function loopThroughContacts() {
  if (CONFIG.useTrigger) {
    ensureTriggerExists();
  } else {
    removeTriggerIfExists();
  }

  const connections = getAllContacts();
  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);

  if (!calendar) {
    Logger.log("⚠️ Calendar not found.");
    return;
  }

  if (CONFIG.cleanupEvents) {
    cleanupOldBirthdayEvents(calendar, connections);
  }

  const currentYear = new Date().getFullYear();
  const startDate = new Date(currentYear - CONFIG.pastYears - 1, 0, 1);
  const endDate = new Date(currentYear + CONFIG.futureYears + 1, 11, 31);
  const allEvents = calendar.getEvents(startDate, endDate);

  for (const person of connections) {
    const birthdayData = person.birthdays?.find(b => b.date);
    if (birthdayData) {
      updateOrCreateBirthDayEvent(person, birthdayData, calendar, allEvents);
    }
  }
}

/**
 * Get all Google Contacts using pagination.
 */
function getAllContacts() {
  const connections = [];
  let nextPageToken;

  do {
    const reqOpts = {
      personFields: 'names,birthdays',
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
 */
function updateOrCreateBirthDayEvent(person, birthdayRaw, calendar, allEvents) {
  const contactName = getContactName(person);
  const contactId = person.resourceName; // e.g. "people/c1234567890"

  // Extract actual date from birthday object
  const birthdayDate = birthdayRaw.date;
  if (!birthdayDate || typeof birthdayDate.day !== 'number' || typeof birthdayDate.month !== 'number') {
    Logger.log(`⚠️ Skipping ${contactName} due to invalid birthdayDate: ${JSON.stringify(birthdayRaw)}`);
    return;
  }

  const currentYear = new Date().getFullYear();
  const startYear = currentYear - CONFIG.pastYears;

  const month = parseInt(birthdayDate.month, 10) - 1;
  const day = parseInt(birthdayDate.day, 10);

  // Build expected title
  let expectedTitle = contactName + ` hat Geburtstag`;

  // Find and update related events
  const relatedEvents = findBirthdayEvents(allEvents, birthdayDate, contactId);
  const deletedSeriesIds = new Set();
  const deletedEventIds = new Set();
  let correctEventExists = false;

  for (const event of relatedEvents) {
    const title = event.getTitle();
    const isTitleOutdated = title !== expectedTitle;
    const isNotAllDay = !event.isAllDayEvent();
    const isRecurrenceMismatch = CONFIG.useRecurrence !== (event.isRecurringEvent && event.isRecurringEvent());

    if (isTitleOutdated || isNotAllDay || isRecurrenceMismatch) {
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
      } catch (e) {
        Logger.log(`❌ Error deleting event: ${title} → ${e}`);
      }
    } else if (CONFIG.useRecurrence) {
      correctEventExists = true;
    }
  }

  if (correctEventExists) {
    return;
  }

  // Create one new event — recurring or one-time
  if (CONFIG.useRecurrence) {
    const recurrence = CalendarApp.newRecurrence()
      .addYearlyRule()
      .until(new Date(currentYear + CONFIG.futureYears, 11, 31));

    const birthdayStartDate = new Date(startYear, month, day);

    let description = `${contactName} hat Geburtstag!`;

    if (birthdayDate.year) {
      description += ` (Jahrgang: ${birthdayDate.year})`;
    }

    const eventSeries = calendar.createAllDayEventSeries(
      expectedTitle,
      birthdayStartDate,
      recurrence,
      { description: description }
    );
    eventSeries.setTag('contactId', contactId);
    if (CONFIG.reminderMinutesBefore > 0) {
      eventSeries.addPopupReminder(CONFIG.reminderMinutesBefore);
    }
    Logger.log(`🎉 Created RECURRING event: ${expectedTitle} [starts ${birthdayStartDate.toDateString()}]`);
  } else {
    for (i = 0; i < CONFIG.pastYears + CONFIG.futureYears + 1; i++) {
      const thisYear = startYear + i;

      if (birthdayDate.year && (thisYear - birthdayDate.year) < 1) {
        continue; // don't add birthdays for years that person wasn't even born
      }

      const thisBirthdayDate = new Date(thisYear, month, day);
      const age = birthdayDate.year ? thisYear - birthdayDate.year : null;
      const description = contactName + ` ` + (age ? `wird ${age} Jahre alt!` : `hat Geburtstag!`);

      const existingEvent = relatedEvents.find(e => e.getStartTime().getFullYear() == thisYear);
      const existingEventId = existingEvent ? existingEvent.getId() : null;

      if (existingEvent && existingEventId && !deletedEventIds.has(existingEventId)) {
        const isDescriptionOutdated = existingEvent.getDescription() !== description;

        if (!isDescriptionOutdated) {
          continue; // now we're sure everything matches
        }

        deletedEventIds.add(existingEventId);
        existingEvent.deleteEvent();
        Logger.log(`🗑️ Deleted outdated event: ${existingEvent.getTitle()}`);
      }

      const event = calendar.createAllDayEvent(
        expectedTitle,
        thisBirthdayDate,
        { description: description }
      );
      event.setTag('contactId', contactId);

      if (CONFIG.reminderMinutesBefore > 0) {
        event.addPopupReminder(CONFIG.reminderMinutesBefore);
      }

      Logger.log(`🎁 Created ONE-TIME event: ${expectedTitle} [${thisBirthdayDate.toDateString()}]`);
    }
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
 * Find existing birthday events for a contact by contactId tag and date.
 */
function findBirthdayEvents(allEvents, birthdayDate, contactId) {
  return allEvents.filter(event => {
    if (!event.isAllDayEvent()) return false;

    const start = event.getStartTime();
    const isBirthdayDateMatch =
      start.getDate() === birthdayDate.day &&
      start.getMonth() === birthdayDate.month - 1;

    if (!isBirthdayDateMatch) return false;

    return event.getTag('contactId') === contactId;
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

  // Collect valid contacts with birthday info
  const contactBirthdays = allContacts
    .map(person => {
      if (!person.names || !person.birthdays) return null;
      const name = getContactName(person);
      const bday = person.birthdays.find(b => b.date);
      if (!bday) return null;

      return {
        name,
        day: bday.date.day,
        month: bday.date.month - 1,
      };
    })
    .filter(Boolean); // remove nulls

  const deletedSeriesIds = new Set();
  const eventsToDelete = [];

  for (const event of allEvents) {
    const title = event.getTitle();
    const start = event.getStartTime();
    const isAllDay = event.isAllDayEvent();

    for (const contact of contactBirthdays) {
      const isNameMatch = title.includes(contact.name);
      const isBirthdayDateMatch =
        start.getDate() === contact.day &&
        start.getMonth() === contact.month;

      const shouldDelete = isAllDay && isNameMatch && isBirthdayDateMatch;

      if (!shouldDelete) continue;

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
          break;
        } catch (e) {
          Logger.log(`⚠️ Failed to resolve recurring series: ${title} — ${e}`);
          Logger.log(`⚠️ Waiting for rate limiter and safely retrying`);
          Utilities.sleep(5000)
          try {
            const series = event.getEventSeries();
            const seriesId = series.getId();
            if (deletedSeriesIds.has(seriesId)) {
              Logger.log(`✅ Second attempt succeeded for: ${title}`);
              break;
            }
            deletedSeriesIds.add(seriesId);
            eventsToDelete.push({ event: series, isSeries: true });
            Logger.log(`✅ Second attempt succeeded for: ${title}`);
            break;
          } catch (e) {
            Logger.log(`⚠️ Second attempt failed — skipping entry ${e}`);
          }
        }
      } else {
        eventsToDelete.push({ event, isSeries: false });
        break;
      }
    }
  }

  // Now delete all marked events
  for (const { event, isSeries } of eventsToDelete) {
    try {
      if (isSeries) {
        event.deleteEventSeries(); // Deletes the entire series
        Logger.log(`🧹 Deleted recurring series → ${event.getTitle()}`);
      } else {
        event.deleteEvent();
        Logger.log(`🧹 Deleted single event → ${event.getTitle()} on ${event.getStartTime().toDateString()}`);
      }
    } catch (e) {
      Logger.log(`❌ Failed to delete event → ${event.getTitle()} | Reason: ${e}`);
    }
  }

  Logger.log(`✅ Total deleted events: ${eventsToDelete.length}`);
}

/**
 * Ensure a correct time-based trigger exists for loopThroughContacts.
 * Removes old one and creates a new one according to CONFIG.
 */
function ensureTriggerExists() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'loopThroughContacts') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("🗑️ Removed outdated trigger for loopThroughContacts");
    }
  }

  const builder = ScriptApp.newTrigger('loopThroughContacts').timeBased();
  if (CONFIG.triggerFrequency === 'hourly') {
    builder.everyHours(1);
    Logger.log("✅ Created hourly trigger");
  } else {
    builder.everyDays(1).atHour(CONFIG.triggerHour);
    Logger.log(`✅ Created daily trigger at ${CONFIG.triggerHour}:00`);
  }

  builder.create();
}

/**
 * Remove time-based trigger if disabled in config.
 */
function removeTriggerIfExists() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'loopThroughContacts') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("🗑️ Removed existing trigger");
    }
  }
}