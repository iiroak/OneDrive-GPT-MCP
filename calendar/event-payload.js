const { DEFAULT_TIMEZONE } = require('../config');

const EVENT_FIELDS = new Set([
  'calendarId', 'subject', 'start', 'end', 'body', 'location', 'attendees', 'categories',
  'isAllDay', 'isReminderOn', 'responseRequested', 'allowNewTimeProposals', 'hideAttendees',
  'isOnlineMeeting', 'importance', 'sensitivity', 'showAs', 'onlineMeetingProvider',
  'reminderMinutesBeforeStart', 'recurrence', 'transactionId', 'eventId'
]);

function normalizeDateTime(value) {
  if (typeof value === 'string' && value) {
    return { dateTime: value, timeZone: DEFAULT_TIMEZONE };
  }
  if (value && typeof value === 'object' && typeof value.dateTime === 'string' && value.dateTime) {
    return { dateTime: value.dateTime, timeZone: value.timeZone || DEFAULT_TIMEZONE };
  }
  throw new Error('start and end must contain a non-empty dateTime');
}

function normalizeAttendees(value) {
  if (!Array.isArray(value)) throw new Error('attendees must be an array');
  return value.map(attendee => {
    if (typeof attendee === 'string' && attendee) {
      return { emailAddress: { address: attendee }, type: 'required' };
    }
    if (!attendee || typeof attendee !== 'object' || !attendee.emailAddress?.address) {
      throw new Error('attendees must contain email strings or emailAddress objects');
    }
    return {
      emailAddress: {
        address: attendee.emailAddress.address,
        ...(attendee.emailAddress.name ? { name: attendee.emailAddress.name } : {})
      },
      type: attendee.type || 'required'
    };
  });
}

function normalizeBody(value) {
  if (typeof value === 'string') return { contentType: 'HTML', content: value };
  if (value && typeof value === 'object' && typeof value.content === 'string') {
    return { contentType: value.contentType || 'HTML', content: value.content };
  }
  throw new Error('body must be a string or content object');
}

function normalizeLocation(value) {
  if (typeof value === 'string') return { displayName: value };
  if (value && typeof value === 'object' && typeof value.displayName === 'string') return value;
  throw new Error('location must be a string or location object');
}

function buildEventPayload(args, { requireTimes = true, includeAttendees = true, requireSubject = true } = {}) {
  const unknownFields = Object.keys(args).filter(field => !EVENT_FIELDS.has(field));
  if (unknownFields.length) throw new Error(`Unsupported event fields: ${unknownFields.join(', ')}`);
  const payload = {};
  if (args.subject !== undefined) payload.subject = args.subject;
  if (args.start !== undefined) payload.start = normalizeDateTime(args.start);
  if (args.end !== undefined) payload.end = normalizeDateTime(args.end);
  if (requireTimes && (!payload.start || !payload.end)) throw new Error('Subject, start, and end times are required to create an event.');
  if (args.body !== undefined) payload.body = normalizeBody(args.body);
  if (args.location !== undefined) payload.location = normalizeLocation(args.location);
  if (args.attendees !== undefined && includeAttendees) payload.attendees = normalizeAttendees(args.attendees);
  if (args.categories !== undefined) {
    if (!Array.isArray(args.categories) || args.categories.some(category => typeof category !== 'string' || !category)) {
      throw new Error('categories must be an array of non-empty strings');
    }
    payload.categories = args.categories;
  }
  for (const field of [
    'isAllDay',
    'isReminderOn',
    'responseRequested',
    'allowNewTimeProposals',
    'hideAttendees',
    'isOnlineMeeting'
  ]) {
    if (args[field] !== undefined) {
      if (typeof args[field] !== 'boolean') throw new Error(`${field} must be a boolean`);
      payload[field] = args[field];
    }
  }
  for (const field of ['importance', 'sensitivity', 'showAs', 'onlineMeetingProvider']) {
    if (args[field] !== undefined) {
      if (typeof args[field] !== 'string' || !args[field]) throw new Error(`${field} must be a non-empty string`);
      payload[field] = args[field];
    }
  }
  if (args.reminderMinutesBeforeStart !== undefined) {
    if (!Number.isInteger(args.reminderMinutesBeforeStart) || args.reminderMinutesBeforeStart < 0) {
      throw new Error('reminderMinutesBeforeStart must be a non-negative integer');
    }
    payload.reminderMinutesBeforeStart = args.reminderMinutesBeforeStart;
  }
  if (args.recurrence !== undefined) {
    if (!args.recurrence || typeof args.recurrence !== 'object') throw new Error('recurrence must be an object');
    payload.recurrence = args.recurrence;
  }
  if (args.transactionId !== undefined) {
    if (typeof args.transactionId !== 'string' || !args.transactionId) throw new Error('transactionId must be a non-empty string');
    payload.transactionId = args.transactionId;
  }
  if (requireSubject && (!payload.subject || typeof payload.subject !== 'string')) {
    throw new Error('Subject, start, and end times are required to create an event.');
  }
  return payload;
}

module.exports = { buildEventPayload, normalizeDateTime };
