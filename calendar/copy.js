const crypto = require('crypto');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { calendarEventsPath, eventPath, calendarPath } = require('./paths');

function sourcePayload(source, args) {
  if (source.isOnlineMeeting && (args.onlineMeetingMode || 'reject') === 'reject') {
    throw new Error('Online meetings require onlineMeetingMode: preserveBodyOnly or newOnlineMeeting.');
  }
  if (source.type === 'seriesMaster' && !source.recurrence) throw new Error('The recurring series has no recurrence definition to copy.');
  if (source.type === 'occurrence' || source.type === 'exception') {
    if (source.recurrence) throw new Error('Recurring occurrences must be copied as standalone events without recurrence.');
  }
  const payload = {
    subject: source.subject || '(untitled event)',
    body: source.body || { contentType: 'HTML', content: source.bodyPreview || '' },
    start: source.start,
    end: source.end,
    location: source.location,
    isAllDay: source.isAllDay,
    categories: args.categoriesMode === 'clear' ? [] : (source.categories || []),
    ...(source.type === 'seriesMaster' ? { recurrence: source.recurrence } : {}),
    ...(source.isReminderOn !== undefined ? { isReminderOn: source.isReminderOn } : {}),
    ...(source.reminderMinutesBeforeStart !== undefined ? { reminderMinutesBeforeStart: source.reminderMinutesBeforeStart } : {}),
    ...(source.importance ? { importance: source.importance } : {}),
    ...(source.sensitivity ? { sensitivity: source.sensitivity } : {}),
    ...(source.showAs ? { showAs: source.showAs } : {}),
    transactionId: args.transactionId || crypto.randomUUID()
  };
  if (source.isOnlineMeeting && args.onlineMeetingMode === 'newOnlineMeeting') {
    payload.isOnlineMeeting = true;
    if (source.onlineMeetingProvider) payload.onlineMeetingProvider = source.onlineMeetingProvider;
  }
  if (args.includeAttendees) {
    if (source.isOrganizer === false) throw new Error('Attendees can only be recreated for events you organize.');
    payload.attendees = source.attendees || [];
  }
  return payload;
}

async function handleCopyEvent(args = {}) {
  if (!args.sourceEventId || !args.targetCalendarId) {
    return { content: [{ type: 'text', text: 'sourceEventId and targetCalendarId are required to copy an event.' }] };
  }
  if (args.sourceCalendarId && args.sourceCalendarId === args.targetCalendarId) {
    return { content: [{ type: 'text', text: 'Source and target calendars must be different.' }] };
  }
  try {
    const accessToken = await ensureAuthenticated();
    const target = await callGraphAPI(accessToken, 'GET', calendarPath(args.targetCalendarId));
    if (target.canEdit === false) throw new Error('The target calendar is not editable.');
    const source = await callGraphAPI(accessToken, 'GET', eventPath(args.sourceEventId, args.sourceCalendarId));
    const payload = sourcePayload(source, args);
    const copied = await callGraphAPI(accessToken, 'POST', calendarEventsPath(args.targetCalendarId), payload);
    if (!copied?.id) throw new Error('Graph did not return the copied event ID.');
    const verified = await callGraphAPI(accessToken, 'GET', eventPath(copied.id, args.targetCalendarId));
    const warnings = [];
    if (!args.includeAttendees && source.attendees?.length) warnings.push('Attendees were not copied.');
    if (source.isOnlineMeeting && args.onlineMeetingMode === 'preserveBodyOnly') warnings.push('The original online meeting link was retained only in the body.');
    return {
      content: [{ type: 'text', text: `Event '${source.subject || args.sourceEventId}' was copied to calendar ${args.targetCalendarId}.` }],
      structuredContent: {
        sourceEventId: args.sourceEventId,
        targetCalendarId: args.targetCalendarId,
        copiedEvent: verified,
        sourceRemoved: false,
        warnings
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') return { content: [{ type: 'text', text: "Authentication required. Complete the MCP OAuth flow first." }] };
    return { content: [{ type: 'text', text: `Error copying event: ${error.message}` }] };
  }
}

module.exports = handleCopyEvent;
module.exports.sourcePayload = sourcePayload;
