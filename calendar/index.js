/**
 * Calendar module for Outlook MCP server.
 */
const handleListEvents = require('./list');
const handleGetEvent = require('./get');
const handleCreateEvent = require('./create');
const handleUpdateEvent = require('./update');
const handleAcceptEvent = require('./accept');
const handleDeclineEvent = require('./decline');
const handleCancelEvent = require('./cancel');
const handleDeleteEvent = require('./delete');
const handleCopyEvent = require('./copy');
const handleMigrateEvents = require('./migrate');
const handleListMasterCategories = require('./master-categories');
const { handleCreateMasterCategory, handleUpdateMasterCategory, handleDeleteMasterCategory } = require('./categories');
const { handleListCalendars, handleCreateCalendar, handleUpdateCalendar, handleDeleteCalendar } = require('./calendars');

const id = description => ({ type: 'string', minLength: 1, description });
const date = description => ({ description });
const eventFields = {
  calendarId: id('Optional calendar ID; omit for the default calendar'),
  subject: { type: 'string', minLength: 1 },
  start: date('ISO 8601 string or dateTime/timeZone object'),
  end: date('ISO 8601 string or dateTime/timeZone object'),
  body: { description: 'String or body object with contentType and content' },
  location: { description: 'Location string or location object' },
  attendees: { type: 'array', items: { description: 'Email string or attendee object' } },
  categories: { type: 'array', items: { type: 'string' }, description: 'Outlook category names' },
  isAllDay: { type: 'boolean' },
  isReminderOn: { type: 'boolean' },
  reminderMinutesBeforeStart: { type: 'integer', minimum: 0 },
  responseRequested: { type: 'boolean' },
  allowNewTimeProposals: { type: 'boolean' },
  hideAttendees: { type: 'boolean' },
  isOnlineMeeting: { type: 'boolean' },
  onlineMeetingProvider: { type: 'string' },
  importance: { type: 'string' },
  sensitivity: { type: 'string' },
  showAs: { type: 'string' },
  recurrence: { type: 'object' },
  transactionId: id('Graph idempotency identifier')
};

const eventAction = (description, handler, extra = {}) => ({
  description,
  inputSchema: {
    type: 'object',
    properties: { eventId: id('Event ID'), calendarId: id('Optional calendar ID'), ...extra },
    required: ['eventId'],
    additionalProperties: false
  },
  handler
});

const calendarTools = [
  {
    name: 'list-calendars',
    ...{ description: 'Lists calendars available in the connected Outlook mailbox', inputSchema: { type: 'object', properties: { count: { type: 'integer', minimum: 1, maximum: 100 } }, additionalProperties: false }, handler: handleListCalendars }
  },
  {
    name: 'create-calendar',
    ...{ description: 'Creates a new blank calendar', inputSchema: { type: 'object', properties: { name: id('Calendar name') }, required: ['name'], additionalProperties: false }, handler: handleCreateCalendar }
  },
  {
    name: 'update-calendar',
    ...{ description: 'Updates the name or color of a calendar', inputSchema: { type: 'object', properties: { calendarId: id('Calendar ID'), name: { type: 'string' }, color: { type: 'string' } }, required: ['calendarId'], additionalProperties: false }, handler: handleUpdateCalendar }
  },
  {
    name: 'delete-calendar',
    ...{ description: 'Deletes a non-default Outlook calendar', inputSchema: { type: 'object', properties: { calendarId: id('Calendar ID') }, required: ['calendarId'], additionalProperties: false }, handler: handleDeleteCalendar }
  },
  {
    name: 'list-events',
    ...{ description: 'Lists events in a calendar view', inputSchema: { type: 'object', properties: { calendarId: id('Optional calendar ID'), count: { type: 'integer', minimum: 1, maximum: 50 }, startDateTime: { type: 'string' }, endDateTime: { type: 'string' } }, additionalProperties: false }, handler: handleListEvents }
  },
  {
    name: 'get-event',
    ...eventAction('Gets the complete data for an Outlook event', handleGetEvent)
  },
  {
    name: 'create-event',
    description: 'Creates an event in the default or specified calendar',
    inputSchema: { type: 'object', properties: eventFields, required: ['subject', 'start', 'end'], additionalProperties: false },
    handler: handleCreateEvent
  },
  {
    name: 'update-event',
    description: 'Updates an existing event, including its Outlook categories',
    inputSchema: { type: 'object', properties: { eventId: id('Event ID'), ...eventFields }, required: ['eventId'], additionalProperties: false },
    handler: handleUpdateEvent
  },
  {
    name: 'copy-event',
    description: 'Copies an event to another calendar without deleting the source',
    inputSchema: {
      type: 'object',
      properties: {
        sourceEventId: id('Source event ID'), sourceCalendarId: id('Optional source calendar ID'), targetCalendarId: id('Target calendar ID'),
        categoriesMode: { type: 'string', enum: ['preserve', 'clear'] }, includeAttendees: { type: 'boolean' },
        onlineMeetingMode: { type: 'string', enum: ['reject', 'preserveBodyOnly', 'newOnlineMeeting'] }, transactionId: id('Graph idempotency identifier')
      },
      required: ['sourceEventId', 'targetCalendarId'],
      additionalProperties: false
    },
    handler: handleCopyEvent
  },
  {
    name: 'migrate-events',
    description: 'Copies events to another calendar and deletes originals after verification',
    inputSchema: { type: 'object', properties: { sourceEventIds: { type: 'array', minItems: 1, items: id('Source event ID') }, sourceCalendarId: id('Optional source calendar ID'), targetCalendarId: id('Target calendar ID'), confirm: { type: 'string', enum: ['MIGRATE_EVENTS'] }, categoriesMode: { type: 'string', enum: ['preserve', 'clear'] }, includeAttendees: { type: 'boolean' }, onlineMeetingMode: { type: 'string', enum: ['reject', 'preserveBodyOnly', 'newOnlineMeeting'] }, expectedChangeKeys: { type: 'object' }, transactionIdPrefix: { type: 'string' } }, required: ['sourceEventIds', 'targetCalendarId', 'confirm'], additionalProperties: false },
    handler: handleMigrateEvents
  },
  {
    name: 'list-master-categories',
    ...{ description: 'Lists Outlook master categories and colors', inputSchema: { type: 'object', properties: {}, additionalProperties: false }, handler: handleListMasterCategories }
  },
  {
    name: 'create-master-category',
    ...{ description: 'Creates an Outlook master category', inputSchema: { type: 'object', properties: { displayName: id('Category display name'), color: id('Preset color') }, required: ['displayName', 'color'], additionalProperties: false }, handler: handleCreateMasterCategory }
  },
  {
    name: 'update-master-category',
    ...{ description: 'Changes an Outlook master category color', inputSchema: { type: 'object', properties: { categoryId: id('Category ID'), color: id('Preset color') }, required: ['categoryId', 'color'], additionalProperties: false }, handler: handleUpdateMasterCategory }
  },
  {
    name: 'delete-master-category',
    ...{ description: 'Deletes an Outlook master category', inputSchema: { type: 'object', properties: { categoryId: id('Category ID') }, required: ['categoryId'], additionalProperties: false }, handler: handleDeleteMasterCategory }
  },
  { name: 'accept-event', ...eventAction('Accepts a calendar event and sends the response to the organizer', handleAcceptEvent, { comment: { type: 'string' } }) },
  { name: 'decline-event', ...eventAction('Declines a calendar event', handleDeclineEvent, { comment: { type: 'string' } }) },
  { name: 'cancel-event', ...eventAction('Cancels a calendar event', handleCancelEvent, { comment: { type: 'string' } }) },
  { name: 'delete-event', ...eventAction('Deletes a calendar event', handleDeleteEvent) }
];

module.exports = {
  calendarTools,
  handleListEvents,
  handleGetEvent,
  handleCreateEvent,
  handleUpdateEvent,
  handleAcceptEvent,
  handleDeclineEvent,
  handleCancelEvent,
  handleDeleteEvent,
  handleCopyEvent,
  handleMigrateEvents,
  handleListMasterCategories,
  handleCreateMasterCategory,
  handleUpdateMasterCategory,
  handleDeleteMasterCategory,
  handleListCalendars,
  handleCreateCalendar,
  handleUpdateCalendar,
  handleDeleteCalendar
};
