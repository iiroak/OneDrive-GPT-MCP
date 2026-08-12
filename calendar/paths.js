function segment(value, name = 'ID') {
  if (typeof value !== 'string' || !value) throw new Error(`${name} is required`);
  return encodeURIComponent(value);
}

function calendarPath(calendarId) {
  return calendarId ? `me/calendars/${segment(calendarId, 'calendarId')}` : 'me/calendar';
}

function calendarEventsPath(calendarId) {
  return calendarId ? `${calendarPath(calendarId)}/events` : 'me/events';
}

function calendarViewPath(calendarId) {
  return calendarId ? `${calendarPath(calendarId)}/calendarView` : 'me/calendarView';
}

function eventPath(eventId, calendarId) {
  return `${calendarEventsPath(calendarId)}/${segment(eventId, 'eventId')}`;
}

function masterCategoryPath(categoryId) {
  return `me/outlook/masterCategories/${segment(categoryId, 'categoryId')}`;
}

module.exports = {
  calendarPath,
  calendarEventsPath,
  calendarViewPath,
  eventPath,
  masterCategoryPath,
  segment
};
