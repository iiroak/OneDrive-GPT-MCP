const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { calendarPath } = require('./paths');

function calendarSummary(calendar) {
  return {
    id: calendar.id,
    name: calendar.name,
    color: calendar.color,
    hexColor: calendar.hexColor,
    isDefaultCalendar: calendar.isDefaultCalendar,
    canEdit: calendar.canEdit,
    canShare: calendar.canShare,
    canViewPrivateItems: calendar.canViewPrivateItems,
    isRemovable: calendar.isRemovable,
    owner: calendar.owner
  };
}

async function handleListCalendars(args = {}) {
  try {
    const accessToken = await ensureAuthenticated();
    const count = Math.min(Math.max(Number(args.count) || 50, 1), 100);
    const response = await callGraphAPIPaginated(accessToken, 'GET', 'me/calendars', {
      $top: count,
      $select: 'id,name,color,hexColor,isDefaultCalendar,canEdit,canShare,canViewPrivateItems,isRemovable,owner'
    }, count);
    const calendars = (response.value || []).map(calendarSummary);
    return {
      content: [{ type: 'text', text: calendars.length
        ? `Found ${calendars.length} calendars:\n\n${calendars.map(calendar => `${calendar.name} (${calendar.id})`).join('\n')}`
        : 'No calendars found.' }],
      structuredContent: { calendars, truncated: calendars.length >= count }
    };
  } catch (error) {
    return errorResponse('listing calendars', error);
  }
}

async function handleCreateCalendar(args = {}) {
  if (typeof args.name !== 'string' || !args.name.trim()) return message('Calendar name is required.');
  try {
    const accessToken = await ensureAuthenticated();
    const calendar = await callGraphAPI(accessToken, 'POST', 'me/calendars', { name: args.name.trim() });
    return {
      content: [{ type: 'text', text: `Calendar '${args.name.trim()}' has been successfully created.` }],
      structuredContent: { calendar: calendarSummary(calendar) }
    };
  } catch (error) {
    return errorResponse('creating calendar', error);
  }
}

async function handleUpdateCalendar(args = {}) {
  if (!args.calendarId) return message('Calendar ID is required to update a calendar.');
  const body = {};
  if (args.name !== undefined) body.name = args.name;
  if (args.color !== undefined) body.color = args.color;
  if (!Object.keys(body).length) return message('At least one calendar field is required to update a calendar.');
  try {
    const accessToken = await ensureAuthenticated();
    const calendar = await callGraphAPI(accessToken, 'PATCH', calendarPath(args.calendarId), body);
    return {
      content: [{ type: 'text', text: `Calendar with ID ${args.calendarId} has been successfully updated.` }],
      structuredContent: { calendar: calendarSummary(calendar) }
    };
  } catch (error) {
    return errorResponse('updating calendar', error);
  }
}

async function handleDeleteCalendar(args = {}) {
  if (!args.calendarId) return message('Calendar ID is required to delete a calendar.');
  try {
    const accessToken = await ensureAuthenticated();
    await callGraphAPI(accessToken, 'DELETE', calendarPath(args.calendarId));
    return {
      content: [{ type: 'text', text: `Calendar with ID ${args.calendarId} has been successfully deleted.` }],
      structuredContent: { calendarId: args.calendarId, deleted: true }
    };
  } catch (error) {
    return errorResponse('deleting calendar', error);
  }
}

function message(text) {
  return { content: [{ type: 'text', text }] };
}

function errorResponse(action, error) {
  if (error.message === 'Authentication required') return message("Authentication required. Please use the 'authenticate' tool first.");
  return message(`Error ${action}: ${error.message}`);
}

module.exports = {
  handleListCalendars,
  handleCreateCalendar,
  handleUpdateCalendar,
  handleDeleteCalendar,
  calendarSummary
};
