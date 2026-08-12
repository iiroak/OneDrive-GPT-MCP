/**
 * Create event functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { calendarEventsPath } = require('./paths');
const { buildEventPayload } = require('./event-payload');

async function handleCreateEvent(args = {}) {
  let body;
  try {
    body = buildEventPayload(args);
  } catch (error) {
    return { content: [{ type: 'text', text: error.message }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const event = await callGraphAPI(accessToken, 'POST', calendarEventsPath(args.calendarId), body);
    return {
      content: [{ type: 'text', text: `Event '${body.subject}' has been successfully created.` }],
      structuredContent: {
        event: {
          id: event.id,
          calendarId: args.calendarId || null,
          subject: event.subject || body.subject,
          start: event.start || body.start,
          end: event.end || body.end,
          webLink: event.webLink
        }
      }
    };
  } catch (error) {
    if (error.message === 'Authentication required') return { content: [{ type: 'text', text: "Authentication required. Please use the 'authenticate' tool first." }] };
    return { content: [{ type: 'text', text: `Error creating event: ${error.message}` }] };
  }
}

module.exports = handleCreateEvent;
