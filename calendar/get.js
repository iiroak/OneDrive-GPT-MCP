const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { eventPath } = require('./paths');

async function handleGetEvent(args = {}) {
  if (!args.eventId) return { content: [{ type: 'text', text: 'Event ID is required to get an event.' }] };
  try {
    const accessToken = await ensureAuthenticated();
    const event = await callGraphAPI(accessToken, 'GET', eventPath(args.eventId, args.calendarId));
    return {
      content: [{ type: 'text', text: `Event '${event.subject || args.eventId}' retrieved.` }],
      structuredContent: { event }
    };
  } catch (error) {
    if (error.message === 'Authentication required') return { content: [{ type: 'text', text: "Authentication required. Please use the 'authenticate' tool first." }] };
    return { content: [{ type: 'text', text: `Error getting event: ${error.message}` }] };
  }
}

module.exports = handleGetEvent;
