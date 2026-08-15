/**
 * Update event functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { eventPath } = require('./paths');
const { buildEventPayload } = require('./event-payload');

async function handleUpdateEvent(args = {}) {
  if (!args.eventId) return { content: [{ type: 'text', text: 'Event ID is required to update an event.' }] };
  try {
    const body = buildEventPayload(args, { requireTimes: false, requireSubject: false });
    if (!Object.keys(body).length) throw new Error('At least one event field is required to update an event.');
    delete body.transactionId;
    const accessToken = await ensureAuthenticated();
    await callGraphAPI(accessToken, 'PATCH', eventPath(args.eventId, args.calendarId), body);
    return {
      content: [{ type: 'text', text: `Event with ID ${args.eventId} has been successfully updated.` }],
      structuredContent: { eventId: args.eventId, updatedFields: Object.keys(body), ...(body.categories ? { categories: body.categories } : {}) }
    };
  } catch (error) {
    if (error.message === 'Authentication required') return { content: [{ type: 'text', text: "Authentication required. Complete the MCP OAuth flow first." }] };
    return { content: [{ type: 'text', text: `Error updating event: ${error.message}` }] };
  }
}

module.exports = handleUpdateEvent;
