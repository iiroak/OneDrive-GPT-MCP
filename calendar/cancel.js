/**
 * Cancel event functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { eventPath } = require('./paths');

/**
 * Cancel event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCancelEvent(args) {
  const { eventId, comment } = args;

  if (!eventId) {
    return {
      content: [{
        type: "text",
        text: "Event ID is required to cancel an event."
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Build API endpoint
    const endpoint = `${eventPath(eventId, args.calendarId)}/cancel`;

    // Request body
    const body = {
      comment: comment || "Cancelled via API"
    };

    // Make API call
    await callGraphAPI(accessToken, 'POST', endpoint, body);

    return {
      content: [{
        type: "text",
        text: `Event with ID ${eventId} has been successfully cancelled.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Complete the MCP OAuth flow first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error cancelling event: ${error.message}`
      }]
    };
  }
}

module.exports = handleCancelEvent;
