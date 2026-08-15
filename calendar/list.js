/**
 * List events functionality
 */
const config = require('../config');
const { callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { calendarViewPath } = require('./paths');

function formatDateTime(dateTimeData) {
  if (!dateTimeData) return '';
  const dateTime = typeof dateTimeData === 'string' ? dateTimeData : dateTimeData.dateTime || '';
  const timeZone = typeof dateTimeData === 'object' ? dateTimeData.timeZone : undefined;
  if (!dateTime) return '';
  const hasOffset = /[zZ]$|[+\-]\d{2}:\d{2}$/.test(dateTime);
  if (timeZone && timeZone !== 'UTC' && !hasOffset) return `${dateTime} (${timeZone})`;
  const date = new Date(dateTime.endsWith('Z') || hasOffset ? dateTime : `${dateTime}Z`);
  if (Number.isNaN(date.getTime())) return dateTime;
  return date.toLocaleString('en-US', {
    year: 'numeric', month: 'long', day: 'numeric', hour: 'numeric', minute: '2-digit', hour12: true,
    timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
  });
}

async function handleListEvents(args = {}) {
  const count = Math.min(Math.max(Number(args.count) || 10, 1), config.MAX_RESULT_COUNT);
  const startDate = args.startDateTime ? new Date(args.startDateTime) : new Date();
  const endDate = args.endDateTime ? new Date(args.endDateTime) : new Date(startDate.getTime() + 30 * 24 * 60 * 60 * 1000);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime()) || endDate <= startDate) {
    return { content: [{ type: 'text', text: 'Invalid date range. Provide valid startDateTime/endDateTime with endDateTime > startDateTime.' }] };
  }
  try {
    const accessToken = await ensureAuthenticated();
    const response = await callGraphAPIPaginated(accessToken, 'GET', calendarViewPath(args.calendarId), {
      startDateTime: startDate.toISOString(),
      endDateTime: endDate.toISOString(),
      $top: count,
      $orderby: 'start/dateTime',
      $select: config.CALENDAR_SELECT_FIELDS
    }, count);
    const events = response.value || [];
    const eventList = events.map((event, index) => {
      const location = event.location?.displayName || 'No location';
      const categories = event.categories?.length ? event.categories.join(', ') : 'None';
      return `${index + 1}. ${event.subject} - Location: ${location}\nStart: ${formatDateTime(event.start)}\nEnd: ${formatDateTime(event.end)}\nCategories: ${categories}\nSummary: ${event.bodyPreview || ''}\nID: ${event.id}\n`;
    }).join('\n');
    return {
      content: [{ type: 'text', text: events.length ? `Found ${events.length} events:\n\n${eventList}` : 'No calendar events found.' }],
      structuredContent: { calendarId: args.calendarId || null, events, truncated: events.length >= count }
    };
  } catch (error) {
    if (error.message === 'Authentication required') return { content: [{ type: 'text', text: "Authentication required. Complete the MCP OAuth flow first." }] };
    return { content: [{ type: 'text', text: `Error listing events: ${error.message}` }] };
  }
}

module.exports = handleListEvents;
module.exports.formatDateTime = formatDateTime;
