const handleCopyEvent = require('./copy');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { eventPath } = require('./paths');

async function handleMigrateEvents(args = {}) {
  if (args.confirm !== 'MIGRATE_EVENTS') return { content: [{ type: 'text', text: 'Migration requires confirm: MIGRATE_EVENTS.' }] };
  if (!Array.isArray(args.sourceEventIds) || !args.sourceEventIds.length || !args.targetCalendarId) {
    return { content: [{ type: 'text', text: 'sourceEventIds and targetCalendarId are required to migrate events.' }] };
  }
  try {
    const accessToken = await ensureAuthenticated();
    const results = [];
    for (const sourceEventId of args.sourceEventIds) {
      const copy = await handleCopyEvent({ ...args, sourceEventId, transactionId: `${args.transactionIdPrefix || 'migrate'}-${sourceEventId}` });
      const copied = copy.structuredContent?.copiedEvent;
      if (!copied?.id) {
        results.push({ sourceEventId, state: 'copy_failed', message: copy.content?.[0]?.text });
        continue;
      }
      const current = await callGraphAPI(accessToken, 'GET', eventPath(sourceEventId, args.sourceCalendarId));
      if (args.expectedChangeKeys?.[sourceEventId] && args.expectedChangeKeys[sourceEventId] !== current.changeKey) {
        results.push({ sourceEventId, targetEventId: copied.id, state: 'copied_source_retained', message: 'Source changed after copy.' });
        continue;
      }
      await callGraphAPI(accessToken, 'DELETE', eventPath(sourceEventId, args.sourceCalendarId));
      results.push({ sourceEventId, targetEventId: copied.id, state: 'migrated' });
    }
    return {
      content: [{ type: 'text', text: `Migration finished: ${results.filter(result => result.state === 'migrated').length}/${results.length} events migrated.` }],
      structuredContent: { targetCalendarId: args.targetCalendarId, results }
    };
  } catch (error) {
    if (error.message === 'Authentication required') return { content: [{ type: 'text', text: "Authentication required. Please use the 'authenticate' tool first." }] };
    return { content: [{ type: 'text', text: `Error migrating events: ${error.message}` }] };
  }
}

module.exports = handleMigrateEvents;
