const handleCopyEvent = require('../../calendar/copy');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleCopyEvent', () => {
  beforeEach(() => {
    callGraphAPI.mockReset();
    ensureAuthenticated.mockReset();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
  });

  test('reads, copies, and verifies an appointment without removing the source', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'target-calendar', canEdit: true })
      .mockResolvedValueOnce({
        id: 'source-event', subject: 'Class', type: 'singleInstance',
        body: { contentType: 'HTML', content: 'Details' },
        start: { dateTime: '2026-08-11T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2026-08-11T11:00:00', timeZone: 'UTC' },
        categories: ['Clases UMayor'], attendees: [{ emailAddress: { address: 'a@example.com' }, type: 'required' }]
      })
      .mockResolvedValueOnce({ id: 'copied-event' })
      .mockResolvedValueOnce({ id: 'copied-event', subject: 'Class' });

    const result = await handleCopyEvent({
      sourceEventId: 'source-event',
      targetCalendarId: 'target-calendar',
      categoriesMode: 'clear'
    });

    expect(callGraphAPI).toHaveBeenNthCalledWith(1, 'dummy_access_token', 'GET', 'me/calendars/target-calendar');
    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'dummy_access_token', 'GET', 'me/events/source-event');
    expect(callGraphAPI).toHaveBeenNthCalledWith(3, 'dummy_access_token', 'POST', 'me/calendars/target-calendar/events', expect.objectContaining({ categories: [], transactionId: expect.any(String) }));
    expect(callGraphAPI).toHaveBeenNthCalledWith(4, 'dummy_access_token', 'GET', 'me/calendars/target-calendar/events/copied-event');
    expect(result.structuredContent.sourceRemoved).toBe(false);
    expect(result.structuredContent.warnings).toContain('Attendees were not copied.');
  });

  test('rejects online meetings by default before creating a target', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'target-calendar', canEdit: true })
      .mockResolvedValueOnce({ id: 'source-event', subject: 'Meeting', isOnlineMeeting: true });

    const result = await handleCopyEvent({ sourceEventId: 'source-event', targetCalendarId: 'target-calendar' });

    expect(result.content[0].text).toContain('Online meetings require');
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
  });

  test('rejects identical source and target calendars', async () => {
    const result = await handleCopyEvent({ sourceEventId: 'source-event', sourceCalendarId: 'same', targetCalendarId: 'same' });

    expect(result.content[0].text).toBe('Source and target calendars must be different.');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });
});
