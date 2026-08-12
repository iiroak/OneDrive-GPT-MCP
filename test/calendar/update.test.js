const handleUpdateEvent = require('../../calendar/update');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleUpdateEvent', () => {
  beforeEach(() => {
    callGraphAPI.mockReset();
    ensureAuthenticated.mockReset();
  });

  test('updates event categories with a PATCH request', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ id: 'event-123' });

    const result = await handleUpdateEvent({
      eventId: 'event-123',
      categories: ['Clases UMayor']
    });

    expect(callGraphAPI).toHaveBeenCalledWith(
      'dummy_access_token',
      'PATCH',
      'me/events/event-123',
      { categories: ['Clases UMayor'] }
    );
    expect(result.structuredContent).toEqual({
      eventId: 'event-123',
      updatedFields: ['categories'],
      categories: ['Clases UMayor']
    });
  });

  test('maps supported event fields to Graph payloads', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({
      eventId: 'event-123',
      start: '2026-08-11T10:00:00',
      end: { dateTime: '2026-08-11T11:00:00', timeZone: 'Pacific Standard Time' },
      location: 'Room 4',
      body: 'Updated details',
      attendees: ['student@example.com'],
      isAllDay: false
    });

    expect(callGraphAPI.mock.calls[0][3]).toEqual({
      start: { dateTime: '2026-08-11T10:00:00', timeZone: 'Central European Standard Time' },
      end: { dateTime: '2026-08-11T11:00:00', timeZone: 'Pacific Standard Time' },
      location: { displayName: 'Room 4' },
      body: { contentType: 'HTML', content: 'Updated details' },
      attendees: [{ emailAddress: { address: 'student@example.com' }, type: 'required' }],
      isAllDay: false
    });
  });

  test('allows categories to be cleared with an empty array', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({ eventId: 'event-123', categories: [] });

    expect(callGraphAPI.mock.calls[0][3]).toEqual({ categories: [] });
  });

  test('updates recurrence before authenticating with Graph', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({});

    await handleUpdateEvent({
      eventId: 'event-123',
      recurrence: {
        pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday'] },
        range: { type: 'noEnd', startDate: '2026-08-11' }
      }
    });

    expect(callGraphAPI.mock.calls[0][3].recurrence).toEqual({
      pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday'] },
      range: { type: 'noEnd', startDate: '2026-08-11' }
    });
  });

  test('requires an event ID', async () => {
    const result = await handleUpdateEvent({ categories: ['Clases UMayor'] });

    expect(result.content[0].text).toBe('Event ID is required to update an event.');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });
});
