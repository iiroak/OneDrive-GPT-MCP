const {
  handleListCalendars,
  handleCreateCalendar,
  handleUpdateCalendar,
  handleDeleteCalendar
} = require('../../calendar/calendars');
const { callGraphAPI, callGraphAPIPaginated } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('calendar container handlers', () => {
  beforeEach(() => {
    callGraphAPI.mockReset();
    callGraphAPIPaginated.mockReset();
    ensureAuthenticated.mockReset();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
  });

  test('lists calendars through Graph with selected fields', async () => {
    callGraphAPIPaginated.mockResolvedValue({
      value: [{ id: 'calendar/1', name: 'Classes', canEdit: true, isRemovable: true }]
    });

    const result = await handleListCalendars({ count: 10 });

    expect(callGraphAPIPaginated).toHaveBeenCalledWith(
      'dummy_access_token',
      'GET',
      'me/calendars',
      expect.objectContaining({ $top: 10 }),
      10
    );
    expect(result.structuredContent.calendars[0]).toMatchObject({ id: 'calendar/1', name: 'Classes' });
  });

  test('creates a named calendar and returns its ID', async () => {
    callGraphAPI.mockResolvedValue({ id: 'new-calendar', name: 'Clases 2° semestre 2026', canEdit: true });

    const result = await handleCreateCalendar({ name: ' Clases 2° semestre 2026 ' });

    expect(callGraphAPI).toHaveBeenCalledWith(
      'dummy_access_token',
      'POST',
      'me/calendars',
      { name: 'Clases 2° semestre 2026' }
    );
    expect(result.structuredContent.calendar).toMatchObject({ id: 'new-calendar', name: 'Clases 2° semestre 2026' });
  });

  test('rejects an empty calendar name before authentication', async () => {
    ensureAuthenticated.mockReset();

    const result = await handleCreateCalendar({ name: '  ' });

    expect(result.content[0].text).toBe('Calendar name is required.');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('updates and deletes an encoded calendar path', async () => {
    callGraphAPI.mockResolvedValue({ id: 'calendar/1', name: 'Updated' });

    await handleUpdateCalendar({ calendarId: 'calendar/1', name: 'Updated' });
    await handleDeleteCalendar({ calendarId: 'calendar/1' });

    expect(callGraphAPI).toHaveBeenNthCalledWith(1, 'dummy_access_token', 'PATCH', 'me/calendars/calendar%2F1', { name: 'Updated' });
    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'dummy_access_token', 'DELETE', 'me/calendars/calendar%2F1');
  });
});
