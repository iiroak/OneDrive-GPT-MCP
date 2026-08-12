const handleListMasterCategories = require('../../calendar/master-categories');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleListMasterCategories', () => {
  beforeEach(() => {
    callGraphAPI.mockReset();
    ensureAuthenticated.mockReset();
  });

  test('lists Outlook master categories with their colors', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({
      value: [{ id: 'category-1', displayName: 'Clases UMayor', color: 'preset9' }]
    });

    const result = await handleListMasterCategories();

    expect(callGraphAPI).toHaveBeenCalledWith(
      'dummy_access_token',
      'GET',
      'me/outlook/masterCategories'
    );
    expect(result.structuredContent).toEqual({
      categories: [{ id: 'category-1', displayName: 'Clases UMayor', color: 'preset9' }]
    });
    expect(result.content[0].text).toContain('Clases UMayor (preset9)');
  });

  test('handles an empty master category list', async () => {
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    callGraphAPI.mockResolvedValue({ value: [] });

    const result = await handleListMasterCategories();

    expect(result.content[0].text).toBe('No Outlook master categories found.');
    expect(result.structuredContent).toEqual({ categories: [] });
  });
});
