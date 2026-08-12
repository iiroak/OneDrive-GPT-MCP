const { handleCreateMasterCategory, handleUpdateMasterCategory, handleDeleteMasterCategory } = require('../../calendar/categories');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('master category mutation handlers', () => {
  beforeEach(() => {
    callGraphAPI.mockReset();
    ensureAuthenticated.mockReset();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
  });

  test('creates, updates, and deletes master categories using encoded IDs', async () => {
    callGraphAPI.mockResolvedValue({ id: 'category/1', displayName: 'Classes', color: 'preset9' });

    await handleCreateMasterCategory({ displayName: 'Classes', color: 'preset9' });
    await handleUpdateMasterCategory({ categoryId: 'category/1', color: 'preset10' });
    await handleDeleteMasterCategory({ categoryId: 'category/1' });

    expect(callGraphAPI).toHaveBeenNthCalledWith(1, 'dummy_access_token', 'POST', 'me/outlook/masterCategories', { displayName: 'Classes', color: 'preset9' });
    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'dummy_access_token', 'PATCH', 'me/outlook/masterCategories/category%2F1', { color: 'preset10' });
    expect(callGraphAPI).toHaveBeenNthCalledWith(3, 'dummy_access_token', 'DELETE', 'me/outlook/masterCategories/category%2F1');
  });
});
