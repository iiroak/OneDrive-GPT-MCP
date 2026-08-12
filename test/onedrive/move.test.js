const handleMoveItem = require('../../onedrive/move');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleMoveItem', () => {
  beforeEach(() => {
    callGraphAPI.mockReset();
    ensureAuthenticated.mockReset();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
  });

  test('moves and renames an item with one PATCH without downloading it', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'source-id', name: 'recording.mp3', file: {} })
      .mockResolvedValueOnce({ id: 'destination-id', name: 'Spanish', folder: {} })
      .mockResolvedValueOnce({ id: 'source-id', name: 'lesson-01.mp3', webUrl: 'https://onedrive.example/item' });

    const result = await handleMoveItem({
      itemId: 'source-id',
      destinationPath: '/Audio Clases/Spanish',
      newName: 'lesson-01.mp3'
    });

    expect(callGraphAPI).toHaveBeenNthCalledWith(1, 'dummy_access_token', 'GET', 'me/drive/items/source-id');
    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'dummy_access_token', 'GET', 'me/drive/root:/Audio Clases/Spanish');
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      3,
      'dummy_access_token',
      'PATCH',
      'me/drive/items/source-id',
      { parentReference: { id: 'destination-id' }, name: 'lesson-01.mp3' }
    );
    expect(result.structuredContent).toEqual({
      id: 'source-id', name: 'lesson-01.mp3', moved: true, renamed: true
    });
  });

  test('resolves source and destination paths and preserves the current name', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'source-id', name: 'recording.mp3', file: {} })
      .mockResolvedValueOnce({ id: 'root-id', name: 'root', folder: {} })
      .mockResolvedValueOnce({ id: 'source-id', name: 'recording.mp3' });

    await handleMoveItem({ path: '/recording.mp3', destinationPath: '/' });

    expect(callGraphAPI).toHaveBeenNthCalledWith(1, 'dummy_access_token', 'GET', 'me/drive/root:/recording.mp3');
    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'dummy_access_token', 'GET', 'me/drive/root');
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      3,
      'dummy_access_token',
      'PATCH',
      'me/drive/items/source-id',
      { parentReference: { id: 'root-id' } }
    );
  });

  test('resolves the root destination by its documented alias', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'source-id', name: 'recording.mp3' })
      .mockResolvedValueOnce({ id: 'root-id', name: 'root', folder: {} })
      .mockResolvedValueOnce({ id: 'source-id', name: 'recording.mp3' });

    await handleMoveItem({ itemId: 'source-id', destinationPath: 'root' });

    expect(callGraphAPI).toHaveBeenNthCalledWith(2, 'dummy_access_token', 'GET', 'me/drive/root');
  });

  test('rejects an empty operation before authenticating', async () => {
    const result = await handleMoveItem({ itemId: 'source-id' });

    expect(result.content[0].text).toBe('Either destinationPath or newName is required.');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });

  test('rejects a destination item that is not a folder', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ id: 'source-id', name: 'recording.mp3' })
      .mockResolvedValueOnce({ id: 'file-id', name: 'other.mp3', file: {} });

    const result = await handleMoveItem({ itemId: 'source-id', destinationPath: '/other.mp3' });

    expect(result.content[0].text).toBe('Destination is not a folder: /other.mp3');
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
  });
});
