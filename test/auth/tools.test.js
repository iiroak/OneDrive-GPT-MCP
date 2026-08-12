jest.mock('../../auth/token-storage-instance', () => ({
  getValidAccessToken: jest.fn(),
  getExpiryTime: jest.fn()
}));

const tokenStorage = require('../../auth/token-storage-instance');
const { handleCheckAuthStatus } = require('../../auth/tools');

describe('Authentication tools', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('uses TokenStorage so an expired token can be refreshed', async () => {
    tokenStorage.getValidAccessToken.mockResolvedValue('refreshed-access-token');
    tokenStorage.getExpiryTime.mockReturnValue(Date.now() + 3600000);

    const result = await handleCheckAuthStatus();

    expect(tokenStorage.getValidAccessToken).toHaveBeenCalledTimes(1);
    expect(result.content[0].text).toBe('Authenticated and ready');
  });

  it('reports unauthenticated when TokenStorage cannot provide a token', async () => {
    tokenStorage.getValidAccessToken.mockResolvedValue(null);

    const result = await handleCheckAuthStatus();

    expect(result.content[0].text).toBe('Not authenticated');
  });
});
