const axios = require('axios');

/**
 * Get a valid access token, refreshing if necessary
 * @param {Object} req - Express request object containing session
 * @returns {Promise<string>} - Valid access token
 */
async function getValidAccessToken(req) {
  console.log('🔍 [TOKEN REFRESH] ===== CHECKING TOKEN VALIDITY =====');
  
  // Check if we have tokens in session
  if (!req.session.accessToken || !req.session.refreshToken) {
    console.log('🔍 [TOKEN REFRESH] ❌ No tokens in session');
    throw new Error('No authentication tokens found. Please log in again.');
  }

  const now = new Date();
  const expiresAt = req.session.tokenExpiresAt ? new Date(req.session.tokenExpiresAt) : null;
  
  console.log('🔍 [TOKEN REFRESH] Current time:', now.toISOString());
  console.log('🔍 [TOKEN REFRESH] Token expires at:', expiresAt ? expiresAt.toISOString() : 'Not set');
  
  // Check if token is expired (with 5 minute buffer)
  const bufferTime = 5 * 60 * 1000; // 5 minutes in milliseconds
  const isExpired = !expiresAt || (now.getTime() + bufferTime) >= expiresAt.getTime();
  
  if (!isExpired) {
    console.log('🔍 [TOKEN REFRESH] ✅ Token is still valid');
    console.log('🔍 [TOKEN REFRESH] Returning token preview:', req.session.accessToken ? `${req.session.accessToken.substring(0, 20)}...${req.session.accessToken.substring(req.session.accessToken.length - 20)}` : 'NULL');
    return req.session.accessToken;
  }
  
  console.log('🔍 [TOKEN REFRESH] ⏰ Token is expired, refreshing...');
  
  try {
    // Refresh the token
    const refreshResponse = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      refresh_token: req.session.refreshToken,
      grant_type: 'refresh_token',
      scope: 'https://graph.microsoft.com/Mail.ReadWrite offline_access'
    }, {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });
    
    console.log('🔍 [TOKEN REFRESH] ✅ Token refresh successful');
    console.log('🔍 [TOKEN REFRESH] New access token preview:', refreshResponse.data.access_token ? `${refreshResponse.data.access_token.substring(0, 20)}...${refreshResponse.data.access_token.substring(refreshResponse.data.access_token.length - 20)}` : 'NULL');
    console.log('🔍 [TOKEN REFRESH] Refresh response data:', JSON.stringify(refreshResponse.data, null, 2));
    
    // Update session with new tokens
    req.session.accessToken = refreshResponse.data.access_token;
    req.session.refreshToken = refreshResponse.data.refresh_token;
    
    // Calculate new expiration time (typically 1 hour from now)
    const newExpiresAt = new Date(Date.now() + (refreshResponse.data.expires_in * 1000));
    req.session.tokenExpiresAt = newExpiresAt.toISOString();
    
    console.log('🔍 [TOKEN REFRESH] New token expires at:', newExpiresAt.toISOString());
    console.log('🔍 [TOKEN REFRESH] Returning refreshed token preview:', refreshResponse.data.access_token ? `${refreshResponse.data.access_token.substring(0, 20)}...${refreshResponse.data.access_token.substring(refreshResponse.data.access_token.length - 20)}` : 'NULL');
    console.log('🔍 [TOKEN REFRESH] ======================================');
    
    return refreshResponse.data.access_token;
    
  } catch (error) {
    console.error('🔍 [TOKEN REFRESH] ❌ Token refresh failed:');
    console.error('🔍 [TOKEN REFRESH] Error status:', error.response?.status);
    console.error('🔍 [TOKEN REFRESH] Error data:', error.response?.data);
    
    // Clear invalid tokens from session
    delete req.session.accessToken;
    delete req.session.refreshToken;
    delete req.session.tokenExpiresAt;
    
    throw new Error('Token refresh failed. Please log in again.');
  }
}

/**
 * Test if a token is valid by making a simple API call
 * @param {string} accessToken - Access token to test
 * @returns {Promise<boolean>} - True if token is valid
 */
async function testTokenValidity(accessToken) {
  try {
    const response = await axios.get('https://graph.microsoft.com/v1.0/me', {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });
    return response.status === 200;
  } catch (error) {
    console.log('🔍 [TOKEN REFRESH] Token validation failed:', error.response?.status);
    return false;
  }
}

module.exports = {
  getValidAccessToken,
  testTokenValidity
};
