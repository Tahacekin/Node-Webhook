const express = require('express');
const session = require('express-session');
const { Client } = require('@microsoft/microsoft-graph-client');
const axios = require('axios');
const crypto = require('crypto');
const { sequelize, Subscription } = require('./models');
const renewalService = require('./services/renewalService');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.json());
app.use(express.static('public'));
app.use(session({
  secret: process.env.SESSION_SECRET || 'your-session-secret',
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false } // Set to true in production with HTTPS
}));

// Microsoft Graph API configuration
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI || 'http://localhost:3000/callback';
// Use hardcoded webhook URL to avoid environment variable issues
const WEBHOOK_URL = 'https://natural-sparkle-production.up.railway.app/webhook';
const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET;

// Helper function to get valid access token (with refresh logic)
async function getValidAccessToken(req) {
  console.log("=== getValidAccessToken DEBUG START ===");
  console.log("Checking token validity...");
  console.log("Session accessToken type:", typeof req.session.accessToken);
  console.log("Session accessToken value:", req.session.accessToken);
  console.log("Session refreshToken type:", typeof req.session.refreshToken);
  console.log("Session refreshToken value:", req.session.refreshToken);
  console.log("Session tokenExpiresAt:", req.session.tokenExpiresAt);
  
  // If no access token in session, return null
  if (!req.session.accessToken) {
    console.log("No access token in session, returning null");
    return null;
  }
  
  // Check if token is expired (with 5 minute buffer)
  const now = new Date();
  const expiresAt = new Date(req.session.tokenExpiresAt);
  const bufferTime = 5 * 60 * 1000; // 5 minutes in milliseconds
  
  if (expiresAt && now.getTime() > (expiresAt.getTime() - bufferTime)) {
    console.log("Access token expired. Attempting to refresh...");
    
    // If no refresh token, return null
    if (!req.session.refreshToken) {
      console.log("No refresh token available, returning null");
      return null;
    }
    
    try {
      // Refresh the token
      const refreshResponse = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        refresh_token: req.session.refreshToken,
        grant_type: 'refresh_token'
      });
      
      // Update session with new tokens
      req.session.accessToken = refreshResponse.data.access_token;
      req.session.refreshToken = refreshResponse.data.refresh_token || req.session.refreshToken;
      req.session.tokenExpiresAt = new Date(Date.now() + (refreshResponse.data.expires_in * 1000));
      
      console.log("Token refresh successful!");
      console.log("New token expires at:", req.session.tokenExpiresAt);
      
      // Store tokens in renewal service for webhook operations
      if (renewalService && req.session.userId) {
        renewalService.storeUserTokens(req.session.userId, {
          accessToken: req.session.accessToken,
          refreshToken: req.session.refreshToken,
          expiresAt: req.session.tokenExpiresAt
        });
      }
      
    } catch (error) {
      console.error("Token refresh failed:", error.response?.data || error.message);
      // Clear invalid tokens from session
      req.session.accessToken = null;
      req.session.refreshToken = null;
      req.session.tokenExpiresAt = null;
      return null;
    }
  } else {
    console.log("Access token is still valid");
  }
  
  const tokenString = String(req.session.accessToken);
  console.log("Returning access token string:", tokenString);
  console.log("=== getValidAccessToken DEBUG END ===");
  return tokenString;
}

// Helper function to get Graph client
function getGraphClient(accessToken) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
}

// Part 1: Authentication Routes

// Login route - redirects to Microsoft login
app.get('/login', (req, res) => {
  // Use hardcoded redirect URI to avoid environment variable issues
  const hardcodedRedirectUri = 'https://natural-sparkle-production.up.railway.app/callback';
  
  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
    `client_id=${CLIENT_ID}&` +
    `response_type=code&` +
    `redirect_uri=${encodeURIComponent(hardcodedRedirectUri)}&` +
    `response_mode=query&` +
    `scope=https://graph.microsoft.com/Mail.Read&` +
    `state=12345`;
  
  res.redirect(authUrl);
});

// Callback route - handles the authorization code exchange
app.get('/callback', async (req, res) => {
  const { code, error } = req.query;
  
  if (error) {
    return res.status(400).json({ error: 'Authentication failed', details: error });
  }
  
  try {
    // Exchange authorization code for access token
    const hardcodedRedirectUri = 'https://natural-sparkle-production.up.railway.app/callback';
    const tokenResponse = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      code: code,
      grant_type: 'authorization_code',
      redirect_uri: hardcodedRedirectUri,
      scope: 'https://graph.microsoft.com/Mail.Read'
    }, {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });
    
    const { access_token, refresh_token, expires_in } = tokenResponse.data;
    
    console.log("=== /callback TOKEN STORAGE DEBUG ===");
    console.log("access_token type:", typeof access_token);
    console.log("access_token value:", access_token);
    console.log("access_token length:", access_token ? access_token.length : 'null/undefined');
    console.log("refresh_token type:", typeof refresh_token);
    console.log("refresh_token value:", refresh_token);
    console.log("expires_in:", expires_in);
    
    // Store tokens in session
    req.session.accessToken = access_token;
    req.session.refreshToken = refresh_token;
    req.session.userId = 'user-' + Date.now(); // Simple user ID generation
    req.session.tokenExpiresAt = new Date(Date.now() + (expires_in * 1000));
    
    console.log("Stored in session - accessToken:", req.session.accessToken);
    console.log("Stored in session - refreshToken:", req.session.refreshToken);
    console.log("Stored in session - tokenExpiresAt:", req.session.tokenExpiresAt);
    console.log("=== /callback TOKEN STORAGE DEBUG END ===");
    
    // Store tokens in renewal service for webhook renewal
    renewalService.storeUserTokens(
      req.session.userId,
      access_token,
      refresh_token,
      expires_in
    );
    
    console.log("SESSION DATA BEFORE REDIRECT:", req.session);
    res.redirect('/');
  } catch (error) {
    console.error('Token exchange error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to exchange authorization code' });
  }
});

// Part 2: Manual Email Fetching

// API endpoint to fetch emails
app.get('/fetch-emails', async (req, res) => {
  console.log("=== /fetch-emails DEBUG START ===");
  console.log("Session accessToken before getValidAccessToken:", req.session.accessToken);
  
  const validToken = await getValidAccessToken(req);
  console.log("Token received by controller:", validToken);
  console.log("Token type:", typeof validToken);
  console.log("Token length:", validToken ? validToken.length : 'null/undefined');
  
  if (!validToken) {
    console.log("No valid token, returning 401");
    return res.status(401).json({ error: 'Not authenticated. Please login first.' });
  }
  
  try {
    const graphClient = getGraphClient(validToken);
    
    // Fetch the top 10 most recent emails
    const messages = await graphClient
      .api('/me/messages')
      .select('subject,receivedDateTime,from,isRead')
      .top(10)
      .orderby('receivedDateTime desc')
      .get();
    
    res.json({
      success: true,
      emails: messages.value.map(email => ({
        subject: email.subject,
        receivedDateTime: email.receivedDateTime,
        from: email.from?.emailAddress?.name || 'Unknown',
        isRead: email.isRead
      }))
    });
  } catch (error) {
    console.error('Error fetching emails:', error);
    res.status(500).json({ 
      error: 'Failed to fetch emails',
      details: error.message 
    });
  }
});

// Part 3: Webhook Implementation

// Create subscription endpoint
app.post('/create-subscription', async (req, res) => {
  const validToken = await getValidAccessToken(req);
  
  if (!validToken) {
    return res.status(401).json({ error: 'Not authenticated. Please login first.' });
  }
  
  try {
    const graphClient = getGraphClient(validToken);
    
    // Create a subscription for new mail notifications
    const subscription = await graphClient
      .api('/subscriptions')
      .post({
        changeType: 'created',
        notificationUrl: WEBHOOK_URL,
        resource: '/me/messages',
        expirationDateTime: new Date(Date.now() + 4230 * 60 * 1000).toISOString(), // ~3 days
        clientState: WEBHOOK_SECRET
      });
    
    // Store subscription ID in session for management
    req.session.subscriptionId = subscription.id;
    
    // Store subscription in database for renewal service
    try {
      await Subscription.create({
        subscriptionId: subscription.id,
        expirationDateTime: new Date(subscription.expirationDateTime),
        userId: req.session.userId || 'default-user'
      });
      console.log(`Subscription ${subscription.id} stored in database for user: ${req.session.userId}`);
    } catch (dbError) {
      console.warn('Database not available, subscription not stored for renewal:', dbError.message);
      // Continue even if database storage fails
    }
    
    res.json({
      success: true,
      subscription: {
        id: subscription.id,
        expirationDateTime: subscription.expirationDateTime
      }
    });
  } catch (error) {
    console.error('Error creating subscription:', error);
    res.status(500).json({ 
      error: 'Failed to create subscription',
      details: error.message 
    });
  }
});

// Webhook endpoint for receiving notifications
app.post('/webhook', async (req, res) => {
  const validationToken = req.query.validationToken;
  
  // Handle initial validation request
  if (validationToken) {
    console.log('Webhook validation request received');
    return res.status(200).send(validationToken);
  }
  
  // Handle notification
  const notifications = req.body.value;
  if (!notifications || !Array.isArray(notifications)) {
    return res.status(400).json({ error: 'Invalid notification format' });
  }
  
  console.log(`Received ${notifications.length} notification(s)`);
  
  // Process each notification
  for (const notification of notifications) {
    try {
      // Verify the client state if needed
      if (notification.clientState !== WEBHOOK_SECRET) {
        console.warn('Invalid client state in notification');
        continue;
      }
      
      // Get the resource (email) that triggered the notification
      const resource = notification.resource;
      console.log('New email notification for resource:', resource);
      
      // Extract message ID from the resource URL
      // Resource format: /me/messages/{messageId}
      const messageId = resource.split('/').pop();
      console.log('Extracted message ID:', messageId);
      
      // Fetch the full email details using Microsoft Graph API
      try {
        // Try to find the subscription to get the correct user ID
        let userId = 'default-user'; // Fallback
        try {
          // Look up subscription by resource to find the user
          // This is a simplified approach - in production you'd want better mapping
          const subscription = await Subscription.findOne({
            where: {
              // We'll use a simple approach for now
            },
            order: [['createdAt', 'DESC']] // Get the most recent subscription
          });
          
          if (subscription && subscription.userId) {
            userId = subscription.userId;
            console.log('Found user ID from subscription:', userId);
          }
        } catch (lookupError) {
          console.warn('Could not lookup subscription, using default user:', lookupError.message);
        }
        
        const accessToken = await renewalService.getAccessToken(userId);
        
        if (!accessToken) {
          console.warn('No access token available for user:', userId);
          continue;
        }
        
        // Create Graph client with the access token
        const graphClient = getGraphClient(accessToken);
        
        // Fetch the specific email details
        const email = await graphClient
          .api(`/me/messages/${messageId}`)
          .select('subject,from,body,receivedDateTime,isRead')
          .get();
        
        // Log the email subject
        console.log("New Email Received - Subject:", email.subject);
        console.log("From:", email.from?.emailAddress?.name || 'Unknown');
        console.log("Received:", email.receivedDateTime);
        console.log("Is Read:", email.isRead);
        
        // Log additional details if needed
        console.log('Full email details:', {
          id: email.id,
          subject: email.subject,
          from: email.from?.emailAddress,
          receivedDateTime: email.receivedDateTime,
          isRead: email.isRead,
          bodyPreview: email.body?.content ? email.body.content.substring(0, 100) + '...' : 'No body content'
        });
        
      } catch (fetchError) {
        console.error('Error fetching email details:', fetchError);
        console.error('Error details:', fetchError.response?.data || fetchError.message);
      }
      
    } catch (error) {
      console.error('Error processing notification:', error);
    }
  }
  
  res.status(200).json({ success: true });
});

// Serve the main page
app.get('/', (req, res) => {
  console.log("A user visited the root route. Sending index.html now.");
  res.sendFile(__dirname + '/public/index.html');
});

// Logout route
app.post('/logout', (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      return res.status(500).json({ error: 'Could not log out' });
    }
    res.json({ success: true, message: 'Logged out successfully' });
  });
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Manual renewal check endpoint (for testing)
app.post('/manual-renewal-check', async (req, res) => {
  try {
    await renewalService.manualRenewalCheck();
    res.json({ success: true, message: 'Manual renewal check completed' });
  } catch (error) {
    console.error('Error in manual renewal check:', error);
    res.status(500).json({ error: 'Manual renewal check failed', details: error.message });
  }
});

// Start server
app.listen(PORT, async () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Visit http://localhost:${PORT} to start`);
  
  try {
    // Connect to database
    await sequelize.authenticate();
    console.log('Database connection established successfully.');
    
    // Sync database models
    await sequelize.sync();
    console.log('Database models synchronized.');
    
    // Start renewal service
    renewalService.start();
    console.log('Renewal service started.');
  } catch (error) {
    console.warn('Database connection failed, running without renewal service:', error.message);
    console.log('Webhook functionality will work, but subscription renewal is disabled.');
  }
});

module.exports = app;
