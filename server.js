const express = require('express');
const session = require('express-session');
const SequelizeStore = require('connect-session-sequelize')(session.Store);
const { Client } = require('@microsoft/microsoft-graph-client');
const axios = require('axios');
const crypto = require('crypto');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const { sequelize, Subscription, User } = require('./models');
const renewalService = require('./services/renewalService');
const imapService = require('./services/imapService');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Configure Sequelize session store
const sessionStore = new SequelizeStore({
  db: sequelize,
  tableName: 'Sessions'
});

// Middleware
app.use(express.json());
app.use(express.static('public'));

// Session middleware with database store for production
app.use(session({
  secret: process.env.SESSION_SECRET || 'your-session-secret',
  store: sessionStore,
  resave: false,
  saveUninitialized: false,
  proxy: true, // Essential for Railway
  cookie: { 
    secure: process.env.NODE_ENV === 'production', // true in production with HTTPS
    httpOnly: true,
    maxAge: 24 * 60 * 60 * 1000 // 24 hours
  }
}));

// Microsoft Graph API configuration
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI || 'http://localhost:3000/callback';
// Use hardcoded webhook URL to avoid environment variable issues
const WEBHOOK_URL = 'https://natural-sparkle-production.up.railway.app/webhook';
const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET;

// JWT configuration
const JWT_SECRET = process.env.JWT_SECRET || 'your-jwt-secret-key';

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

// Email/Password Registration Route
app.post('/auth/register', async (req, res) => {
  console.log("=== /auth/register endpoint hit ===");
  console.log("Request body:", req.body);
  try {
    const { name, email, password } = req.body;
    
    // Validate required fields
    if (!name || !email || !password) {
      return res.status(400).json({ 
        error: 'Missing required fields', 
        details: 'Name, email, and password are required' 
      });
    }
    
    // Check if user already exists
    const existingUser = await User.findOne({ where: { email } });
    if (existingUser) {
      return res.status(409).json({ 
        error: 'User already exists', 
        details: 'A user with this email already exists' 
      });
    }
    
    // Hash the password
    const saltRounds = 12;
    const hashedPassword = await bcrypt.hash(password, saltRounds);
    
    // Create new user
    const user = await User.create({
      name,
      email,
      password: hashedPassword,
      provider: 'local',
      authProvider: 'email'
    });
    
    // Generate JWT token
    const token = jwt.sign(
      { 
        userId: user.id, 
        email: user.email, 
        provider: 'local' 
      },
      JWT_SECRET,
      { expiresIn: '24h' }
    );
    
    // Store user info in session
    req.session.userId = user.id;
    req.session.userEmail = user.email;
    req.session.userName = user.name;
    req.session.provider = 'local';
    req.session.authProvider = user.authProvider;
    
    res.status(201).json({
      success: true,
      message: 'User registered successfully',
      user: {
        id: user.id,
        name: user.name,
        email: user.email,
        provider: user.provider
      },
      token
    });
    
  } catch (error) {
    console.error('Registration error:', error);
    res.status(500).json({ 
      error: 'Registration failed', 
      details: error.message 
    });
  }
});

// Email/Password Login Route
app.post('/auth/login', async (req, res) => {
  console.log("=== /auth/login endpoint hit ===");
  console.log("Request body:", req.body);
  try {
    const { email, password } = req.body;
    
    // Validate required fields
    if (!email || !password) {
      return res.status(400).json({ 
        error: 'Missing required fields', 
        details: 'Email and password are required' 
      });
    }
    
    // Find user by email
    const user = await User.findOne({ where: { email, provider: 'local' } });
    if (!user) {
      return res.status(401).json({ 
        error: 'Invalid credentials', 
        details: 'User not found or invalid login method' 
      });
    }
    
    // Check if user has a password (should always be true for local users)
    if (!user.password) {
      return res.status(401).json({ 
        error: 'Invalid credentials', 
        details: 'User account not properly configured' 
      });
    }
    
    // Compare password
    const isPasswordValid = await bcrypt.compare(password, user.password);
    if (!isPasswordValid) {
      return res.status(401).json({ 
        error: 'Invalid credentials', 
        details: 'Incorrect password' 
      });
    }
    
    // Generate JWT token
    const token = jwt.sign(
      { 
        userId: user.id, 
        email: user.email, 
        provider: 'local' 
      },
      JWT_SECRET,
      { expiresIn: '24h' }
    );
    
    // Store user info in session
    req.session.userId = user.id;
    req.session.userEmail = user.email;
    req.session.userName = user.name;
    req.session.provider = 'local';
    req.session.authProvider = user.authProvider;
    
    res.json({
      success: true,
      message: 'Login successful',
      user: {
        id: user.id,
        name: user.name,
        email: user.email,
        provider: user.provider
      },
      token
    });
    
  } catch (error) {
    console.error('Login error:', error);
    res.status(500).json({ 
      error: 'Login failed', 
      details: error.message 
    });
  }
});

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
    req.session.authProvider = 'microsoft';
    
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
  console.log("Session authProvider:", req.session.authProvider);
  console.log("Session userId:", req.session.userId);
  
  // Check if user is authenticated
  if (!req.session.userId) {
    console.log("No user ID in session, returning 401");
    return res.status(401).json({ error: 'Not authenticated. Please login first.' });
  }
  
  try {
    let emails = [];
    
    if (req.session.authProvider === 'microsoft') {
      // Microsoft OAuth flow
      console.log("Using Microsoft Graph API for email fetching");
      const validToken = await getValidAccessToken(req);
      console.log("Token received by controller:", validToken);
      
      if (!validToken) {
        console.log("No valid Microsoft token, returning 401");
        return res.status(401).json({ error: 'Microsoft authentication expired. Please login again.' });
      }
      
      const graphClient = getGraphClient(validToken);
      
      // Fetch the top 10 most recent emails
      const messages = await graphClient
        .api('/me/messages')
        .select('subject,receivedDateTime,from,isRead')
        .top(10)
        .orderby('receivedDateTime desc')
        .get();
      
      emails = messages.value.map(email => ({
        subject: email.subject,
        receivedDateTime: email.receivedDateTime,
        from: email.from?.emailAddress?.name || 'Unknown',
        isRead: email.isRead
      }));
      
    } else if (req.session.authProvider === 'email') {
      // IMAP flow for standard email providers
      console.log("Using IMAP for email fetching");
      
      // Get user from database to access email and password
      const user = await User.findByPk(req.session.userId);
      if (!user) {
        return res.status(401).json({ error: 'User not found. Please login again.' });
      }
      
      if (!user.email || !user.password) {
        return res.status(400).json({ 
          error: 'Email credentials not available. Please re-register with your email credentials.' 
        });
      }
      
      // For IMAP, we need to temporarily decrypt the password
      // Note: This is a security consideration - in production, consider using OAuth2 for email providers
      const emailPassword = req.body.emailPassword || req.query.emailPassword;
      
      if (!emailPassword) {
        return res.status(400).json({ 
          error: 'Email password required for IMAP access. Please provide your email password.',
          requiresPassword: true
        });
      }
      
      // Test connection first
      try {
        await imapService.testConnection(user.email, emailPassword);
      } catch (connectionError) {
        return res.status(400).json({ 
          error: 'Invalid email credentials. Please check your email and password.',
          details: connectionError.message
        });
      }
      
      // Fetch emails using IMAP
      emails = await imapService.fetchEmails(user.email, emailPassword, 10);
      
    } else {
      return res.status(400).json({ 
        error: 'Unknown authentication provider. Please login again.' 
      });
    }
    
    console.log(`Successfully fetched ${emails.length} emails`);
    res.json({
      success: true,
      emails: emails,
      authProvider: req.session.authProvider
    });
    
  } catch (error) {
    console.error('Error fetching emails:', error);
    res.status(500).json({ 
      error: 'Failed to fetch emails',
      details: error.message 
    });
  }
});

// POST endpoint for fetch-emails (for IMAP password handling)
app.post('/fetch-emails', async (req, res) => {
  console.log("=== /fetch-emails POST DEBUG START ===");
  console.log("Session authProvider:", req.session.authProvider);
  console.log("Session userId:", req.session.userId);
  
  // Check if user is authenticated
  if (!req.session.userId) {
    console.log("No user ID in session, returning 401");
    return res.status(401).json({ error: 'Not authenticated. Please login first.' });
  }
  
  try {
    let emails = [];
    
    if (req.session.authProvider === 'microsoft') {
      // Microsoft OAuth flow
      console.log("Using Microsoft Graph API for email fetching");
      const validToken = await getValidAccessToken(req);
      console.log("Token received by controller:", validToken);
      
      if (!validToken) {
        console.log("No valid Microsoft token, returning 401");
        return res.status(401).json({ error: 'Microsoft authentication expired. Please login again.' });
      }
      
      const graphClient = getGraphClient(validToken);
      
      // Fetch the top 10 most recent emails
      const messages = await graphClient
        .api('/me/messages')
        .select('subject,receivedDateTime,from,isRead')
        .top(10)
        .orderby('receivedDateTime desc')
        .get();
      
      emails = messages.value.map(email => ({
        subject: email.subject,
        receivedDateTime: email.receivedDateTime,
        from: email.from?.emailAddress?.name || 'Unknown',
        isRead: email.isRead
      }));
      
    } else if (req.session.authProvider === 'email') {
      // IMAP flow for standard email providers
      console.log("Using IMAP for email fetching");
      
      // Get user from database to access email and password
      const user = await User.findByPk(req.session.userId);
      if (!user) {
        return res.status(401).json({ error: 'User not found. Please login again.' });
      }
      
      if (!user.email) {
        return res.status(400).json({ 
          error: 'Email not available. Please re-register with your email credentials.' 
        });
      }
      
      // Get email password from request body
      const emailPassword = req.body.emailPassword;
      
      if (!emailPassword) {
        return res.status(400).json({ 
          error: 'Email password required for IMAP access. Please provide your email password.',
          requiresPassword: true
        });
      }
      
      // Test connection first
      try {
        await imapService.testConnection(user.email, emailPassword);
      } catch (connectionError) {
        return res.status(400).json({ 
          error: 'Invalid email credentials. Please check your email and password.',
          details: connectionError.message
        });
      }
      
      // Fetch emails using IMAP
      emails = await imapService.fetchEmails(user.email, emailPassword, 10);
      
    } else {
      return res.status(400).json({ 
        error: 'Unknown authentication provider. Please login again.' 
      });
    }
    
    console.log(`Successfully fetched ${emails.length} emails`);
    res.json({
      success: true,
      emails: emails,
      authProvider: req.session.authProvider
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
    
    // Sync session store to create Sessions table
    await sessionStore.sync();
    console.log('Session store synchronized.');
    
    // Start renewal service
    renewalService.start();
    console.log('Renewal service started.');
  } catch (error) {
    console.warn('Database connection failed, running without renewal service:', error.message);
    console.log('Webhook functionality will work, but subscription renewal is disabled.');
  }
});

module.exports = app;
