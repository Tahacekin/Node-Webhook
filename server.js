const express = require('express');
const session = require('express-session');
const { Client } = require('@microsoft/microsoft-graph-client');
const axios = require('axios');
const crypto = require('crypto');
const { Sequelize } = require('sequelize');
const { Subscription } = require('./models');
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

// Database connection
const sequelize = new Sequelize(
  process.env.DB_NAME || 'webhook_renewal',
  process.env.DB_USER || 'webhook_user',
  process.env.DB_PASSWORD || 'your_secure_password',
  {
    host: process.env.DB_HOST || 'localhost',
    port: process.env.DB_PORT || 5432,
    dialect: 'postgres',
    logging: false
  }
);

// Test database connection
sequelize.authenticate()
  .then(() => {
    console.log('Database connection established successfully.');
  })
  .catch(err => {
    console.error('Unable to connect to the database:', err);
  });

// Microsoft Graph API configuration
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const APP_URL = process.env.APP_URL || 'https://node-webhook-mi3nuu5rt-taha-cekins-projects.vercel.app';
const REDIRECT_URI = `${APP_URL}/callback`;
// Dynamic webhook URL based on environment
const WEBHOOK_URL = process.env.WEBHOOK_URL || `${APP_URL}/webhook`;
const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET;

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
  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
    `client_id=${CLIENT_ID}&` +
    `response_type=code&` +
    `redirect_uri=${encodeURIComponent(REDIRECT_URI)}&` +
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
    const tokenResponse = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      code: code,
      grant_type: 'authorization_code',
      redirect_uri: REDIRECT_URI,
      scope: 'https://graph.microsoft.com/Mail.Read'
    }, {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });
    
    const { access_token, refresh_token } = tokenResponse.data;
    
    // Store tokens in session
    req.session.accessToken = access_token;
    req.session.refreshToken = refresh_token;
    
    res.redirect('/');
  } catch (error) {
    console.error('Token exchange error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to exchange authorization code' });
  }
});

// Part 2: Manual Email Fetching

// API endpoint to fetch emails
app.get('/fetch-emails', async (req, res) => {
  if (!req.session.accessToken) {
    return res.status(401).json({ error: 'Not authenticated. Please login first.' });
  }
  
  try {
    const graphClient = getGraphClient(req.session.accessToken);
    
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
  if (!req.session.accessToken) {
    return res.status(401).json({ error: 'Not authenticated. Please login first.' });
  }
  
  try {
    const graphClient = getGraphClient(req.session.accessToken);
    
    // Get user info to store with subscription
    const user = await graphClient.api('/me').get();
    const userId = user.id;
    
    // Create a subscription for new mail notifications
    const subscription = await graphClient
      .api('/subscriptions')
      .post({
        changeType: 'created',
        notificationUrl: WEBHOOK_URL,
        resource: '/me/messages',
        expirationDateTime: new Date(Date.now() + 1 * 60 * 1000).toISOString(), // 1 minute for testing
        clientState: WEBHOOK_SECRET
      });
    
    // Store subscription in database
    const dbSubscription = await Subscription.create({
      subscriptionId: subscription.id,
      expirationDateTime: new Date(subscription.expirationDateTime),
      userId: userId
    });
    
    // Store subscription ID in session for management
    req.session.subscriptionId = subscription.id;
    
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
      
      // Here you could fetch the specific email details
      // For now, we'll just log the notification
      console.log('Notification details:', {
        changeType: notification.changeType,
        resource: resource,
        clientState: notification.clientState
      });
      
    } catch (error) {
      console.error('Error processing notification:', error);
    }
  }
  
  res.status(200).json({ success: true });
});

// Serve the main page
app.get('/', (req, res) => {
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

// Start renewal service
renewalService.start();

// Manual renewal endpoint for testing
app.post('/manual-renewal', async (req, res) => {
  try {
    await renewalService.manualRenewalCheck();
    res.json({ success: true, message: 'Manual renewal check completed' });
  } catch (error) {
    console.error('Manual renewal error:', error);
    res.status(500).json({ error: 'Manual renewal failed', details: error.message });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Visit http://localhost:${PORT} to start`);
});

module.exports = app;
