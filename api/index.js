const express = require('express');
const session = require('express-session');
const { Client } = require('@microsoft/microsoft-graph-client');
const axios = require('axios');
const crypto = require('crypto');
const { Sequelize } = require('sequelize');
require('dotenv').config();

const app = express();

// Middleware
app.use(express.json());
app.use(express.static('public'));
app.use(session({
  secret: process.env.SESSION_SECRET || 'your-session-secret',
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false }
}));

// Database connection - simplified for serverless
let sequelize;
try {
  if (process.env.DATABASE_URL) {
    sequelize = new Sequelize(process.env.DATABASE_URL, {
      dialect: 'postgres',
      logging: false,
      dialectOptions: {
        ssl: {
          require: true,
          rejectUnauthorized: false
        }
      }
    });
  }
} catch (error) {
  console.error('Database connection error:', error);
}

// Microsoft Graph API configuration
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const APP_URL = process.env.APP_URL || 'https://node-webhook-mi3nuu5rt-taha-cekins-projects.vercel.app';
const REDIRECT_URI = `${APP_URL}/callback`;
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

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Database test endpoint
app.get('/test-db', async (req, res) => {
  try {
    if (!sequelize) {
      return res.status(500).json({ error: 'Database not configured' });
    }
    
    await sequelize.authenticate();
    const result = await sequelize.query('SELECT NOW() as current_time');
    res.json({ 
      status: 'Database connected', 
      time: result[0][0].current_time,
      environment: process.env.NODE_ENV 
    });
  } catch (error) {
    res.status(500).json({ 
      error: 'Database connection failed', 
      details: error.message,
      environment: process.env.NODE_ENV 
    });
  }
});

// Login route
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

// Callback route
app.get('/callback', async (req, res) => {
  const { code, error } = req.query;
  
  if (error) {
    return res.status(400).json({ error: 'Authentication failed', details: error });
  }
  
  try {
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
    
    req.session.accessToken = access_token;
    req.session.refreshToken = refresh_token;
    
    res.redirect('/');
  } catch (error) {
    console.error('Token exchange error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to exchange authorization code' });
  }
});

// Fetch emails endpoint
app.get('/fetch-emails', async (req, res) => {
  if (!req.session.accessToken) {
    return res.status(401).json({ error: 'Not authenticated. Please login first.' });
  }
  
  try {
    const graphClient = getGraphClient(req.session.accessToken);
    
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

// Create subscription endpoint
app.post('/create-subscription', async (req, res) => {
  if (!req.session.accessToken) {
    return res.status(401).json({ error: 'Not authenticated. Please login first.' });
  }
  
  try {
    const graphClient = getGraphClient(req.session.accessToken);
    
    // Get user info
    const user = await graphClient.api('/me').get();
    const userId = user.id;
    
    // Create subscription
    const subscription = await graphClient
      .api('/subscriptions')
      .post({
        changeType: 'created',
        notificationUrl: WEBHOOK_URL,
        resource: '/me/messages',
        expirationDateTime: new Date(Date.now() + 1 * 60 * 1000).toISOString(), // 1 minute for testing
        clientState: WEBHOOK_SECRET
      });
    
    // Store in database if available
    if (sequelize) {
      try {
        const { Subscription } = require('../models');
        await Subscription.create({
          subscriptionId: subscription.id,
          expirationDateTime: new Date(subscription.expirationDateTime),
          userId: userId
        });
        console.log('Subscription stored in database');
      } catch (dbError) {
        console.error('Database error:', dbError);
      }
    }
    
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

// Webhook endpoint
app.post('/webhook', async (req, res) => {
  const validationToken = req.query.validationToken;
  
  if (validationToken) {
    console.log('Webhook validation request received');
    return res.status(200).send(validationToken);
  }
  
  const notifications = req.body.value;
  if (!notifications || !Array.isArray(notifications)) {
    return res.status(400).json({ error: 'Invalid notification format' });
  }
  
  console.log(`Received ${notifications.length} notification(s)`);
  
  for (const notification of notifications) {
    try {
      if (notification.clientState !== WEBHOOK_SECRET) {
        console.warn('Invalid client state in notification');
        continue;
      }
      
      const resource = notification.resource;
      console.log('New email notification for resource:', resource);
      
    } catch (error) {
      console.error('Error processing notification:', error);
    }
  }
  
  res.status(200).json({ success: true });
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

// Serve the main page
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/../public/index.html');
});

module.exports = app;
