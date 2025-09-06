const cron = require('node-cron');
const { Op } = require('sequelize');
const { Subscription } = require('../models');
const axios = require('axios');
require('dotenv').config();

// Simple in-memory token storage (in production, use Redis or database)
const tokenStore = new Map();

class RenewalService {
  constructor() {
    this.isRunning = false;
  }

  // Start the renewal service
  start() {
    if (this.isRunning) {
      console.log('Renewal service is already running');
      return;
    }

    // Run once per day at 2 AM UTC
    this.cronJob = cron.schedule('0 2 * * *', async () => {
      console.log('Starting subscription renewal check...');
      await this.checkAndRenewSubscriptions();
    }, {
      scheduled: false,
      timezone: 'UTC'
    });

    this.cronJob.start();
    this.isRunning = true;
    console.log('Renewal service started - will run daily at 2 AM UTC');
  }

  // Stop the renewal service
  stop() {
    if (this.cronJob) {
      this.cronJob.stop();
      this.isRunning = false;
      console.log('Renewal service stopped');
    }
  }

  // Check for subscriptions that expire in the next 24 hours and renew them
  async checkAndRenewSubscriptions() {
    try {
      const in24Hours = new Date();
      in24Hours.setHours(in24Hours.getHours() + 24); // 24 hours from now

      const now = new Date();

      // Find subscriptions expiring in the next 24 hours
      const expiringSubscriptions = await Subscription.findAll({
        where: {
          expirationDateTime: {
            [Op.between]: [now, in24Hours]
          }
        }
      });

      console.log(`Found ${expiringSubscriptions.length} subscriptions expiring in the next 24 hours`);

      for (const subscription of expiringSubscriptions) {
        try {
          await this.renewSubscription(subscription);
        } catch (error) {
          console.error(`Failed to renew subscription ${subscription.subscriptionId}:`, error.message);
        }
      }

      console.log('Subscription renewal check completed');
    } catch (error) {
      console.error('Error in subscription renewal check:', error);
    }
  }

  // Renew a single subscription
  async renewSubscription(subscription) {
    try {
      console.log(`Renewing subscription ${subscription.subscriptionId} for user ${subscription.userId}`);

      // Calculate new expiration date (3 days from now)
      const newExpirationDate = new Date();
      newExpirationDate.setDate(newExpirationDate.getDate() + 3);

      // Make PATCH request to Microsoft Graph API
      const response = await axios.patch(
        `https://graph.microsoft.com/v1.0/subscriptions/${subscription.subscriptionId}`,
        {
          expirationDateTime: newExpirationDate.toISOString()
        },
        {
          headers: {
            'Authorization': `Bearer ${await this.getAccessToken(subscription.userId)}`,
            'Content-Type': 'application/json'
          }
        }
      );

      // Update the subscription in our database
      await subscription.update({
        expirationDateTime: newExpirationDate
      });

      console.log(`Successfully renewed subscription ${subscription.subscriptionId}. New expiration: ${newExpirationDate.toISOString()}`);
    } catch (error) {
      console.error(`Error renewing subscription ${subscription.subscriptionId}:`, error.response?.data || error.message);
      throw error;
    }
  }

  // Get access token for a user
  async getAccessToken(userId) {
    try {
      const userTokens = tokenStore.get(userId);
      
      if (!userTokens) {
        console.warn(`No tokens found for user ${userId}`);
        return null;
      }

      // Check if token is still valid (with 5 minute buffer)
      const now = new Date();
      const tokenExpiry = new Date(userTokens.expiresAt);
      
      if (now < tokenExpiry) {
        return userTokens.accessToken;
      }

      // Token is expired, try to refresh it
      if (userTokens.refreshToken) {
        console.log(`Refreshing token for user ${userId}`);
        const newTokens = await this.refreshAccessToken(userTokens.refreshToken);
        
        if (newTokens) {
          // Store new tokens
          tokenStore.set(userId, {
            accessToken: newTokens.access_token,
            refreshToken: newTokens.refresh_token || userTokens.refreshToken,
            expiresAt: new Date(Date.now() + (newTokens.expires_in * 1000))
          });
          
          return newTokens.access_token;
        }
      }

      console.error(`Unable to get valid token for user ${userId}`);
      return null;
    } catch (error) {
      console.error(`Error getting access token for user ${userId}:`, error);
      return null;
    }
  }

  // Refresh access token using refresh token
  async refreshAccessToken(refreshToken) {
    try {
      const response = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: 'https://graph.microsoft.com/Mail.Read'
      }, {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        }
      });

      return response.data;
    } catch (error) {
      console.error('Error refreshing access token:', error.response?.data || error.message);
      return null;
    }
  }

  // Store tokens for a user (called when user logs in)
  storeUserTokens(userId, accessToken, refreshToken, expiresIn) {
    tokenStore.set(userId, {
      accessToken,
      refreshToken,
      expiresAt: new Date(Date.now() + (expiresIn * 1000))
    });
  }

  // Manual renewal check (for testing)
  async manualRenewalCheck() {
    console.log('Running manual subscription renewal check...');
    await this.checkAndRenewSubscriptions();
  }
}

module.exports = new RenewalService();
