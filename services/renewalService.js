const cron = require('node-cron');
const { Op } = require('sequelize');
const { Subscription } = require('../models');
const axios = require('axios');
require('dotenv').config();

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

    // Run every 30 seconds for testing
    this.cronJob = cron.schedule('*/30 * * * * *', async () => {
      console.log('Starting subscription renewal check...');
      await this.checkAndRenewSubscriptions();
    }, {
      scheduled: false,
      timezone: 'UTC'
    });

    this.cronJob.start();
    this.isRunning = true;
    console.log('Renewal service started - will run every 30 seconds for testing');
  }

  // Stop the renewal service
  stop() {
    if (this.cronJob) {
      this.cronJob.stop();
      this.isRunning = false;
      console.log('Renewal service stopped');
    }
  }

  // Check for subscriptions that expire in the next 2 minutes and renew them (for testing)
  async checkAndRenewSubscriptions() {
    try {
      const in2Minutes = new Date();
      in2Minutes.setMinutes(in2Minutes.getMinutes() + 2); // 2 minutes from now

      const now = new Date();

      // Find subscriptions expiring in the next 2 minutes
      const expiringSubscriptions = await Subscription.findAll({
        where: {
          expirationDateTime: {
            [Op.between]: [now, in2Minutes]
          }
        }
      });

      console.log(`Found ${expiringSubscriptions.length} subscriptions expiring in the next 2 minutes`);

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

      // Calculate new expiration date (1 minute from now for testing)
      const newExpirationDate = new Date();
      newExpirationDate.setMinutes(newExpirationDate.getMinutes() + 1);

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

  // Get access token for a user (this is a simplified version)
  // In a real application, you'd need to implement proper token storage and refresh logic
  async getAccessToken(userId) {
    // This is a placeholder - you'll need to implement proper token management
    // For now, we'll return null and handle the error in the renewal process
    console.warn(`Access token retrieval not implemented for user ${userId}. This needs to be implemented with proper token storage.`);
    return null;
  }

  // Manual renewal check (for testing)
  async manualRenewalCheck() {
    console.log('Running manual subscription renewal check...');
    await this.checkAndRenewSubscriptions();
  }
}

module.exports = new RenewalService();
