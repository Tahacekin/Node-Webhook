const Imap = require('node-imap');
const { simpleParser } = require('mailparser');

class ImapService {
  constructor() {
    // IMAP configuration for common email providers
    this.imapConfigs = {
      'gmail.com': {
        host: 'imap.gmail.com',
        port: 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false }
      },
      'outlook.com': {
        host: 'outlook.office365.com',
        port: 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false }
      },
      'hotmail.com': {
        host: 'outlook.office365.com',
        port: 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false }
      },
      'yahoo.com': {
        host: 'imap.mail.yahoo.com',
        port: 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false }
      },
      'icloud.com': {
        host: 'imap.mail.me.com',
        port: 993,
        tls: true,
        tlsOptions: { rejectUnauthorized: false }
      }
    };
  }

  /**
   * Get IMAP configuration for a given email domain
   * @param {string} email - User's email address
   * @returns {Object} IMAP configuration
   */
  getImapConfig(email) {
    const domain = email.split('@')[1].toLowerCase();
    return this.imapConfigs[domain] || {
      host: process.env.IMAP_HOST || 'imap.gmail.com',
      port: parseInt(process.env.IMAP_PORT) || 993,
      tls: true,
      tlsOptions: { rejectUnauthorized: false }
    };
  }

  /**
   * Fetch emails using IMAP
   * @param {string} email - User's email address
   * @param {string} password - User's email password
   * @param {number} limit - Number of emails to fetch (default: 10)
   * @returns {Promise<Array>} Array of email objects
   */
  async fetchEmails(email, password, limit = 10) {
    return new Promise((resolve, reject) => {
      const config = this.getImapConfig(email);
      const imap = new Imap({
        user: email,
        password: password,
        ...config
      });

      const emails = [];

      imap.once('ready', () => {
        imap.openBox('INBOX', false, (err, box) => {
          if (err) {
            imap.end();
            return reject(new Error(`Failed to open INBOX: ${err.message}`));
          }

          // Search for unseen emails first, then all emails
          imap.search(['UNSEEN'], (err, results) => {
            if (err) {
              imap.end();
              return reject(new Error(`Search failed: ${err.message}`));
            }

            if (results.length === 0) {
              // If no unseen emails, get recent emails
              imap.search(['ALL'], (err, allResults) => {
                if (err) {
                  imap.end();
                  return reject(new Error(`Search failed: ${err.message}`));
                }

                if (allResults.length === 0) {
                  imap.end();
                  return resolve([]);
                }

                // Get the most recent emails
                const recentResults = allResults.slice(-limit);
                this.fetchEmailBodies(imap, recentResults, emails, resolve, reject, limit);
              });
            } else {
              // Get unseen emails
              const unseenResults = results.slice(-limit);
              this.fetchEmailBodies(imap, unseenResults, emails, resolve, reject, limit);
            }
          });
        });
      });

      imap.once('error', (err) => {
        reject(new Error(`IMAP connection failed: ${err.message}`));
      });

      imap.once('end', () => {
        // Connection ended
      });

      imap.connect();
    });
  }

  /**
   * Fetch email bodies and parse them
   * @param {Object} imap - IMAP connection
   * @param {Array} messageIds - Array of message IDs
   * @param {Array} emails - Array to store parsed emails
   * @param {Function} resolve - Promise resolve function
   * @param {Function} reject - Promise reject function
   * @param {number} limit - Maximum number of emails to process
   */
  fetchEmailBodies(imap, messageIds, emails, resolve, reject, limit) {
    if (messageIds.length === 0) {
      imap.end();
      return resolve(emails);
    }

    const fetch = imap.fetch(messageIds, {
      bodies: '',
      struct: true
    });

    fetch.on('message', (msg, seqno) => {
      let buffer = '';

      msg.on('body', (stream, info) => {
        stream.on('data', (chunk) => {
          buffer += chunk.toString('utf8');
        });

        stream.once('end', () => {
          simpleParser(buffer, (err, parsed) => {
            if (err) {
              console.error('Error parsing email:', err);
              return;
            }

            const email = {
              subject: parsed.subject || '(No Subject)',
              from: parsed.from?.text || 'Unknown',
              receivedDateTime: parsed.date || new Date(),
              isRead: false, // IMAP doesn't provide read status easily
              body: parsed.text || parsed.html || '',
              messageId: parsed.messageId
            };

            emails.push(email);

            if (emails.length >= limit) {
              imap.end();
              resolve(emails);
            }
          });
        });
      });

      msg.once('attributes', (attrs) => {
        // You can access email flags here if needed
      });

      msg.once('end', () => {
        // Message processing complete
      });
    });

    fetch.once('error', (err) => {
      imap.end();
      reject(new Error(`Fetch failed: ${err.message}`));
    });

    fetch.once('end', () => {
      imap.end();
      resolve(emails);
    });
  }

  /**
   * Test IMAP connection
   * @param {string} email - User's email address
   * @param {string} password - User's email password
   * @returns {Promise<boolean>} True if connection successful
   */
  async testConnection(email, password) {
    return new Promise((resolve, reject) => {
      const config = this.getImapConfig(email);
      const imap = new Imap({
        user: email,
        password: password,
        ...config
      });

      imap.once('ready', () => {
        imap.end();
        resolve(true);
      });

      imap.once('error', (err) => {
        reject(new Error(`IMAP connection failed: ${err.message}`));
      });

      imap.connect();
    });
  }
}

module.exports = new ImapService();
