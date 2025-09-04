# 🔄 Microsoft Graph Webhook Renewal System

A comprehensive Node.js application that automatically manages Microsoft Graph webhook subscriptions, ensuring they never expire by implementing a PostgreSQL-based renewal system.

## ✨ Features

- **🔐 Microsoft Graph Integration**: OAuth 2.0 authentication with Microsoft accounts
- **📧 Email Management**: Fetch and display Outlook emails
- **🔔 Webhook Subscriptions**: Create and manage webhook subscriptions for real-time notifications
- **🔄 Automatic Renewal**: Background service that automatically renews expiring subscriptions
- **💾 Database Persistence**: PostgreSQL database to store subscription data
- **⚡ Serverless Ready**: Optimized for Vercel deployment
- **🧪 Testing Mode**: 1-minute subscription duration for quick testing

## 🏗️ Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   Frontend UI   │    │   Node.js API    │    │  PostgreSQL DB  │
│                 │◄──►│                  │◄──►│                 │
│ - Login/Logout  │    │ - OAuth Flow     │    │ - Subscriptions │
│ - Email Display │    │ - Webhook Mgmt   │    │ - User Data     │
│ - Webhook Setup │    │ - Renewal Service│    │ - Expiration    │
└─────────────────┘    └──────────────────┘    └─────────────────┘
                                │
                                ▼
                       ┌──────────────────┐
                       │ Microsoft Graph  │
                       │      API         │
                       │                  │
                       │ - Email Data     │
                       │ - Webhook Events │
                       │ - Subscription   │
                       │   Management     │
                       └──────────────────┘
```

## 🚀 Quick Start

### Prerequisites

- Node.js (>=18.0.0)
- PostgreSQL database
- Microsoft Graph app registration
- Vercel account (for deployment)

### 1. Clone the Repository

```bash
git clone https://github.com/Tahacekin/Node-Webhook.git
cd Node-Webhook
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Environment Setup

Create a `.env` file with your configuration:

```env
# Microsoft Graph API Configuration
CLIENT_ID=your_microsoft_client_id
CLIENT_SECRET=your_microsoft_client_secret
APP_URL=http://localhost:3000
REDIRECT_URI=http://localhost:3000/callback

# Session Configuration
SESSION_SECRET=your_secure_session_secret

# Webhook Configuration
WEBHOOK_URL=http://localhost:3000/webhook
WEBHOOK_SECRET=your_webhook_secret

# Database Configuration
DB_HOST=localhost
DB_PORT=5432
DB_NAME=webhook_renewal
DB_USER=webhook_user
DB_PASSWORD=your_secure_password
```

### 4. Database Setup

```bash
# Create database and user
psql postgres -c "CREATE DATABASE webhook_renewal;"
psql postgres -c "CREATE USER webhook_user WITH PASSWORD 'your_secure_password';"
psql postgres -c "GRANT ALL PRIVILEGES ON DATABASE webhook_renewal TO webhook_user;"

# Run migrations
npx sequelize-cli db:migrate
```

### 5. Start the Application

```bash
# Development mode
npm run dev

# Production mode
npm start
```

Visit `http://localhost:3000` to access the application.

## 🔧 Configuration

### Microsoft Graph App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to Azure Active Directory → App registrations
3. Create a new app registration
4. Add redirect URI: `https://your-domain.com/callback`
5. Note down the Client ID and Client Secret

### Database Configuration

The application supports both local PostgreSQL and cloud databases:

- **Local Development**: Uses individual database variables
- **Production (Vercel)**: Uses `DATABASE_URL` environment variable

### Testing Configuration

For testing purposes, the application is configured with:
- **Webhook Duration**: 1 minute (instead of 3 days)
- **Renewal Check**: Every 30 seconds
- **Renewal Window**: 2 minutes

## 📁 Project Structure

```
Node-Webhook/
├── api/
│   └── index.js              # Serverless API entry point
├── config/
│   ├── config.json           # Sequelize configuration
│   └── database.js           # Database connection settings
├── migrations/
│   └── *.js                  # Database migrations
├── models/
│   ├── index.js              # Sequelize models index
│   └── subscription.js       # Subscription model
├── services/
│   └── renewalService.js     # Background renewal service
├── public/
│   └── index.html            # Frontend UI
├── server.js                 # Main server file
├── vercel.json              # Vercel deployment config
└── README.md                # This file
```

## 🔄 How It Works

### 1. Authentication Flow
- User clicks "Login with Microsoft"
- Redirected to Microsoft OAuth
- Authorization code exchanged for access token
- Session established

### 2. Webhook Creation
- User creates webhook subscription
- Subscription stored in PostgreSQL database
- Microsoft Graph subscription created with 1-minute expiration

### 3. Automatic Renewal
- Background service checks for expiring subscriptions every 30 seconds
- Finds subscriptions expiring within 2 minutes
- Renews them for another 1 minute
- Updates database with new expiration time

### 4. Webhook Notifications
- Microsoft Graph sends notifications to `/webhook` endpoint
- Notifications processed and logged
- Real-time email updates received

## 🚀 Deployment

### Vercel Deployment

1. **Connect to GitHub**: Link your repository to Vercel
2. **Set Environment Variables**:
   ```env
   DATABASE_URL=postgresql://user:pass@host:port/db
   CLIENT_ID=your_client_id
   CLIENT_SECRET=your_client_secret
   APP_URL=https://your-app.vercel.app
   SESSION_SECRET=your_session_secret
   WEBHOOK_SECRET=your_webhook_secret
   ```
3. **Deploy**: Vercel will automatically deploy on push

### Manual Deployment

```bash
# Install Vercel CLI
npm i -g vercel

# Deploy
vercel --prod
```

## 🧪 Testing

### Local Testing

```bash
# Test database connection
node test-supabase.js

# Test complete setup
node setup-production.js

# Manual renewal test
curl -X POST http://localhost:3000/manual-renewal
```

### Production Testing

1. Visit your deployed URL
2. Login with Microsoft account
3. Create webhook subscription
4. Check database for stored subscription
5. Test manual renewal functionality

## 📊 API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Main application UI |
| `/login` | GET | Microsoft OAuth login |
| `/callback` | GET | OAuth callback handler |
| `/fetch-emails` | GET | Fetch user emails |
| `/create-subscription` | POST | Create webhook subscription |
| `/webhook` | POST | Webhook notification endpoint |
| `/logout` | POST | User logout |
| `/health` | GET | Health check |
| `/test-db` | GET | Database connection test |
| `/manual-renewal` | POST | Manual renewal trigger |

## 🔒 Security

- **Environment Variables**: All secrets stored in environment variables
- **Session Management**: Secure session handling with configurable secrets
- **Database Security**: SSL connections for production databases
- **Input Validation**: Proper validation of webhook notifications
- **Error Handling**: Comprehensive error handling and logging

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Commit changes: `git commit -m 'Add feature'`
4. Push to branch: `git push origin feature-name`
5. Submit a pull request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Troubleshooting

### Common Issues

1. **Database Connection Failed**
   - Check database credentials
   - Verify database is running
   - Check network connectivity

2. **OAuth Redirect URI Mismatch**
   - Ensure redirect URI in Azure matches your app URL
   - Check for trailing slashes

3. **Webhook Not Receiving Notifications**
   - Verify webhook URL is accessible
   - Check Microsoft Graph subscription status
   - Review webhook endpoint logs

4. **Renewal Service Not Working**
   - Check database connection
   - Verify subscription data exists
   - Review renewal service logs

### Getting Help

- Check the [Issues](https://github.com/Tahacekin/Node-Webhook/issues) page
- Review the troubleshooting section above
- Create a new issue with detailed error information

## 🎯 Roadmap

- [ ] Add support for multiple webhook types
- [ ] Implement webhook retry logic
- [ ] Add subscription analytics dashboard
- [ ] Support for multiple Microsoft tenants
- [ ] Webhook notification filtering
- [ ] Real-time subscription monitoring

## 🙏 Acknowledgments

- Microsoft Graph API for webhook capabilities
- Vercel for serverless deployment platform
- PostgreSQL for reliable data storage
- Node.js community for excellent packages

---

**Made with ❤️ for reliable webhook management**