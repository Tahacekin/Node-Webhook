# Webhook Renewal System Setup Instructions

## Overview
This project now includes a comprehensive webhook renewal system with PostgreSQL database integration and automatic subscription renewal.

## Prerequisites
- Node.js (>=18.0.0)
- PostgreSQL
- Homebrew (for macOS)

## Part 1: PostgreSQL Setup

### 1. Install PostgreSQL using Homebrew
```bash
# Install PostgreSQL
brew install postgresql

# Start PostgreSQL service
brew services start postgresql

# Optional: Enable auto-start on boot
brew services enable postgresql
```

### 2. Create Database and User
```bash
# Connect to PostgreSQL as superuser
psql postgres

# Create a new database for your project
CREATE DATABASE webhook_renewal;

# Create a new user for your project
CREATE USER webhook_user WITH PASSWORD 'your_secure_password';

# Grant privileges to the user
GRANT ALL PRIVILEGES ON DATABASE webhook_renewal TO webhook_user;

# Exit psql
\q
```

## Part 2: Environment Configuration

### 1. Create .env file
Copy the `env.example` file to `.env` and update with your actual values:

```bash
cp env.example .env
```

### 2. Update .env with your values
```env
# Microsoft Graph API Configuration
CLIENT_ID=your_actual_client_id
CLIENT_SECRET=your_actual_client_secret
APP_URL=http://localhost:3000
REDIRECT_URI=http://localhost:3000/callback

# Session Configuration
SESSION_SECRET=your_secure_session_secret

# Webhook Configuration
# For production, set this to your Vercel URL + /webhook
# For local development, this will default to APP_URL + /webhook
WEBHOOK_URL=https://your-vercel-app.vercel.app/webhook
WEBHOOK_SECRET=your_webhook_secret

# Database Configuration
DB_HOST=localhost
DB_PORT=5432
DB_NAME=webhook_renewal
DB_USER=webhook_user
DB_PASSWORD=your_secure_password
```

## Part 3: Database Migration

The database tables have already been created. If you need to recreate them:

```bash
# Run migrations
npx sequelize-cli db:migrate

# If you need to reset the database
npx sequelize-cli db:migrate:undo:all
npx sequelize-cli db:migrate
```

## Part 4: Running the Application

### 1. Install Dependencies
```bash
npm install
```

### 2. Start the Application
```bash
# Development mode
npm run dev

# Production mode
npm start
```

### 3. Access the Application
Visit `http://localhost:3000` in your browser.

## Features

### 1. Database Integration
- **Subscription Storage**: All webhook subscriptions are now stored in PostgreSQL
- **User Association**: Subscriptions are linked to specific users
- **Expiration Tracking**: Database tracks subscription expiration dates

### 2. Automatic Renewal System
- **Daily Cron Job**: Runs every day at 2 AM UTC
- **24-Hour Window**: Checks for subscriptions expiring in the next 24 hours
- **Automatic Renewal**: Extends subscriptions by 3 days when they're about to expire
- **Database Updates**: Updates local database with new expiration dates

### 3. Manual Renewal
- **Manual Trigger**: Use the "Manual Renewal" button in the UI
- **Testing**: Useful for testing the renewal system
- **Immediate Check**: Runs the renewal check immediately

### 4. Enhanced UI
- **Logout Button**: Properly clears session and redirects to login
- **Manual Renewal Button**: Allows manual triggering of renewal checks
- **Status Messages**: Clear feedback for all operations

## API Endpoints

### Authentication
- `GET /login` - Redirects to Microsoft login
- `GET /callback` - Handles OAuth callback
- `POST /logout` - Logs out user

### Email Management
- `GET /fetch-emails` - Fetches user's emails
- `POST /create-subscription` - Creates webhook subscription (now stores in DB)

### Webhook Management
- `POST /webhook` - Receives webhook notifications
- `POST /manual-renewal` - Manually triggers renewal check

### System
- `GET /health` - Health check endpoint

## Database Schema

### Subscriptions Table
```sql
CREATE TABLE "Subscriptions" (
  "id" SERIAL PRIMARY KEY,
  "subscriptionId" VARCHAR(255) NOT NULL UNIQUE,
  "expirationDateTime" TIMESTAMP NOT NULL,
  "userId" VARCHAR(255) NOT NULL,
  "createdAt" TIMESTAMP NOT NULL,
  "updatedAt" TIMESTAMP NOT NULL
);
```

## Troubleshooting

### Database Connection Issues
1. Ensure PostgreSQL is running: `brew services list | grep postgresql`
2. Check database credentials in `.env`
3. Verify database exists: `psql -U webhook_user -d webhook_renewal`

### Renewal Service Issues
1. Check logs for renewal service startup messages
2. Verify access tokens are available (currently needs implementation)
3. Test manual renewal using the UI button

### Microsoft Graph API Issues
1. Verify `CLIENT_ID` and `CLIENT_SECRET` are correct
2. Check redirect URI matches your configuration
3. Ensure webhook URL is accessible from Microsoft's servers

## Next Steps

### Token Management
The renewal service currently has a placeholder for access token retrieval. You'll need to implement:

1. **Token Storage**: Store access and refresh tokens for each user
2. **Token Refresh**: Implement automatic token refresh logic
3. **Token Retrieval**: Update `getAccessToken()` method in `renewalService.js`

### Production Deployment

#### Vercel Environment Variables
Set these environment variables in your Vercel project settings:

```env
# Microsoft Graph API Configuration
CLIENT_ID=your_actual_client_id
CLIENT_SECRET=your_actual_client_secret
APP_URL=https://your-vercel-app.vercel.app
REDIRECT_URI=https://your-vercel-app.vercel.app/callback

# Session Configuration
SESSION_SECRET=your_secure_session_secret

# Webhook Configuration
WEBHOOK_URL=https://your-vercel-app.vercel.app/webhook
WEBHOOK_SECRET=your_webhook_secret

# Database Configuration (use Vercel Postgres or external database)
DATABASE_URL=postgresql://username:password@host:port/database
```

#### Additional Steps
1. Update database configuration for production
2. Set up proper environment variables in Vercel
3. Configure production database (e.g., Vercel Postgres)
4. Update Microsoft Graph app registration with production redirect URI

## Testing

### Test the Renewal System
1. Create a webhook subscription
2. Check the database to see the subscription record
3. Use the "Manual Renewal" button to test renewal logic
4. Verify the expiration date is updated in the database

### Test the UI
1. Login with Microsoft account
2. Create a webhook subscription
3. Test the logout functionality
4. Test the manual renewal feature
