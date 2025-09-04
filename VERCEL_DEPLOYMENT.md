# Vercel Deployment Guide

## Overview
This guide will help you deploy your webhook renewal system to Vercel with proper environment configuration.

## Prerequisites
- Vercel account
- GitHub repository with your code
- Microsoft Graph app registration

## Step 1: Deploy to Vercel

### Option A: Deploy from GitHub
1. Connect your GitHub repository to Vercel
2. Import the project
3. Vercel will automatically detect it's a Node.js project

### Option B: Deploy via Vercel CLI
```bash
# Install Vercel CLI
npm i -g vercel

# Deploy
vercel

# Follow the prompts
```

## Step 2: Configure Environment Variables

In your Vercel project dashboard, go to Settings > Environment Variables and add:

### Required Variables
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

# Database Configuration (use Vercel Postgres)
DATABASE_URL=postgresql://username:password@host:port/database
```

### Optional Variables (if not using Vercel Postgres)
```env
DB_HOST=your_db_host
DB_PORT=5432
DB_NAME=your_db_name
DB_USER=your_db_user
DB_PASSWORD=your_db_password
```

## Step 3: Set Up Production Database

### Option A: Vercel Postgres (Recommended)
1. In Vercel dashboard, go to Storage tab
2. Create a new Postgres database
3. Copy the connection string to `DATABASE_URL` environment variable
4. Run migrations: `vercel env pull .env.local && npx sequelize-cli db:migrate`

### Option B: External Database
1. Set up PostgreSQL database (e.g., Railway, Supabase, etc.)
2. Add database credentials to environment variables
3. Run migrations with production database

## Step 4: Update Microsoft Graph App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to Azure Active Directory > App registrations
3. Select your app
4. Go to Authentication
5. Add redirect URI: `https://your-vercel-app.vercel.app/callback`
6. Save changes

## Step 5: Test Deployment

1. Visit your Vercel URL: `https://your-vercel-app.vercel.app`
2. Test the login flow
3. Create a webhook subscription
4. Test the manual renewal feature

## Step 6: Configure Webhook URL in Microsoft Graph

1. In your Microsoft Graph app registration
2. Go to API permissions
3. Ensure you have the necessary permissions for webhooks
4. The webhook URL should be: `https://your-vercel-app.vercel.app/webhook`

## Environment-Specific Behavior

### Local Development
- `APP_URL=http://localhost:3000`
- `REDIRECT_URI=http://localhost:3000/callback`
- `WEBHOOK_URL=http://localhost:3000/webhook` (or custom if set)

### Production (Vercel)
- `APP_URL=https://your-vercel-app.vercel.app`
- `REDIRECT_URI=https://your-vercel-app.vercel.app/callback`
- `WEBHOOK_URL=https://your-vercel-app.vercel.app/webhook`

## Troubleshooting

### Common Issues

1. **OAuth Redirect URI Mismatch**
   - Ensure the redirect URI in Azure matches your Vercel URL
   - Check that `APP_URL` environment variable is set correctly

2. **Database Connection Issues**
   - Verify `DATABASE_URL` is set correctly
   - Ensure database is accessible from Vercel
   - Run migrations: `npx sequelize-cli db:migrate`

3. **Webhook Not Receiving Notifications**
   - Check that `WEBHOOK_URL` is accessible
   - Verify webhook endpoint is working: `https://your-vercel-app.vercel.app/webhook`
   - Check Vercel function logs for errors

4. **Session Issues**
   - Ensure `SESSION_SECRET` is set
   - For production, consider using external session storage (Redis)

### Debugging

1. **Check Vercel Function Logs**
   - Go to Vercel dashboard > Functions tab
   - View logs for your deployed functions

2. **Test Endpoints**
   - Health check: `https://your-vercel-app.vercel.app/health`
   - Manual renewal: `POST https://your-vercel-app.vercel.app/manual-renewal`

3. **Environment Variables**
   - Verify all required variables are set
   - Check for typos in variable names

## Security Considerations

1. **Environment Variables**
   - Never commit `.env` files to version control
   - Use strong, unique secrets for production
   - Rotate secrets regularly

2. **Database Security**
   - Use strong database passwords
   - Enable SSL connections
   - Restrict database access to Vercel IPs if possible

3. **Session Security**
   - Use secure session secrets
   - Consider using external session storage for production
   - Set appropriate session timeouts

## Monitoring

1. **Vercel Analytics**
   - Monitor function performance
   - Track error rates
   - Set up alerts for failures

2. **Application Logs**
   - Monitor renewal service logs
   - Track webhook delivery success
   - Monitor database connection health

## Scaling Considerations

1. **Database Connections**
   - Consider connection pooling for high traffic
   - Monitor database performance

2. **Function Timeouts**
   - Vercel functions have execution time limits
   - Consider breaking long-running tasks into smaller functions

3. **Rate Limiting**
   - Implement rate limiting for API endpoints
   - Monitor Microsoft Graph API rate limits
