# Outlook Email Reader with Webhook Support

A Node.js application that connects to Microsoft Outlook via Graph API to read emails manually and receive real-time notifications through webhooks. Deployable on Vercel.

## Features

- 🔐 OAuth 2.0 authentication with Microsoft Graph API
- 📧 Manual email fetching with a beautiful web interface
- 🔔 Real-time email notifications via webhooks
- 🚀 Ready for Vercel deployment
- 📱 Responsive design

## Prerequisites

- Node.js 18+ installed
- Microsoft Azure account
- Vercel account (for deployment)

## Part 1: Azure Portal Setup

### Step 1: Register a New Application

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Navigate to **Azure Active Directory** (now called "Microsoft Entra ID")
3. Click on **App registrations** in the left sidebar
4. Click **New registration**
5. Fill in the application details:
   - **Name**: `Outlook Email Reader` (or any name you prefer)
   - **Supported account types**: Select "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: 
     - Platform: **Web**
     - URI: `http://localhost:3000/callback` (for local development)
6. Click **Register**

### Step 2: Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions**
5. Search for and select:
   - `Mail.Read` - Read user mail
6. Click **Add permissions**
7. Click **Grant admin consent** (if you have admin rights)

### Step 3: Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description: `Webhook App Secret`
4. Choose expiration (recommend 12 months)
5. Click **Add**
6. **IMPORTANT**: Copy the secret value immediately - you won't be able to see it again!

### Step 4: Note Your Application Details

From the **Overview** page, copy:
- **Application (client) ID**
- **Directory (tenant) ID** (you might need this later)

## Part 2: Local Development Setup

### Step 1: Install Dependencies

```bash
npm install
```

### Step 2: Environment Configuration

1. Copy the example environment file:
   ```bash
   cp env.example .env
   ```

2. Edit `.env` and add your Azure app details:
   ```env
   CLIENT_ID=your_client_id_from_azure
   CLIENT_SECRET=your_client_secret_from_azure
   REDIRECT_URI=http://localhost:3000/callback
   SESSION_SECRET=your_random_session_secret
   WEBHOOK_URL=https://your-vercel-app.vercel.app/webhook
   WEBHOOK_SECRET=your_random_webhook_secret
   ```

### Step 3: Run the Application

```bash
# Development mode with auto-restart
npm run dev

# Or production mode
npm start
```

Visit `http://localhost:3000` to see the application.

## Part 3: Vercel Deployment

### Step 1: Install Vercel CLI

```bash
npm install -g vercel
```

### Step 2: Deploy to Vercel

```bash
# Login to Vercel
vercel login

# Deploy the project
vercel

# Follow the prompts:
# - Set up and deploy? Yes
# - Which scope? (select your account)
# - Link to existing project? No
# - Project name: outlook-webhook-app (or your preferred name)
# - Directory: ./
# - Override settings? No
```

### Step 3: Configure Environment Variables

1. Go to your Vercel dashboard
2. Select your project
3. Go to **Settings** → **Environment Variables**
4. Add all the variables from your `.env` file:
   - `CLIENT_ID`
   - `CLIENT_SECRET`
   - `SESSION_SECRET`
   - `WEBHOOK_URL` (use your Vercel app URL)
   - `WEBHOOK_SECRET`

### Step 4: Update Azure Redirect URI

1. Go back to your Azure app registration
2. Go to **Authentication**
3. Add a new redirect URI:
   - Platform: **Web**
   - URI: `https://your-vercel-app.vercel.app/callback`
4. Save the changes

## Part 4: Webhook Configuration

### Understanding Webhooks

The application includes webhook functionality that allows Microsoft Graph to send notifications when new emails arrive. Here's how it works:

1. **Subscription Creation**: When you click "Setup Webhook", the app creates a subscription with Microsoft Graph
2. **Validation**: Microsoft Graph sends a validation request to your webhook URL
3. **Notifications**: When new emails arrive, Microsoft Graph sends notifications to your webhook endpoint
4. **Processing**: The webhook endpoint processes these notifications

### Webhook Endpoints

- `POST /create-subscription` - Creates a webhook subscription
- `POST /webhook` - Receives notifications from Microsoft Graph

### Testing Webhooks Locally

For local development, you'll need to expose your local server to the internet. You can use tools like:

- **ngrok**: `ngrok http 3000`
- **localtunnel**: `npx localtunnel --port 3000`

Update your `WEBHOOK_URL` in `.env` to use the ngrok/localtunnel URL.

## API Endpoints

### Authentication
- `GET /login` - Initiates OAuth flow
- `GET /callback` - Handles OAuth callback
- `POST /logout` - Logs out user

### Email Operations
- `GET /fetch-emails` - Fetches recent emails
- `POST /create-subscription` - Creates webhook subscription
- `POST /webhook` - Webhook endpoint for notifications

### Utility
- `GET /health` - Health check endpoint

## Project Structure

```
Node-Webhook/
├── server.js              # Main Express server
├── package.json           # Dependencies and scripts
├── vercel.json           # Vercel deployment config
├── env.example           # Environment variables template
├── public/
│   └── index.html        # Frontend interface
└── README.md             # This file
```

## Security Considerations

1. **Environment Variables**: Never commit `.env` files to version control
2. **Session Security**: Use strong, random session secrets
3. **HTTPS**: Always use HTTPS in production
4. **Webhook Validation**: The app validates webhook requests using client state
5. **Token Storage**: Access tokens are stored in server-side sessions

## Troubleshooting

### Common Issues

1. **"Invalid client" error**: Check your CLIENT_ID in the environment variables
2. **"Invalid redirect URI"**: Ensure the redirect URI in Azure matches exactly
3. **"Insufficient privileges"**: Make sure you've granted the Mail.Read permission
4. **Webhook not working**: Verify the webhook URL is accessible from the internet

### Debug Mode

Set `NODE_ENV=development` to enable detailed error logging.

## License

MIT License - feel free to use this project for your own applications.

## Support

For issues and questions:
1. Check the troubleshooting section above
2. Review Microsoft Graph API documentation
3. Check Vercel deployment logs
