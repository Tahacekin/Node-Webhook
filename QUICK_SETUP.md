# 🚀 Quick Production Setup Guide

## What I Can Do For You ✅
- ✅ Fixed port conflicts
- ✅ Created setup scripts
- ✅ Prepared database migrations
- ✅ Configured dynamic URLs

## What You Need To Do 🔧

### Step 1: Get Supabase Connection String (2 minutes)
1. Go to [supabase.com/dashboard](https://supabase.com/dashboard)
2. Select your project
3. Go to **Settings** → **Database**
4. Copy the **URI** connection string (looks like: `postgresql://postgres:[password]@[host]:5432/postgres`)

### Step 2: Set Vercel Environment Variables (3 minutes)
1. Go to [vercel.com/dashboard](https://vercel.com/dashboard)
2. Select your `node-webhook-sage` project
3. Go to **Settings** → **Environment Variables**
4. Add these variables:

```
DATABASE_URL=postgresql://postgres:[your-password]@[your-host]:5432/postgres
CLIENT_ID=your_microsoft_client_id
CLIENT_SECRET=your_microsoft_client_secret
APP_URL=https://node-webhook-sage.vercel.app
SESSION_SECRET=your_secure_session_secret
WEBHOOK_SECRET=your_webhook_secret
```

### Step 3: Run Setup Script (1 minute)
```bash
# Set your DATABASE_URL first
export DATABASE_URL="postgresql://postgres:[password]@[host]:5432/postgres"

# Run the complete setup
node setup-production.js
```

### Step 4: Update Microsoft Graph (2 minutes)
1. Go to [Azure Portal](https://portal.azure.com)
2. **Azure Active Directory** → **App registrations**
3. Select your app → **Authentication**
4. Add redirect URI: `https://node-webhook-sage.vercel.app/callback`
5. Save changes

### Step 5: Deploy (Automatic)
- Push your changes to GitHub
- Vercel will automatically redeploy
- Test: [https://node-webhook-sage.vercel.app/](https://node-webhook-sage.vercel.app/)

## 🎯 Total Time: ~8 minutes

## 🧪 Testing Your Setup

Once everything is configured:

1. **Test Local**: `npm start` → Visit `http://localhost:3000`
2. **Test Production**: Visit [https://node-webhook-sage.vercel.app/](https://node-webhook-sage.vercel.app/)
3. **Test Login**: Click "Login with Microsoft"
4. **Test Webhook**: Create a webhook subscription
5. **Test Database**: Check Supabase dashboard for stored subscriptions

## 🆘 Need Help?

If you get stuck on any step, just let me know and I'll help you troubleshoot!
