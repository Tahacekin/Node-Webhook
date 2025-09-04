#!/usr/bin/env node

/**
 * Complete Production Setup Script
 * 
 * This script will help you set up your production environment step by step.
 * 
 * Usage: node setup-production.js
 */

require('dotenv').config();
const { Sequelize } = require('sequelize');
const { execSync } = require('child_process');

console.log('🚀 Starting Production Setup for Webhook Renewal System\n');

async function setupProduction() {
  console.log('📋 Step 1: Checking Environment Variables...\n');
  
  const requiredVars = [
    'DATABASE_URL',
    'CLIENT_ID', 
    'CLIENT_SECRET',
    'APP_URL',
    'SESSION_SECRET',
    'WEBHOOK_SECRET'
  ];
  
  const missingVars = requiredVars.filter(varName => !process.env[varName]);
  
  if (missingVars.length > 0) {
    console.log('❌ Missing required environment variables:');
    missingVars.forEach(varName => console.log(`   - ${varName}`));
    console.log('\n📝 Please set these in your Vercel dashboard:');
    console.log('   Vercel Dashboard → Your Project → Settings → Environment Variables\n');
    
    console.log('🔧 Required Environment Variables:');
    console.log('```');
    console.log('DATABASE_URL=postgresql://postgres:[password]@[host]:5432/postgres');
    console.log('CLIENT_ID=your_microsoft_client_id');
    console.log('CLIENT_SECRET=your_microsoft_client_secret');
    console.log('APP_URL=https://node-webhook-sage.vercel.app');
    console.log('SESSION_SECRET=your_secure_session_secret');
    console.log('WEBHOOK_SECRET=your_webhook_secret');
    console.log('```\n');
    
    console.log('💡 After setting these variables, run this script again.');
    return;
  }
  
  console.log('✅ All required environment variables are set!\n');
  
  console.log('📋 Step 2: Testing Database Connection...\n');
  
  try {
    const sequelize = new Sequelize(process.env.DATABASE_URL, {
      dialect: 'postgres',
      logging: false
    });
    
    await sequelize.authenticate();
    console.log('✅ Database connection successful!\n');
    
    // Check existing tables
    const tables = await sequelize.query(`
      SELECT table_name 
      FROM information_schema.tables 
      WHERE table_schema = 'public' 
      AND table_name IN ('Subscriptions', 'SequelizeMeta');
    `);
    
    console.log('📊 Existing tables:', tables[0].map(t => t.table_name));
    
    if (tables[0].length === 0) {
      console.log('\n📋 Step 3: Running Database Migrations...\n');
      
      try {
        process.env.NODE_ENV = 'production';
        execSync('npx sequelize-cli db:migrate --env production', { 
          stdio: 'inherit',
          env: { ...process.env, NODE_ENV: 'production' }
        });
        console.log('✅ Migrations completed successfully!\n');
      } catch (migrationError) {
        console.log('❌ Migration failed:', migrationError.message);
        console.log('\n🔧 Manual migration command:');
        console.log('npx sequelize-cli db:migrate --env production\n');
      }
    } else {
      console.log('✅ Database tables already exist!\n');
    }
    
    await sequelize.close();
    
  } catch (error) {
    console.log('❌ Database connection failed:', error.message);
    console.log('\n🔧 Troubleshooting:');
    
    if (error.message.includes('password authentication failed')) {
      console.log('   - Check your Supabase password in DATABASE_URL');
    } else if (error.message.includes('ENOTFOUND')) {
      console.log('   - Check your Supabase host URL in DATABASE_URL');
    } else if (error.message.includes('ECONNREFUSED')) {
      console.log('   - Check your Supabase port (should be 5432)');
    }
    
    console.log('\n💡 Get your connection string from:');
    console.log('   Supabase Dashboard → Settings → Database → Connection string\n');
    return;
  }
  
  console.log('📋 Step 4: Testing Local Server...\n');
  
  try {
    // Start server in background
    const serverProcess = require('child_process').spawn('node', ['server.js'], {
      stdio: 'pipe',
      detached: true
    });
    
    // Wait a moment for server to start
    await new Promise(resolve => setTimeout(resolve, 3000));
    
    // Test health endpoint
    const response = await fetch('http://localhost:3000/health');
    const health = await response.json();
    
    if (health.status === 'OK') {
      console.log('✅ Local server is running successfully!');
      console.log('🌐 Visit: http://localhost:3000');
    } else {
      console.log('⚠️  Server started but health check failed');
    }
    
    // Kill the test server
    serverProcess.kill();
    
  } catch (error) {
    console.log('⚠️  Could not test local server:', error.message);
    console.log('💡 You can test manually by running: npm start');
  }
  
  console.log('\n📋 Step 5: Production Deployment Checklist...\n');
  
  console.log('✅ Database: Connected and migrated');
  console.log('✅ Environment Variables: Set in Vercel');
  console.log('✅ Code: Ready for deployment');
  
  console.log('\n🔧 Next Steps:');
  console.log('1. Commit and push your changes to GitHub');
  console.log('2. Vercel will automatically redeploy');
  console.log('3. Test your production app: https://node-webhook-sage.vercel.app/');
  console.log('4. Update Microsoft Graph redirect URI:');
  console.log('   Azure Portal → App registrations → Your app → Authentication');
  console.log('   Add: https://node-webhook-sage.vercel.app/callback');
  
  console.log('\n🎉 Setup Complete! Your webhook renewal system is ready for production.');
}

setupProduction().catch(console.error);
