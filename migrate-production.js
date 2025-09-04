#!/usr/bin/env node

/**
 * Production Migration Script
 * 
 * This script helps you run database migrations on your production database.
 * 
 * Usage:
 * 1. Set your DATABASE_URL environment variable
 * 2. Run: node migrate-production.js
 */

require('dotenv').config();
const { Sequelize } = require('sequelize');

async function runMigrations() {
  const databaseUrl = process.env.DATABASE_URL;
  
  if (!databaseUrl) {
    console.error('❌ DATABASE_URL environment variable is not set!');
    console.log('\nPlease set your DATABASE_URL:');
    console.log('export DATABASE_URL="postgres://username:password@host:port/database"');
    console.log('\nOr add it to your .env file:');
    console.log('DATABASE_URL=postgres://username:password@host:port/database');
    process.exit(1);
  }

  console.log('🔄 Connecting to production database...');
  
  const sequelize = new Sequelize(databaseUrl, {
    dialect: 'postgres',
    logging: console.log
  });

  try {
    // Test connection
    await sequelize.authenticate();
    console.log('✅ Database connection established successfully.');

    // Run migrations
    console.log('🔄 Running migrations...');
    const { execSync } = require('child_process');
    
    // Set NODE_ENV to production for migrations
    process.env.NODE_ENV = 'production';
    
    execSync('npx sequelize-cli db:migrate --env production', { 
      stdio: 'inherit',
      env: { ...process.env, NODE_ENV: 'production' }
    });
    
    console.log('✅ Migrations completed successfully!');
    
  } catch (error) {
    console.error('❌ Migration failed:', error.message);
    process.exit(1);
  } finally {
    await sequelize.close();
  }
}

runMigrations();
