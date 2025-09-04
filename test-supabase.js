#!/usr/bin/env node

/**
 * Supabase Connection Test Script
 * 
 * This script tests the connection to your Supabase database.
 * 
 * Usage:
 * 1. Set your DATABASE_URL environment variable
 * 2. Run: node test-supabase.js
 */

require('dotenv').config();
const { Sequelize } = require('sequelize');

async function testSupabaseConnection() {
  const databaseUrl = process.env.DATABASE_URL;
  
  if (!databaseUrl) {
    console.error('❌ DATABASE_URL environment variable is not set!');
    console.log('\nPlease set your DATABASE_URL:');
    console.log('export DATABASE_URL="postgresql://postgres:[password]@[host]:5432/postgres"');
    console.log('\nOr add it to your .env file:');
    console.log('DATABASE_URL=postgresql://postgres:[password]@[host]:5432/postgres');
    process.exit(1);
  }

  console.log('🔄 Testing Supabase connection...');
  console.log('📍 Database URL:', databaseUrl.replace(/:[^:@]+@/, ':***@')); // Hide password
  
  const sequelize = new Sequelize(databaseUrl, {
    dialect: 'postgres',
    logging: false
  });

  try {
    // Test connection
    await sequelize.authenticate();
    console.log('✅ Supabase connection established successfully!');

    // Test query
    const result = await sequelize.query('SELECT version();');
    console.log('📊 PostgreSQL version:', result[0][0].version);

    // Check if our tables exist
    const tables = await sequelize.query(`
      SELECT table_name 
      FROM information_schema.tables 
      WHERE table_schema = 'public' 
      AND table_name IN ('Subscriptions', 'SequelizeMeta');
    `);
    
    console.log('📋 Existing tables:', tables[0].map(t => t.table_name));
    
    if (tables[0].length === 0) {
      console.log('⚠️  No tables found. You need to run migrations.');
      console.log('Run: node migrate-production.js');
    } else {
      console.log('✅ Tables found! Database is ready.');
    }
    
  } catch (error) {
    console.error('❌ Connection failed:', error.message);
    
    if (error.message.includes('password authentication failed')) {
      console.log('\n💡 Check your Supabase password in the connection string.');
    } else if (error.message.includes('ENOTFOUND')) {
      console.log('\n💡 Check your Supabase host URL.');
    } else if (error.message.includes('ECONNREFUSED')) {
      console.log('\n💡 Check your Supabase port (should be 5432).');
    }
    
    process.exit(1);
  } finally {
    await sequelize.close();
  }
}

testSupabaseConnection();
