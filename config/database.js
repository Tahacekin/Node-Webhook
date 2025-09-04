require('dotenv').config();

module.exports = {
  development: {
    username: process.env.DB_USER || 'webhook_user',
    password: process.env.DB_PASSWORD || 'your_secure_password',
    database: process.env.DB_NAME || 'webhook_renewal',
    host: process.env.DB_HOST || 'localhost',
    port: process.env.DB_PORT || 5432,
    dialect: 'postgres',
    logging: false
  },
  test: {
    username: process.env.DB_USER || 'webhook_user',
    password: process.env.DB_PASSWORD || 'your_secure_password',
    database: process.env.DB_NAME + '_test' || 'webhook_renewal_test',
    host: process.env.DB_HOST || 'localhost',
    port: process.env.DB_PORT || 5432,
    dialect: 'postgres',
    logging: false
  },
  production: {
    use_env_variable: 'DATABASE_URL',
    dialect: 'postgres',
    logging: false
  }
};
