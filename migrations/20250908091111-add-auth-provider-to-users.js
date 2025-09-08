'use strict';

/** @type {import('sequelize-cli').Migration} */
module.exports = {
  async up (queryInterface, Sequelize) {
    await queryInterface.addColumn('Users', 'authProvider', {
      type: Sequelize.ENUM('microsoft', 'email'),
      allowNull: false,
      defaultValue: 'email'
    });
  },

  async down (queryInterface, Sequelize) {
    await queryInterface.removeColumn('Users', 'authProvider');
  }
};
