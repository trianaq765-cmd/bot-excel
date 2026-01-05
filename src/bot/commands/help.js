// ═══════════════════════════════════════════════════════════════════════════
// COMMAND: /help - Show help information
// ═══════════════════════════════════════════════════════════════════════════

const { SlashCommandBuilder } = require('discord.js');
const ResponseBuilder = require('../handlers/responseBuilder');

module.exports = {
  cooldown: 3,
  
  data: new SlashCommandBuilder()
    .setName('help')
    .setDescription('❓ Show help information'),

  async execute(interaction) {
    const embed = ResponseBuilder.buildHelpEmbed();
    await interaction.reply({ embeds: [embed] });
  },
};
