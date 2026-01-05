// ═══════════════════════════════════════════════════════════════════════════
// COMMAND: /stats - View data statistics
// ═══════════════════════════════════════════════════════════════════════════

const { SlashCommandBuilder } = require('discord.js');
const engine = require('../../engine');
const fileHandler = require('../handlers/fileHandler');
const ResponseBuilder = require('../handlers/responseBuilder');

module.exports = {
  cooldown: 5,
  
  data: new SlashCommandBuilder()
    .setName('stats')
    .setDescription('📊 View data statistics')
    .addAttachmentOption(option =>
      option
        .setName('file')
        .setDescription('Excel or CSV file to analyze')
        .setRequired(true)
    ),

  async execute(interaction) {
    const attachment = interaction.options.getAttachment('file');

    await interaction.deferReply();

    try {
      const fileData = await fileHandler.downloadAttachment(attachment);
      const analysis = await engine.analyze(fileData.buffer, {
        fileName: fileData.fileName,
      });

      const embed = ResponseBuilder.buildStatsEmbed(
        analysis.summary,
        analysis.columnStats
      );

      await interaction.editReply({ embeds: [embed] });

    } catch (error) {
      console.error('Stats command error:', error);
      await interaction.editReply({
        embeds: [ResponseBuilder.buildErrorEmbed(error, 'Statistics Failed')],
      });
    }
  },
};
