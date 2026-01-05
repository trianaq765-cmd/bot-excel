// ═══════════════════════════════════════════════════════════════════════════
// COMMAND: /analyze - Full intelligent analysis
// ═══════════════════════════════════════════════════════════════════════════

const { SlashCommandBuilder } = require('discord.js');
const engine = require('../../engine');
const fileHandler = require('../handlers/fileHandler');
const ResponseBuilder = require('../handlers/responseBuilder');

module.exports = {
  cooldown: 10,
  
  data: new SlashCommandBuilder()
    .setName('analyze')
    .setDescription('🧠 Full intelligent analysis with auto-fix')
    .addAttachmentOption(option =>
      option
        .setName('file')
        .setDescription('Excel or CSV file to analyze')
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName('mode')
        .setDescription('Analysis mode')
        .setRequired(false)
        .addChoices(
          { name: '🔄 Auto (Recommended)', value: 'auto' },
          { name: '💰 Finance Focus', value: 'finance' },
          { name: '📈 Sales Focus', value: 'sales' },
          { name: '🔍 Strict (Maximum)', value: 'strict' },
        )
    ),

  async execute(interaction) {
    const attachment = interaction.options.getAttachment('file');
    const mode = interaction.options.getString('mode') || 'auto';

    // Send processing message
    await interaction.deferReply();

    try {
      // Download file
      await interaction.editReply({
        embeds: [ResponseBuilder.buildProcessingEmbed('Downloading file')],
      });

      const fileData = await fileHandler.downloadAttachment(attachment);

      // Process file
      await interaction.editReply({
        embeds: [ResponseBuilder.buildProcessingEmbed('Analyzing data')],
      });

      const result = await engine.process(fileData.buffer, {
        fileName: fileData.fileName,
        mode: mode,
      });

      if (!result.success) {
        throw new Error(result.error);
      }

      // Build response
      const embed = ResponseBuilder.buildAnalysisEmbed(result);
      const changesEmbed = ResponseBuilder.buildChangesEmbed(result.changes);
      
      // Create output file
      const outputFile = ResponseBuilder.createAttachment(
        result.output.buffer,
        result.output.filename
      );

      await interaction.editReply({
        embeds: [embed, changesEmbed],
        files: [outputFile],
      });

    } catch (error) {
      console.error('Analyze command error:', error);
      await interaction.editReply({
        embeds: [ResponseBuilder.buildErrorEmbed(error, 'Analysis Failed')],
      });
    }
  },
};
