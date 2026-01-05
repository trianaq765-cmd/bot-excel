// ═══════════════════════════════════════════════════════════════════════════
// COMMAND: /clean - Quick data cleaning
// ═══════════════════════════════════════════════════════════════════════════

const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const engine = require('../../engine');
const fileHandler = require('../handlers/fileHandler');
const ResponseBuilder = require('../handlers/responseBuilder');

module.exports = {
  cooldown: 5,
  
  data: new SlashCommandBuilder()
    .setName('clean')
    .setDescription('🧹 Quick data cleaning')
    .addAttachmentOption(option =>
      option
        .setName('file')
        .setDescription('Excel or CSV file to clean')
        .setRequired(true)
    )
    .addBooleanOption(option =>
      option
        .setName('duplicates')
        .setDescription('Remove duplicate rows (default: true)')
        .setRequired(false)
    )
    .addBooleanOption(option =>
      option
        .setName('empty')
        .setDescription('Remove empty rows (default: true)')
        .setRequired(false)
    )
    .addBooleanOption(option =>
      option
        .setName('trim')
        .setDescription('Trim whitespace (default: true)')
        .setRequired(false)
    )
    .addStringOption(option =>
      option
        .setName('textcase')
        .setDescription('Convert text case')
        .setRequired(false)
        .addChoices(
          { name: 'Title Case', value: 'title' },
          { name: 'UPPERCASE', value: 'upper' },
          { name: 'lowercase', value: 'lower' },
        )
    ),

  async execute(interaction) {
    const attachment = interaction.options.getAttachment('file');
    const removeDuplicates = interaction.options.getBoolean('duplicates') ?? true;
    const removeEmpty = interaction.options.getBoolean('empty') ?? true;
    const trimWhitespace = interaction.options.getBoolean('trim') ?? true;
    const textCase = interaction.options.getString('textcase');

    await interaction.deferReply();

    try {
      const fileData = await fileHandler.downloadAttachment(attachment);

      const result = await engine.quickClean(fileData.buffer, {
        fileName: fileData.fileName,
        removeDuplicates,
        removeEmpty,
        trimWhitespace,
        textCase,
      });

      if (!result.success) {
        throw new Error(result.error || 'Cleaning failed');
      }

      // Build response embed
      const embed = new EmbedBuilder()
        .setColor(0x00B050)
        .setTitle('🧹 Data Cleaning Complete')
        .setDescription(`File cleaned successfully!`)
        .addFields(
          { name: 'Rows Removed', value: result.stats.rowsRemoved.toString(), inline: true },
          { name: 'Cells Modified', value: result.stats.cellsModified.toString(), inline: true },
          { name: 'Total Changes', value: result.stats.totalChanges.toString(), inline: true },
        )
        .setTimestamp();

      const outputFile = ResponseBuilder.createAttachment(
        result.buffer,
        `cleaned_${fileData.fileName}`
      );

      await interaction.editReply({
        embeds: [embed],
        files: [outputFile],
      });

    } catch (error) {
      console.error('Clean command error:', error);
      await interaction.editReply({
        embeds: [ResponseBuilder.buildErrorEmbed(error, 'Cleaning Failed')],
      });
    }
  },
};
