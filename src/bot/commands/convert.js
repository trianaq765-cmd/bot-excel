// ═══════════════════════════════════════════════════════════════════════════
// COMMAND: /convert - Format conversion
// ═══════════════════════════════════════════════════════════════════════════

const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const engine = require('../../engine');
const fileHandler = require('../handlers/fileHandler');
const ResponseBuilder = require('../handlers/responseBuilder');

module.exports = {
  cooldown: 5,
  
  data: new SlashCommandBuilder()
    .setName('convert')
    .setDescription('🔄 Convert Excel to other formats')
    .addAttachmentOption(option =>
      option
        .setName('file')
        .setDescription('Excel or CSV file to convert')
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName('format')
        .setDescription('Target format')
        .setRequired(true)
        .addChoices(
          { name: '📄 CSV', value: 'csv' },
          { name: '📋 JSON', value: 'json' },
          { name: '🌐 HTML', value: 'html' },
          { name: '📝 Markdown', value: 'markdown' },
          { name: '🗃️ SQL', value: 'sql' },
          { name: '📰 XML', value: 'xml' },
        )
    ),

  async execute(interaction) {
    const attachment = interaction.options.getAttachment('file');
    const format = interaction.options.getString('format');

    await interaction.deferReply();

    try {
      const fileData = await fileHandler.downloadAttachment(attachment);

      const result = await engine.convert(fileData.buffer, format, {
        fileName: fileData.fileName,
        pretty: true,
        tableName: 'data',
      });

      if (!result.success) {
        throw new Error(result.error || 'Conversion failed');
      }

      // Determine file extension and content type
      const extensions = {
        csv: '.csv',
        json: '.json',
        html: '.html',
        markdown: '.md',
        sql: '.sql',
        xml: '.xml',
      };

      const baseName = fileData.fileName.replace(/\.[^/.]+$/, '');
      const outputFileName = `${baseName}${extensions[format]}`;

      const embed = new EmbedBuilder()
        .setColor(0x2B579A)
        .setTitle('🔄 Conversion Complete')
        .setDescription(`Converted to ${format.toUpperCase()} format`)
        .addFields(
          { name: 'Original', value: fileData.fileName, inline: true },
          { name: 'Output', value: outputFileName, inline: true },
        )
        .setTimestamp();

      const outputBuffer = Buffer.from(result.output, 'utf-8');
      const outputFile = ResponseBuilder.createAttachment(outputBuffer, outputFileName);

      await interaction.editReply({
        embeds: [embed],
        files: [outputFile],
      });

    } catch (error) {
      console.error('Convert command error:', error);
      await interaction.editReply({
        embeds: [ResponseBuilder.buildErrorEmbed(error, 'Conversion Failed')],
      });
    }
  },
};
