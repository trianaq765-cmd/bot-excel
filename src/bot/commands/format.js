// ═══════════════════════════════════════════════════════════════════════════
// COMMAND: /format - Apply custom formatting
// ═══════════════════════════════════════════════════════════════════════════

const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const engine = require('../../engine');
const fileHandler = require('../handlers/fileHandler');
const ResponseBuilder = require('../handlers/responseBuilder');

module.exports = {
  cooldown: 5,
  
  data: new SlashCommandBuilder()
    .setName('format')
    .setDescription('🎨 Apply custom formatting to Excel')
    .addAttachmentOption(option =>
      option
        .setName('file')
        .setDescription('Excel or CSV file to format')
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName('instructions')
        .setDescription('Formatting instructions (e.g., "header biru, kolom Harga currency")')
        .setRequired(true)
    ),

  async execute(interaction) {
    const attachment = interaction.options.getAttachment('file');
    const instructions = interaction.options.getString('instructions');

    await interaction.deferReply();

    try {
      const fileData = await fileHandler.downloadAttachment(attachment);

      const result = await engine.applyFormat(fileData.buffer, instructions, {
        fileName: fileData.fileName,
      });

      if (!result.success) {
        throw new Error(result.error || 'Formatting failed');
      }

      const embed = new EmbedBuilder()
        .setColor(0x7030A0)
        .setTitle('🎨 Custom Formatting Applied')
        .setDescription(`Applied ${result.instructionsApplied} formatting instructions`)
        .setTimestamp();

      // List applied instructions
      if (result.instructions.length > 0) {
        const instructionList = result.instructions
          .slice(0, 10)
          .map((inst, i) => `${i + 1}. ${inst.type}`)
          .join('\n');
        
        embed.addFields({
          name: 'Instructions Applied',
          value: instructionList,
          inline: false,
        });
      }

      const outputFile = ResponseBuilder.createAttachment(
        result.buffer,
        `formatted_${fileData.fileName}`
      );

      await interaction.editReply({
        embeds: [embed],
        files: [outputFile],
      });

    } catch (error) {
      console.error('Format command error:', error);
      await interaction.editReply({
        embeds: [ResponseBuilder.buildErrorEmbed(error, 'Formatting Failed')],
      });
    }
  },
};
