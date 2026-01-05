// ═══════════════════════════════════════════════════════════════════════════
// COMMAND: /create - Create Excel from text
// ═══════════════════════════════════════════════════════════════════════════

const { SlashCommandBuilder, EmbedBuilder, ModalBuilder, TextInputBuilder, TextInputStyle, ActionRowBuilder } = require('discord.js');
const engine = require('../../engine');
const ResponseBuilder = require('../handlers/responseBuilder');

module.exports = {
  cooldown: 5,
  
  data: new SlashCommandBuilder()
    .setName('create')
    .setDescription('📝 Create Excel from text data')
    .addStringOption(option =>
      option
        .setName('data')
        .setDescription('Paste your data (CSV format or structured text)')
        .setRequired(false)
    )
    .addBooleanOption(option =>
      option
        .setName('addtotals')
        .setDescription('Add total row for numeric columns')
        .setRequired(false)
    ),

  async execute(interaction) {
    const textData = interaction.options.getString('data');
    const addTotals = interaction.options.getBoolean('addtotals') ?? false;

    // If no data provided, show modal for input
    if (!textData) {
      const modal = new ModalBuilder()
        .setCustomId('createExcelModal')
        .setTitle('Create Excel from Text');

      const dataInput = new TextInputBuilder()
        .setCustomId('dataInput')
        .setLabel('Paste your data')
        .setStyle(TextInputStyle.Paragraph)
        .setPlaceholder('Produk, Qty, Harga\nItem A, 10, 50000\nItem B, 5, 75000')
        .setRequired(true);

      const row = new ActionRowBuilder().addComponents(dataInput);
      modal.addComponents(row);

      await interaction.showModal(modal);
      return;
    }

    await interaction.deferReply();

    try {
      const result = await engine.createFromText(textData, { addTotals });

      if (!result.success) {
        throw new Error(result.error || 'Failed to create Excel');
      }

      const embed = new EmbedBuilder()
        .setColor(0x00B050)
        .setTitle('📊 Excel Created from Text')
        .setDescription(`Successfully parsed your text data!`)
        .addFields(
          { name: 'Format Detected', value: result.format, inline: true },
          { name: 'Columns', value: result.headers.length.toString(), inline: true },
          { name: 'Rows', value: result.rowCount.toString(), inline: true },
        )
        .setTimestamp();

      // Add column preview
      if (result.headers.length > 0) {
        embed.addFields({
          name: 'Columns',
          value: result.headers.slice(0, 10).join(', ') + (result.headers.length > 10 ? '...' : ''),
          inline: false,
        });
      }

      const outputFile = ResponseBuilder.createAttachment(
        result.buffer,
        'created_data.xlsx'
      );

      await interaction.editReply({
        embeds: [embed],
        files: [outputFile],
      });

    } catch (error) {
      console.error('Create command error:', error);
      await interaction.editReply({
        embeds: [ResponseBuilder.buildErrorEmbed(error, 'Creation Failed')],
      });
    }
  },
};
