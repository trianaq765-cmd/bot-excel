// ═══════════════════════════════════════════════════════════════════════════
// COMMAND: /template - Generate from template
// ═══════════════════════════════════════════════════════════════════════════

const { SlashCommandBuilder, EmbedBuilder } = require('discord.js');
const engine = require('../../engine');
const ResponseBuilder = require('../handlers/responseBuilder');

module.exports = {
  cooldown: 5,
  
  data: new SlashCommandBuilder()
    .setName('template')
    .setDescription('📋 Generate Excel from template')
    .addStringOption(option =>
      option
        .setName('type')
        .setDescription('Template type')
        .setRequired(true)
        .addChoices(
          { name: '🧾 Invoice', value: 'invoice' },
          { name: '💵 Payroll / Slip Gaji', value: 'payroll' },
          { name: '📦 Inventory', value: 'inventory' },
          { name: '📈 Sales Report', value: 'sales-report' },
          { name: '💰 Budget', value: 'budget' },
          { name: '📅 Attendance', value: 'attendance' },
          { name: '🧾 Expense', value: 'expense' },
        )
    )
    .addStringOption(option =>
      option
        .setName('company')
        .setDescription('Company name (for invoice/payroll)')
        .setRequired(false)
    ),

  async execute(interaction) {
    const templateType = interaction.options.getString('type');
    const companyName = interaction.options.getString('company');

    await interaction.deferReply();

    try {
      const result = await engine.generateTemplate(templateType, {
        companyName,
      });

      if (!result.success) {
        throw new Error(result.error || 'Template generation failed');
      }

      const templates = engine.listTemplates();
      const templateInfo = templates.find(t => t.name === templateType);

      const embed = new EmbedBuilder()
        .setColor(0x00B050)
        .setTitle(`📋 Template Generated: ${templateType}`)
        .setDescription(templateInfo?.description || 'Template ready to use!')
        .addFields(
          { name: 'Template Type', value: templateType, inline: true },
          { name: 'Status', value: '✅ Ready', inline: true },
        )
        .setFooter({ text: 'Fill in the yellow cells with your data' })
        .setTimestamp();

      const outputFile = ResponseBuilder.createAttachment(
        result.buffer,
        `${templateType}_template.xlsx`
      );

      await interaction.editReply({
        embeds: [embed],
        files: [outputFile],
      });

    } catch (error) {
      console.error('Template command error:', error);
      await interaction.editReply({
        embeds: [ResponseBuilder.buildErrorEmbed(error, 'Template Generation Failed')],
      });
    }
  },
};
