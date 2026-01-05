// ═══════════════════════════════════════════════════════════════════════════
// RESPONSE BUILDER - Build formatted Discord responses
// ═══════════════════════════════════════════════════════════════════════════

const { EmbedBuilder, AttachmentBuilder } = require('discord.js');

class ResponseBuilder {
  // ─────────────────────────────────────────────────────────────────────────
  // COLORS
  // ─────────────────────────────────────────────────────────────────────────
  
  static COLORS = {
    PRIMARY: 0x2B579A,
    SUCCESS: 0x00B050,
    WARNING: 0xFFC000,
    ERROR: 0xFF0000,
    INFO: 0x00D2FF,
  };

  // ─────────────────────────────────────────────────────────────────────────
  // ANALYSIS RESULT EMBED
  // ─────────────────────────────────────────────────────────────────────────

  static buildAnalysisEmbed(result) {
    const { summary, analysis, stages } = result;
    
    const embed = new EmbedBuilder()
      .setColor(this.COLORS.SUCCESS)
      .setTitle('🧠 INTELLIGENT ANALYSIS COMPLETE')
      .setDescription(`File processed successfully in ${result.totalTimeFormatted}`)
      .setTimestamp();

    // Quality Score
    const qualityBefore = summary.qualityBefore;
    const qualityAfter = summary.qualityAfter;
    const improvement = qualityAfter - qualityBefore;

    embed.addFields(
      {
        name: '📊 Data Quality Score',
        value: `\`\`\`
Before: ${this._getProgressBar(qualityBefore)} ${qualityBefore}%
After:  ${this._getProgressBar(qualityAfter)} ${qualityAfter}%
        (+${improvement}% improvement)
\`\`\``,
        inline: false,
      }
    );

    // Data Overview
    embed.addFields(
      {
        name: '📋 Data Overview',
        value: `
• **Original Rows:** ${summary.originalRows.toLocaleString()}
• **Cleaned Rows:** ${summary.cleanedRows.toLocaleString()}
• **Rows Removed:** ${summary.rowsRemoved.toLocaleString()}
        `,
        inline: true,
      }
    );

    // Issues Summary
    embed.addFields(
      {
        name: '🔍 Issues Found',
        value: `
🟢 **Auto-Fixed:** ${summary.issuesFixed}
🟡 **Needs Review:** ${summary.issuesNeedReview}
        `,
        inline: true,
      }
    );

    return embed;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // CHANGES EMBED
  // ─────────────────────────────────────────────────────────────────────────

  static buildChangesEmbed(changes) {
    const embed = new EmbedBuilder()
      .setColor(this.COLORS.INFO)
      .setTitle('✅ Changes Applied');

    const summaryChanges = changes.filter(c => c.type === 'SUMMARY');
    
    if (summaryChanges.length === 0) {
      embed.setDescription('No changes were necessary.');
      return embed;
    }

    let description = '';
    summaryChanges.forEach(change => {
      description += `✓ **${change.operation}:** ${change.count} changes\n`;
    });

    embed.setDescription(description);
    return embed;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // ERROR EMBED
  // ─────────────────────────────────────────────────────────────────────────

  static buildErrorEmbed(error, title = 'Error') {
    return new EmbedBuilder()
      .setColor(this.COLORS.ERROR)
      .setTitle(`❌ ${title}`)
      .setDescription(error.message || String(error))
      .setTimestamp();
  }

  // ─────────────────────────────────────────────────────────────────────────
  // PROCESSING EMBED
  // ─────────────────────────────────────────────────────────────────────────

  static buildProcessingEmbed(stage = 'Processing') {
    return new EmbedBuilder()
      .setColor(this.COLORS.PRIMARY)
      .setTitle('⏳ Processing...')
      .setDescription(`\`\`\`
${stage}...
\`\`\``)
      .setTimestamp();
  }

  // ─────────────────────────────────────────────────────────────────────────
  // STATS EMBED
  // ─────────────────────────────────────────────────────────────────────────

  static buildStatsEmbed(stats, columnStats) {
    const embed = new EmbedBuilder()
      .setColor(this.COLORS.PRIMARY)
      .setTitle('📊 Data Statistics')
      .setTimestamp();

    embed.addFields(
      { name: 'Total Rows', value: stats.totalRows?.toLocaleString() || '0', inline: true },
      { name: 'Total Columns', value: stats.totalColumns?.toLocaleString() || '0', inline: true },
      { name: 'Empty Cells', value: stats.emptyCells?.toLocaleString() || '0', inline: true },
    );

    // Add column stats (first 5)
    if (columnStats) {
      const columns = Object.keys(columnStats).slice(0, 5);
      
      columns.forEach(col => {
        const stat = columnStats[col];
        let value = `Non-empty: ${stat.nonEmptyCount}\nUnique: ${stat.uniqueCount}`;
        
        if (stat.numeric) {
          value += `\nSum: ${stat.numeric.sum?.toLocaleString()}\nAvg: ${stat.numeric.average?.toFixed(2)}`;
        }
        
        embed.addFields({ name: `📈 ${col}`, value, inline: true });
      });
    }

    return embed;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // TEMPLATE LIST EMBED
  // ─────────────────────────────────────────────────────────────────────────

  static buildTemplateListEmbed(templates) {
    const embed = new EmbedBuilder()
      .setColor(this.COLORS.PRIMARY)
      .setTitle('📋 Available Templates')
      .setDescription('Use `/template <name>` to generate a template');

    templates.forEach(t => {
      embed.addFields({
        name: `📄 ${t.name}`,
        value: t.description,
        inline: true,
      });
    });

    return embed;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELP EMBED
  // ─────────────────────────────────────────────────────────────────────────

  static buildHelpEmbed() {
    const embed = new EmbedBuilder()
      .setColor(this.COLORS.PRIMARY)
      .setTitle('🤖 Excel Intelligent Bot - Help')
      .setDescription('Your AI-powered Excel assistant for data processing')
      .setTimestamp();

    embed.addFields(
      {
        name: '📊 Analysis & Cleaning',
        value: `
\`/analyze\` - Full intelligent analysis + auto-fix
\`/clean\` - Quick data cleaning
\`/stats\` - View data statistics
        `,
        inline: false,
      },
      {
        name: '🔄 Conversion',
        value: `
\`/convert\` - Convert to CSV, JSON, HTML, etc
        `,
        inline: false,
      },
      {
        name: '🆕 Creation',
        value: `
\`/create\` - Create Excel from text
\`/extract\` - Extract table from image (OCR)
\`/format\` - Apply custom formatting
\`/template\` - Generate from template
        `,
        inline: false,
      },
    );

    embed.setFooter({ text: 'Upload an Excel/CSV file with any command' });

    return embed;
  }

  // ─────────────────────────────────────────────────────────────────────────
  // HELPER METHODS
  // ─────────────────────────────────────────────────────────────────────────

  static _getProgressBar(percentage, length = 20) {
    const filled = Math.round((percentage / 100) * length);
    const empty = length - filled;
    return '█'.repeat(filled) + '░'.repeat(empty);
  }

  /**
   * Create file attachment from buffer
   */
  static createAttachment(buffer, fileName) {
    return new AttachmentBuilder(buffer, { name: fileName });
  }
}

module.exports = ResponseBuilder;
