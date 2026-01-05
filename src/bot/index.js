// ═══════════════════════════════════════════════════════════════════════════
// DISCORD BOT - Main Entry Point
// ═══════════════════════════════════════════════════════════════════════════

const { 
  Client, 
  GatewayIntentBits, 
  Collection, 
  REST, 
  Routes,
  Events,
  ActivityType,
} = require('discord.js');
const fs = require('fs');
const path = require('path');

class ExcelBot {
  constructor() {
    this.client = new Client({
      intents: [
        GatewayIntentBits.Guilds,
        GatewayIntentBits.GuildMessages,
        GatewayIntentBits.MessageContent,
      ],
    });

    this.commands = new Collection();
    this.cooldowns = new Collection();
  }

  // ─────────────────────────────────────────────────────────────────────────
  // INITIALIZATION
  // ─────────────────────────────────────────────────────────────────────────

  async initialize() {
    console.log('🤖 Initializing Excel Intelligent Bot...');

    // Load commands
    await this.loadCommands();

    // Setup event handlers
    this.setupEvents();

    // Login
    await this.login();
  }

  async loadCommands() {
    const commandsPath = path.join(__dirname, 'commands');
    
    if (!fs.existsSync(commandsPath)) {
      console.log('⚠️ Commands folder not found, creating...');
      fs.mkdirSync(commandsPath, { recursive: true });
      return;
    }

    const commandFiles = fs.readdirSync(commandsPath).filter(f => f.endsWith('.js'));

    for (const file of commandFiles) {
      const filePath = path.join(commandsPath, file);
      const command = require(filePath);

      if ('data' in command && 'execute' in command) {
        this.commands.set(command.data.name, command);
        console.log(`  ✅ Loaded command: /${command.data.name}`);
      } else {
        console.log(`  ⚠️ Skipped ${file}: missing data or execute`);
      }
    }

    console.log(`📦 Loaded ${this.commands.size} commands`);
  }

  setupEvents() {
    // Ready event
    this.client.once(Events.ClientReady, (client) => {
      console.log(`\n✅ Bot is online as ${client.user.tag}`);
      console.log(`📊 Serving ${client.guilds.cache.size} servers`);
      
      // Set activity
      client.user.setActivity('Excel files | /help', { 
        type: ActivityType.Watching 
      });
    });

    // Interaction handler
    this.client.on(Events.InteractionCreate, async (interaction) => {
      await this.handleInteraction(interaction);
    });

    // Error handler
    this.client.on(Events.Error, (error) => {
      console.error('❌ Discord client error:', error);
    });
  }

  async handleInteraction(interaction) {
    if (!interaction.isChatInputCommand()) return;

    const command = this.commands.get(interaction.commandName);
    if (!command) return;

    // Cooldown check
    const cooldownAmount = (command.cooldown || 3) * 1000;
    const now = Date.now();
    const timestamps = this.cooldowns.get(command.data.name) || new Collection();
    
    if (timestamps.has(interaction.user.id)) {
      const expirationTime = timestamps.get(interaction.user.id) + cooldownAmount;
      if (now < expirationTime) {
        const timeLeft = ((expirationTime - now) / 1000).toFixed(1);
        return interaction.reply({
          content: `⏳ Please wait ${timeLeft}s before using \`/${command.data.name}\` again.`,
          ephemeral: true,
        });
      }
    }

    timestamps.set(interaction.user.id, now);
    this.cooldowns.set(command.data.name, timestamps);

    // Execute command
    try {
      await command.execute(interaction);
    } catch (error) {
      console.error(`❌ Error executing ${interaction.commandName}:`, error);
      
      const errorMessage = {
        content: '❌ An error occurred while executing this command.',
        ephemeral: true,
      };

      if (interaction.replied || interaction.deferred) {
        await interaction.followUp(errorMessage);
      } else {
        await interaction.reply(errorMessage);
      }
    }
  }

  async login() {
    const token = process.env.DISCORD_TOKEN;
    
    if (!token) {
      throw new Error('DISCORD_TOKEN not found in environment variables');
    }

    await this.client.login(token);
  }

  // ─────────────────────────────────────────────────────────────────────────
  // COMMAND REGISTRATION
  // ─────────────────────────────────────────────────────────────────────────

  async registerCommands() {
    const commands = [];
    
    this.commands.forEach(command => {
      commands.push(command.data.toJSON());
    });

    const rest = new REST().setToken(process.env.DISCORD_TOKEN);

    try {
      console.log(`🔄 Registering ${commands.length} commands...`);

      // Global commands
      await rest.put(
        Routes.applicationCommands(process.env.DISCORD_CLIENT_ID),
        { body: commands }
      );

      console.log('✅ Commands registered successfully');
    } catch (error) {
      console.error('❌ Failed to register commands:', error);
    }
  }
}

module.exports = ExcelBot;
