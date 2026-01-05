// ═══════════════════════════════════════════════════════════════════════════
// MAIN SERVER - Express + Discord Bot
// ═══════════════════════════════════════════════════════════════════════════

require('dotenv').config();

const express = require('express');
const cors = require('cors');
const path = require('path');
const ExcelBot = require('./bot');
const apiRoutes = require('./web/routes/api');

// ─────────────────────────────────────────────────────────────────────────
// EXPRESS SETUP
// ─────────────────────────────────────────────────────────────────────────

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Static files
app.use(express.static(path.join(__dirname, 'web/public')));

// API routes
app.use('/api', apiRoutes);

// Serve frontend for all other routes
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'web/public/index.html'));
});

// Error handler
app.use((err, req, res, next) => {
  console.error('Server error:', err);
  res.status(500).json({
    success: false,
    error: process.env.NODE_ENV === 'production' 
      ? 'Internal server error' 
      : err.message,
  });
});

// ─────────────────────────────────────────────────────────────────────────
// START SERVICES
// ─────────────────────────────────────────────────────────────────────────

async function startServer() {
  console.log('═══════════════════════════════════════════════════════════');
  console.log('   🚀 EXCEL INTELLIGENT BOT - Starting Services');
  console.log('═══════════════════════════════════════════════════════════');

  // Start Express server
  app.listen(PORT, () => {
    console.log(`\n🌐 Web server running on port ${PORT}`);
    console.log(`   Local: http://localhost:${PORT}`);
  });

  // Start Discord bot (if token provided)
  if (process.env.DISCORD_TOKEN) {
    try {
      const bot = new ExcelBot();
      await bot.loadCommands();
      
      // Register commands if CLIENT_ID provided
      if (process.env.DISCORD_CLIENT_ID) {
        await bot.registerCommands();
      }
      
      await bot.initialize();
    } catch (error) {
      console.error('❌ Failed to start Discord bot:', error.message);
      console.log('   Web server will continue running without Discord bot');
    }
  } else {
    console.log('\n⚠️  DISCORD_TOKEN not found - Bot disabled');
    console.log('   Web interface is still available');
  }

  console.log('\n═══════════════════════════════════════════════════════════');
  console.log('   ✅ All services started successfully');
  console.log('═══════════════════════════════════════════════════════════\n');
}

// Handle shutdown
process.on('SIGINT', async () => {
  console.log('\n🛑 Shutting down...');
  process.exit(0);
});

process.on('unhandledRejection', (error) => {
  console.error('Unhandled rejection:', error);
});

// Start
startServer();
