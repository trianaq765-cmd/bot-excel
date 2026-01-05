// ═══════════════════════════════════════════════════════════════════════════
// MAIN SERVER - Express + Discord Bot
// ═══════════════════════════════════════════════════════════════════════════

require('dotenv').config();

const express = require('express');
const cors = require('cors');
const path = require('path');

// ─────────────────────────────────────────────────────────────────────────
// EXPRESS SETUP
// ─────────────────────────────────────────────────────────────────────────

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Request logging
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`);
  next();
});

// Static files
app.use(express.static(path.join(__dirname, 'web/public')));

// API routes
try {
  const apiRoutes = require('./web/routes/api');
  app.use('/api', apiRoutes);
  console.log('✅ API routes loaded');
} catch (error) {
  console.error('❌ Failed to load API routes:', error.message);
  
  // Fallback API
  app.use('/api', (req, res) => {
    res.status(500).json({ 
      success: false, 
      error: 'API not available: ' + error.message 
    });
  });
}

// Serve frontend for all other routes
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'web/public/index.html'));
});

// Global error handler
app.use((err, req, res, next) => {
  console.error('[Server] Error:', err);
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
  const server = app.listen(PORT, () => {
    console.log(`\n🌐 Web server running on port ${PORT}`);
  });

  // Handle server errors
  server.on('error', (error) => {
    console.error('❌ Server error:', error);
  });

  // Start Discord bot (if token provided)
  if (process.env.DISCORD_TOKEN) {
    try {
      const ExcelBot = require('./bot');
      const bot = new ExcelBot();
      
      await bot.loadCommands();
      
      if (process.env.DISCORD_CLIENT_ID) {
        await bot.registerCommands();
      }
      
      await bot.initialize();
      console.log('✅ Discord bot started');
    } catch (error) {
      console.error('❌ Discord bot error:', error.message);
      console.log('   Web server will continue running without Discord bot');
    }
  } else {
    console.log('\n⚠️  DISCORD_TOKEN not found - Bot disabled');
  }

  console.log('\n═══════════════════════════════════════════════════════════');
  console.log('   ✅ Server started successfully');
  console.log('═══════════════════════════════════════════════════════════\n');
}

// Handle shutdown
process.on('SIGINT', () => {
  console.log('\n🛑 Shutting down...');
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('\n🛑 Shutting down...');
  process.exit(0);
});

process.on('unhandledRejection', (error) => {
  console.error('Unhandled rejection:', error);
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught exception:', error);
});

// Start
startServer();
