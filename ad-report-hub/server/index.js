const express = require('express');
const cors = require('cors');
const path = require('path');

// Import routes
const usersRoutes = require('./routes/users');
const groupsRoutes = require('./routes/groups');
const computersRoutes = require('./routes/computers');
const containersRoutes = require('./routes/containers');
const gposRoutes = require('./routes/gpos');
const printersRoutes = require('./routes/printers');
const contactsRoutes = require('./routes/contacts');
const searchRoutes = require('./routes/search');
const logsRoutes = require('./routes/logs');
const domainControllersRoutes = require('./routes/domainControllers');
const sitesSubnetsRoutes = require('./routes/sitesSubnets');
const serviceAccountsRoutes = require('./routes/serviceAccounts');
const topologyRoutes = require('./routes/topology');
const dashboardRoutes = require('./routes/dashboard');
const systemRoutes = require('./routes/system');
const complianceRoutes = require('./routes/compliance');
const securityRoutes = require('./routes/security');

// Import middleware
const activityLogger = require('./middleware/activityLogger');

// Initialize database
require('./utils/database');

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(activityLogger);

// API Routes
app.use('/api/users', usersRoutes);
app.use('/api/groups', groupsRoutes);
app.use('/api/computers', computersRoutes);
app.use('/api/containers', containersRoutes);
app.use('/api/gpos', gposRoutes);
app.use('/api/printers', printersRoutes);
app.use('/api/contacts', contactsRoutes);
app.use('/api/search', searchRoutes);
app.use('/api/activity-logs', logsRoutes);
app.use('/api/domain-controllers', domainControllersRoutes);
app.use('/api/sites-subnets', sitesSubnetsRoutes);
app.use('/api/service-accounts', serviceAccountsRoutes);
app.use('/api/topology', topologyRoutes);
app.use('/api/dashboard', dashboardRoutes);
app.use('/api/system', systemRoutes);
app.use('/api/compliance', complianceRoutes);
app.use('/api/security', securityRoutes);

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Serve static files in production
if (process.env.NODE_ENV === 'production') {
  app.use(express.static(path.join(__dirname, '../build')));
  app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, '../build', 'index.html'));
  });
}

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Server Error:', err);
  res.status(500).json({ 
    error: 'Internal Server Error', 
    message: err.message 
  });
});

app.listen(PORT, () => {
  console.log(`
╔════════════════════════════════════════════════════════════╗
║             AD Report Hub - Server Started                 ║
╠════════════════════════════════════════════════════════════╣
║  Server running on: http://localhost:${PORT}                  ║
║  API Base URL:      http://localhost:${PORT}/api              ║
║  Authentication:    Windows Integrated (current user)      ║
╚════════════════════════════════════════════════════════════╝
  `);
});

module.exports = app;
