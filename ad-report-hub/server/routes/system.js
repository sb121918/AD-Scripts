const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');
const os = require('os');

// GET /api/system/current-user - Get current logged in user info
router.get('/current-user', async (req, res) => {
  try {
    // Try to get detailed user info from AD
    const result = await runPowerShell('Get-CurrentUserInfo.ps1', {}, '');
    // runPowerShell normalizes everything to an array; current-user should be a single object
    const userObj = Array.isArray(result?.data) ? (result.data[0] || {}) : (result?.data || {});
    res.json({ data: userObj, duration: result?.duration });
  } catch (error) {
    // Fallback to environment variables
    const username = process.env.USERNAME || os.userInfo().username;
    const domain = process.env.USERDOMAIN || 'LOCAL';
    const computer = os.hostname();
    
    res.json({
      data: {
        username: username,
        displayName: username,
        domain: domain,
        computer: computer,
        email: null
      }
    });
  }
});

// POST /api/system/shutdown - Shutdown the application
router.post('/shutdown', (req, res) => {
  res.json({ message: 'Shutting down...' });
  
  // Give time for response to be sent
  setTimeout(() => {
    console.log('Application shutdown requested by user');
    process.exit(0);
  }, 500);
});

module.exports = router;
