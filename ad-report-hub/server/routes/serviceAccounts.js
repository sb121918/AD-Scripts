const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for Service Accounts tab
const predefinedQueries = {
  'all-forest': 'Get-AllServiceAccountsForest.ps1',
  'all': 'Get-AllServiceAccounts.ps1',
  'orphaned': 'Get-OrphanedServiceAccounts.ps1',
  'orphaned-recent': 'Get-RecentOrphanedServiceAccounts.ps1',
  'inactive': 'Get-InactiveServiceAccounts.ps1',
  'pwd-never-expires': 'Get-SvcPasswordNeverExpires.ps1',
  'interactive-logon': 'Get-SvcInteractiveLogon.ps1'
};

// GET /api/service-accounts - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, domain, days = 365 } = req.query;
    const queryType = query || 'orphaned';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: 'Invalid query type' });
    }

    const args = {};
    
    // Domain is required for 'all' (per-domain) and 'inactive' queries
    if (queryType === 'all' || queryType === 'inactive') {
      if (!domain) {
        return res.status(400).json({ error: 'Domain parameter is required for this query' });
      }
      args.TargetDomain = domain;
    }
    
    // Days parameter for orphaned-recent and inactive queries
    if (queryType === 'orphaned-recent' || queryType === 'inactive') {
      args.Days = days;
    }

    // Pass domain as optional target for all other queries (from global selector)
    if (!args.TargetDomain && domain && domain.trim()) {
      args.TargetDomain = domain.trim();
    }

    const result = await runPowerShell(scriptName, args, 'service-accounts');
    res.json(result);
  } catch (error) {
    console.error('Service Accounts query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/service-accounts/by-domain - Service Accounts by Domain
router.get('/by-domain', async (req, res) => {
  try {
    const { domain } = req.query;
    if (!domain) {
      return res.status(400).json({ error: 'Domain parameter is required' });
    }
    const result = await runPowerShell('Get-AllServiceAccounts.ps1', { TargetDomain: domain }, 'service-accounts');
    res.json(result);
  } catch (error) {
    console.error('Service accounts by domain error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/service-accounts/orphaned - Orphaned Service Accounts
router.get('/orphaned', async (req, res) => {
  try {
    const result = await runPowerShell('Get-OrphanedServiceAccounts.ps1', {}, 'service-accounts');
    res.json(result);
  } catch (error) {
    console.error('Orphaned accounts error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/service-accounts/orphaned-recent - Recently Created Orphaned Accounts
router.get('/orphaned-recent', async (req, res) => {
  try {
    const { days = 365 } = req.query;
    const result = await runPowerShell('Get-RecentOrphanedServiceAccounts.ps1', { Days: days }, 'service-accounts');
    res.json(result);
  } catch (error) {
    console.error('Recent orphaned accounts error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;
