const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for Active Directory Reports tab
const predefinedQueries = {
  'all': 'Get-AllDomainControllers.ps1',
  'health': 'Get-DCResourceHealth.ps1',
  'services': 'Get-DCServicesStatus.ps1',
  'trusts': 'Get-ADTrusts.ps1',
  'fsmo': 'Get-FSMORoleHolders.ps1',
};

// GET /api/domain-controllers/queries - Debug helper: list supported query types
router.get('/queries', (req, res) => {
  res.json({
    data: {
      supportedQueries: Object.keys(predefinedQueries),
      scripts: predefinedQueries,
      serverTime: new Date().toISOString(),
    },
  });
});

// GET /api/domain-controllers - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, domain, targetDomain } = req.query;
    const queryType = query || 'all';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: `Invalid query type. Valid types: ${Object.keys(predefinedQueries).join(', ')}` });
    }

    const args = {};
    // domain = forest anchor (selected in navbar)
    // If targetDomain is provided, prefer anchoring to it (ensures per-domain runs in the intended forest)
    if (targetDomain) {
      args.ForestDomain = targetDomain;
    } else if (domain) {
      args.ForestDomain = domain;
    }
    // targetDomain = optional per-domain filter override
    if (targetDomain) {
      args.TargetDomain = targetDomain;
    }

    const result = await runPowerShell(scriptName, args, 'domain-controllers');
    res.json(result);
  } catch (error) {
    console.error('Domain Controller query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/domain-controllers/health - DC Resource Health
router.get('/health', async (req, res) => {
  try {
    const { domain, targetDomain } = req.query;
    const args = {};
    if (domain) args.ForestDomain = domain;
    if (targetDomain) args.TargetDomain = targetDomain;
    const result = await runPowerShell('Get-DCResourceHealth.ps1', args, 'domain-controllers');
    res.json(result);
  } catch (error) {
    console.error('DC health error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/domain-controllers/services - DC Services Status
router.get('/services', async (req, res) => {
  try {
    const { domain, targetDomain } = req.query;
    const args = {};
    if (domain) args.ForestDomain = domain;
    if (targetDomain) args.TargetDomain = targetDomain;
    const result = await runPowerShell('Get-DCServicesStatus.ps1', args, 'domain-controllers');
    res.json(result);
  } catch (error) {
    console.error('DC services error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;
