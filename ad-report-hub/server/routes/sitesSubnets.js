const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for AD Sites & Subnets tab
const predefinedQueries = {
  'site-links': 'Get-ADSiteLinks.ps1',
  'subnets': 'Get-ADSubnets.ps1',
  'subnets-detailed': 'Get-ADSubnetsDetailed.ps1',
  'sites-no-dc': 'Get-SitesWithNoDC.ps1',
  'unassigned-subnets': 'Get-UnassignedSubnets.ps1'
};

// GET /api/sites-subnets - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, domain } = req.query;
    const queryType = query || 'subnets';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: 'Invalid query type' });
    }

    const args = {};
    if (domain) {
      args.TargetDomain = domain;
    }

    const result = await runPowerShell(scriptName, args, 'sites-subnets');
    res.json(result);
  } catch (error) {
    console.error('Sites/Subnets query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/sites-subnets/site-links - AD Site Links
router.get('/site-links', async (req, res) => {
  try {
    const result = await runPowerShell('Get-ADSiteLinks.ps1', {}, 'sites-subnets');
    res.json(result);
  } catch (error) {
    console.error('Site links error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/sites-subnets/subnets - All Subnets
router.get('/subnets', async (req, res) => {
  try {
    const result = await runPowerShell('Get-ADSubnets.ps1', {}, 'sites-subnets');
    res.json(result);
  } catch (error) {
    console.error('Subnets error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/sites-subnets/subnets-detailed - Detailed Subnets
router.get('/subnets-detailed', async (req, res) => {
  try {
    const result = await runPowerShell('Get-ADSubnetsDetailed.ps1', {}, 'sites-subnets');
    res.json(result);
  } catch (error) {
    console.error('Subnets detailed error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;
