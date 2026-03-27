const express = require('express');
const router = express.Router();
const path = require('path');
const { runPowerShell } = require('../utils/powershell');

const predefinedQueries = {
  'forest-info': 'Get-ForestInfo.ps1',
  'domains': 'Get-AllDomains.ps1',
  'trusts': 'Get-ADTrusts.ps1',
  'naming-contexts': 'Get-NamingContexts.ps1',
  'fsmo-roles': 'Get-FSMORoles.ps1',
  'ad-sites': 'Get-ADSitesSummary.ps1',
  'site-links': 'Get-SiteLinks.ps1',
  'dc-replication': 'Get-DCReplicationStatus.ps1',
};

router.get('/', async (req, res) => {
  const { query, domain } = req.query;

  if (!query || !predefinedQueries[query]) {
    return res.status(400).json({ error: 'Invalid query type' });
  }

  try {
    const args = {};
    // Pass domain as TargetDomain so scripts query the selected forest
    if (domain) {
      args.TargetDomain = domain;
    }

    const result = await runPowerShell(predefinedQueries[query], args, 'topology');

    // Normalize to array
    const dataArray = Array.isArray(result.data) ? result.data : [result.data];

    res.json({ data: dataArray, duration: result.duration });
  } catch (error) {
    console.error('Topology query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/topology/forests - Return configured domains from static config
router.get('/forests', (req, res) => {
  try {
    const fs = require('fs');
    const configPath = path.join(__dirname, '..', 'config', 'domains.json');
    const raw = fs.readFileSync(configPath, 'utf8');
    const domainsConfig = JSON.parse(raw);
    res.json({ data: domainsConfig });
  } catch (error) {
    console.error('Failed to load domains config:', error);
    res.status(500).json({ error: 'Failed to load domain configuration' });
  }
});

module.exports = router;
