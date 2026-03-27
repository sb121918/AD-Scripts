const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for Computers tab
const predefinedQueries = {
  'all': 'Get-AllComputers.ps1',
  'by-guid': 'Get-ComputerByGuid.ps1',
  'by-name': 'Get-ComputerByName.ps1',
  'by-sid': 'Get-ComputerBySid.ps1',
  'created-30-days': 'Get-ComputersCreatedLastXDays.ps1',
  'created-x-days': 'Get-ComputersCreatedLastXDays.ps1',
  'deleted-30-days': 'Get-ComputersDeletedLastXDays.ps1',
  'disabled-30-days': 'Get-ComputersDisabledLastXDays.ps1',
  'modified-30-days': 'Get-ComputersModifiedLastXDays.ps1',
  'os-server-2022': 'Get-ComputersByOS.ps1',
  'os-server-2019': 'Get-ComputersByOS.ps1',
  'os-server-2016': 'Get-ComputersByOS.ps1',
  'os-windows-11': 'Get-ComputersByOS.ps1',
  'os-windows-10': 'Get-ComputersByOS.ps1',
  'never-logged-on': 'Get-ComputersNeverLoggedOn.ps1',
  'not-logged-60-days': 'Get-ComputersNotLoggedOnXDays.ps1',
  'enabled': 'Get-EnabledComputers.ps1',
  'disabled': 'Get-DisabledComputers.ps1',
  // New forest-wide queries
  'enabled-forest': 'Get-EnabledComputersForest.ps1',
  'disabled-forest': 'Get-DisabledComputersForest.ps1',
  // New per-domain queries
  'all-by-domain': 'Get-ComputersByDomain.ps1',
  'os-by-domain': 'Get-ComputersByOSDomain.ps1'
};

const osVersions = {
  'os-server-2022': 'Windows Server 2022',
  'os-server-2019': 'Windows Server 2019',
  'os-server-2016': 'Windows Server 2016',
  'os-windows-11': 'Windows 11',
  'os-windows-10': 'Windows 10'
};

// Queries that require domain parameter
const domainRequiredQueries = ['all-by-domain', 'os-by-domain'];

// GET /api/computers - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, days, value, domain, os } = req.query;
    const queryType = query || 'all';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: `Invalid query type: ${queryType}` });
    }

    const args = {};
    
    // Handle days parameter
    if (queryType.includes('days')) {
      args.Days = days || 30;
    }
    if (value) {
      args.SearchValue = value;
    }
    if (queryType.startsWith('os-') && !queryType.includes('domain')) {
      args.OperatingSystem = osVersions[queryType];
    }
    
    // Handle domain-specific queries (required for these types)
    if (domainRequiredQueries.includes(queryType)) {
      if (!domain) {
        return res.status(400).json({ error: 'Domain parameter is required for this query' });
      }
      args.TargetDomain = domain;
    } else if (domain) {
      // Pass domain as optional target for all other queries (from global selector)
      args.TargetDomain = domain;
    }
    
    // Handle OS parameter for os-by-domain query
    if (queryType === 'os-by-domain' && os) {
      args.OperatingSystem = os;
    }

    const result = await runPowerShell(scriptName, args, 'computers');
    res.json(result);
  } catch (error) {
    console.error('Computer query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/computers/search - Search computers
router.get('/search', async (req, res) => {
  try {
    const { q } = req.query;
    if (!q) {
      return res.status(400).json({ error: 'Search query required' });
    }

    const result = await runPowerShell('Search-Computer.ps1', {
      SearchValue: q
    }, 'computers');
    res.json(result);
  } catch (error) {
    console.error('Computer search error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/computers/:name/details - Get computer details
router.get('/:name/details', async (req, res) => {
  try {
    const { name } = req.params;
    const result = await runPowerShell('Get-ComputerDetails.ps1', {
      ComputerName: name
    }, 'computers');
    res.json(result);
  } catch (error) {
    console.error('Computer details error:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST /api/computers/export - Export computers
router.post('/export', async (req, res) => {
  try {
    const { filters, format = 'json' } = req.body;
    const result = await runPowerShell('Export-Computers.ps1', {
      ...filters,
      Format: format
    }, 'computers');
    res.json(result);
  } catch (error) {
    console.error('Computer export error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;
