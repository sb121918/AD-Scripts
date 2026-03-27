const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for GPOs tab
const predefinedQueries = {
  'all': 'Get-AllGPOs.ps1',
  'by-name': 'Get-GPOByName.ps1',
  'created-30-days': 'Get-GPOsCreatedLastXDays.ps1',
  'modified-30-days': 'Get-GPOsModifiedLastXDays.ps1',
  'enabled': 'Get-EnabledGPOs.ps1',
  'disabled': 'Get-DisabledGPOs.ps1',
  'disabled-gpos': 'Get-DisabledGPOs.ps1',
  'disabled-by-domain': 'Get-DisabledGPOsByDomain.ps1',
  'linked-domain': 'Get-DomainLinkedGPOs.ps1',
  'linked-ou': 'Get-OULinkedGPOs.ps1',
  'unlinked': 'Get-UnlinkedGPOs.ps1',
  'unlinked-by-domain': 'Get-UnlinkedGPOsByDomain.ps1',
  'password-policies': 'Get-AllPasswordPolicies.ps1',
  'gpos-by-domain': 'Get-GPOsByDomain.ps1'
};

// Queries that require domain parameter
const domainRequiredQueries = ['gpos-by-domain'];
// Queries that optionally accept domain parameter
const domainOptionalQueries = ['unlinked-by-domain', 'disabled-by-domain'];

// GET /api/gpos - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, value, domain } = req.query;
    const queryType = query || 'all';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: `Invalid query type: ${queryType}` });
    }

    const args = {};
    if (value) {
      args.SearchValue = value;
    }
    
    // Handle domain-specific queries (required)
    if (domainRequiredQueries.includes(queryType)) {
      if (!domain) {
        return res.status(400).json({ error: 'Domain parameter is required for this query' });
      }
      args.TargetDomain = domain;
    }
    
    // Handle domain-specific queries (optional) - if domain provided, use domain-specific script
    if (domainOptionalQueries.includes(queryType)) {
      if (domain && domain.trim()) {
        args.TargetDomain = domain.trim();
      }
    }

    // Pass domain as optional target for all other queries (from global selector)
    if (!args.TargetDomain && domain && domain.trim()) {
      args.TargetDomain = domain.trim();
    }

    // Password policies script is in domain-controllers folder
    const scriptFolder = queryType === 'password-policies' ? 'domain-controllers' : 'gpos';
    const result = await runPowerShell(scriptName, args, scriptFolder);
    res.json(result);
  } catch (error) {
    console.error('GPO query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/gpos/search - Search GPOs
router.get('/search', async (req, res) => {
  try {
    const { q } = req.query;
    if (!q) {
      return res.status(400).json({ error: 'Search query required' });
    }

    const result = await runPowerShell('Search-GPO.ps1', {
      SearchValue: q
    }, 'gpos');
    res.json(result);
  } catch (error) {
    console.error('GPO search error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/gpos/:name/links - Get GPO links
router.get('/:name/links', async (req, res) => {
  try {
    const { name } = req.params;
    const result = await runPowerShell('Get-GPOLinks.ps1', {
      GPOName: name
    }, 'gpos');
    res.json(result);
  } catch (error) {
    console.error('GPO links error:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST /api/gpos/export - Export GPOs
router.post('/export', async (req, res) => {
  try {
    const { filters, format = 'json' } = req.body;
    const result = await runPowerShell('Export-GPOs.ps1', {
      ...filters,
      Format: format
    }, 'gpos');
    res.json(result);
  } catch (error) {
    console.error('GPO export error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;
